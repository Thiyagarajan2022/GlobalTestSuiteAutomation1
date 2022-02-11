//USEUNIT ExcelUtils
//USEUNIT EnvParams
//USEUNIT ValidationUtils
/*Closes workspaces after job completes in maconomy*/

var Language = "";
var Maconomy_Parent = "";
var Sys_Maconomy_Parent = ""

function closeAllWorkspaces(){
 if((Language == "English") || (Language == "Chinese (Simplified)")){
  Sys.Desktop.KeyDown(0x12); //Alt
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); 
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
  }
  else if(Language =="Spanish"){ 
  Sys.Desktop.KeyDown(0x12); //Alt
  Sys.Desktop.KeyDown(0x56); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); 
  Sys.Desktop.KeyUp(0x56);
  Sys.Desktop.KeyUp(0x0D);
  }
}

function closeMaconomy(){ 
var menuBar = eval(Sys_Maconomy_Parent).SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();

 if((Language == "English") || (Language == "Chinese (Simplified)")){
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
    
    }
else if(Language =="Spanish"){ 
    Sys.Desktop.KeyDown(0x12); //Alt
  Sys.Desktop.KeyDown(0x41); //A
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyDown(0x4C);//L
  Sys.Desktop.KeyUp(0x12); 
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
  Sys.Desktop.KeyUp(0x4C);
  }
}

function SpanishcloseAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Alt
  Sys.Desktop.KeyDown(0x41); //A
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyDown(0x4C);//L
  Sys.Desktop.KeyUp(0x12); 
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
  Sys.Desktop.KeyUp(0x4C);
}

// Calculate time difference between startTime and endTime
function timeDifference(stime, etime)
{
  try{
  var seconds = (etime.getTime() - stime.getTime()) / 1000;
  var minutes = Math.floor(seconds / 60);
  var remainingSeconds = Math.floor(seconds%60);
  if(remainingSeconds<9)
    remainingSeconds = "0"+ remainingSeconds
  return minutes+"."+remainingSeconds;
  }
  catch(e)
  {
    Log.Message("");
  }
}


function VPWSearchByValue(ObjectAddrs,popupName,value,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
//    aqUtils.Delay(3000, Indicator.Text);;

    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "")
  waitForObj(code);
  code.Click();
    code.setText(value);
//    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    waitForObj(serch);

  serch.Click();    
//Sys.HighlightObject(serch);
//    if(serch.isEnabled())
//  serch.Click();
//  else{ 
//    aqUtils.Delay(3000, Indicator.Text);;
//   serch.Click(); 
//  }
//    aqUtils.Delay(5000, Indicator.Text);;
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
    var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
    waitForObj(OK);
    Sys.HighlightObject(table); 
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  waitForObj(OK);
  OK.Click();
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
  waitForObj(cancel);
  cancel.Click();

          Sys.HighlightObject(ObjectAddrs);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        waitForObj(cancel);
        cancel.Click();

      Sys.HighlightObject(ObjectAddrs);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}


function SearchByValue(ObjectAddrs,popupName,value,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, popupName);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);


    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  waitForObj(code);
  code.Click();

    code.setText(value);

    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    waitForObj(serch);

  serch.Click();
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())

    waitForObj(OK);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  waitForObj(OK);
  OK.Click();

          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
  waitForObj(cancel);
  cancel.Click();

          Sys.HighlightObject(ObjectAddrs);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        waitForObj(cancel);
        cancel.Click();

      Sys.HighlightObject(ObjectAddrs);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    
    return checkmark;
}

function TDS(ObjectAddrs,popupName,value,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, popupName);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);


    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  waitForObj(code);
  code.Click();
Log.Message(value);
    code.setText(value);

    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    waitForObj(serch);

  serch.Click();
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())

    waitForObj(OK);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  waitForObj(OK);
  OK.Click();

          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
  waitForObj(cancel);
  cancel.Click();

          Sys.HighlightObject(ObjectAddrs);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        waitForObj(cancel);
        cancel.Click();

      Sys.HighlightObject(ObjectAddrs);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    
    return checkmark;
}


function SearchByValues(ObjectAddrs,popupName,value,fieldName,alljobs){

var checkmark = false; 
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(10000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    var alljob = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
    alljob.Click();
    aqUtils.Delay(7000, Indicator.Text);; 
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(1).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
         break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}



function SearchByValuesjob(ObjectAddrs,popupName,value,fieldName){

var checkmark = false; 
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    var alljob = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
    alljob.Click();
    aqUtils.Delay(2000, Indicator.Text);; 
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(1).OleValue.toString().trim()==value){ 
      Sys.Desktop.KeyDown(0x10);     
      Sys.Desktop.KeyDown(0x09);
      Sys.Desktop.KeyUp(0x10);
      Sys.Desktop.KeyUp(0x09);
      Delay(1000);
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         break;
        ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}


function SearchByValues_Col_1(ObjectAddrs,popupName,value,fieldName){

var checkmark = false; 
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);;

    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy"); 
         break;
         
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}


function AccessLevel_Add(ObjectAddrs,popupName,value,fieldName){

var checkmark = false; 
  aqUtils.Delay(1000, popupName);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    waitForObj(code)
    code.setText(value);
//    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    waitForObj(serch)
    Sys.HighlightObject(serch);
    serch.Click();
//    aqUtils.Delay(5000, Indicator.Text);;
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
    waitForObj(table)
    waitForObj(OK)
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy"); 
         break;
         
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          waitForObj(ObjectAddrs)
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      waitForObj(ObjectAddrs)
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}


function SearchByValues_all_Col_1(ObjectAddrs,popupName,value,fieldName){

var checkmark = false; 
  aqUtils.Delay(1000, popupName);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);

    var alljob = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "All Jobs").OleValue.toString().trim());
    waitForObj(alljob);
    alljob.Click();
//    aqUtils.Delay(2000, Indicator.Text);; 
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    waitForObj(code);
    code.setText(value);
//    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
waitForObj(serch);
serch.Click();    

    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
     waitForObj(table);
     var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
     waitForObj(OK);
   
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
//         if(OK.isEnabled())
  OK.Click();

         checkmark = true;
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
         break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
          waitForObj(cancel);

          waitForObj(ObjectAddrs);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())

      waitForObj(ObjectAddrs);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}





function SearchByValues_Wiz2_Col_1(ObjectAddrs,popupName,value,fieldName){

var checkmark = false; 
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);;

    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         break;
        ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}


function SearchByValues_Wiz2_Col_2(ObjectAddrs,popupName,value,fieldName){

var checkmark = false; 
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    var grid = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    waitForObj(grid);
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(500, Indicator.Text);;
//    var alljob = eval(Sys_Maconomy_Parent).SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Jobs");
//    alljob.Click();
//    aqUtils.Delay(2000, Indicator.Text);; 
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    waitForObj(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
    waitForObj(OK);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(1).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
       waitForObj(OK);
         OK.Click();
         checkmark = true;
         break;
        ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
          waitForObj(cancel);
           cancel.Click();
//          aqUtils.Delay(1000, Indicator.Text);;
Sys.HighlightObject(ObjectAddrs)
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      waitForObj(cancel);
      cancel.Click();
//      aqUtils.Delay(1000, Indicator.Text);;
Sys.HighlightObject(ObjectAddrs)
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}








function SearchByValues_all_Col_2(ObjectAddrs,popupName,value,fieldName,all){

var checkmark = false; 
  aqUtils.Delay(1000, popupName);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
    waitForObj(code);
    code.Click();
    Sys.Desktop.KeyDown(0x10);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x10);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, popupName);;
//    aqUtils.Delay(3000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    var alljob = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", all);
    waitForObj(alljob);
    alljob.Click();
//    aqUtils.Delay(2000, Indicator.Text);; 
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    waitForObj(code);
    code.setText(value);
//    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    waitForObj(serch);
    Sys.HighlightObject(serch);
    serch.Click();
//    aqUtils.Delay(5000, Indicator.Text);;
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
    waitForObj(table);
    var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
    waitForObj(OK);
//   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(1).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
         break;
        
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          waitForObj(ObjectAddrs);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      waitForObj(ObjectAddrs);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}

function SearchByValues_Col_1_all(ObjectAddrs,popupName,value,fieldName,all){

var checkmark = false; 
  aqUtils.Delay(1000, popupName);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);

    var alljob = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", all);
    WorkspaceUtils.waitForObj(alljob);
    alljob.Click();
//    aqUtils.Delay(2000, Indicator.Text);; 
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    WorkspaceUtils.waitForObj(code);
    code.setText(value);
//    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    WorkspaceUtils.waitForObj(serch);
    Sys.HighlightObject(serch);
    serch.Click();
//    aqUtils.Delay(5000, Indicator.Text);;
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
    WorkspaceUtils.waitForObj(table);
    var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
    WorkspaceUtils.waitForObj(OK);
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
         break;
        
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          WorkspaceUtils.waitForObj(ObjectAddrs);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      WorkspaceUtils.waitForObj(ObjectAddrs);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}

// Only (SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", ""))
function SearchByValuePicker_Col_1(ObjectAddrs,popupName,value,fieldName){
var checkmark =  false;
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);;
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
          OK.Click();
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
          cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}




// Only (SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", ""))
function SearchByValuePicker_Col_2(ObjectAddrs,popupName,value,fieldName){
var checkmark =  false;
  ObjectAddrs.Click();
  aqUtils.Delay(3000, Indicator.Text);;

//  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    var grid = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    waitForObj(grid);
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    waitForObj(code);
    var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
    waitForObj(cancel);
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    code.setText(value);
//    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    waitForObj(grid);
    serch.Click();
//    aqUtils.Delay(5000, Indicator.Text);;
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
    waitForObj(OK);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(1).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
       waitForObj(OK);
          OK.Click();
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
          waitForObj(cancel);
          cancel.Click();
//          aqUtils.Delay(1000, Indicator.Text);;
          Sys.HighlightObject(ObjectAddrs);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      Sys.HighlightObject(ObjectAddrs);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}




function SearchByValuePicker(ObjectAddrs,popupName,value){ 
var checkmark =  false;
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);;
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
          OK.Click();
          checkmark = true;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
          cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
    }
    return checkmark;
}






function CalenderDateSelection(ObjectAddrs,value){ 
    var temp = "";
      temp = value.split(/[\/ | -]/g);
  
  var leapYear = false;
  if((temp[2]>=1800) &&(temp[2]<2500)){
  if(temp[2]%4 == 0)
    {
        if( temp[2]%100 == 0)
        {
            // year is divisible by 400, hence the year is a leap year
            if ( temp[2]%400 == 0){ 
            leapYear = true;
            if((temp[0]==1)||(temp[0]==3)||(temp[0]==5)||(temp[0]==7)||(temp[0]==8)||(temp[0]==10)||(temp[0]==12)){ 
              if((temp[1]>0) &&(temp[1]<=31)){ 
                ObjectAddrs.setText(value)
              }else{ 
                Log.Message("Date is invalid");
              }
              
            }else{ 
            if(temp[0]==2){ 
              if((temp[1]>0) &&(temp[1]<30)){ 
                ObjectAddrs.setText(value)
              }else{ 
                Log.Message("Date is invalid");
              }
            }else if((temp[0]>0) &&(temp[0]<13)){ 
              if((temp[1]>0) &&(temp[1]<31)){ 
                ObjectAddrs.setText(value)
              }else{ 
                Log.Message("Date is invalid");
              }
            }else{ 
               Log.Message("Month is invalid");
            }  
        }
    }      
        }
        
    }
     if(!leapYear){ 
        if((temp[0]==1)||(temp[0]==3)||(temp[0]==5)||(temp[0]==7)||(temp[0]==8)||(temp[0]==10)||(temp[0]==12)){ 
              if((temp[1]>0) &&(temp[1]<=31)){ 
                ObjectAddrs.setText(value)
              }else{ 
                Log.Message("Date is invalid");
              }
              
            }else{ 
            if(temp[0]==2){ 
              if((temp[1]>0) &&(temp[1]<29)){ 
                ObjectAddrs.setText(value)
              }else{ 
                Log.Message("Date is invalid");
              }
            }else if((temp[0]>0) &&(temp[0]<13)){ 
              if((temp[1]>0) &&(temp[1]<31)){ 
                ObjectAddrs.setText(value)
              }else{ 
                Log.Message("Date is invalid");
              }
            }else{ 
               Log.Message("Month is invalid");
            }  
        }
  }
    }else{ 
      Log.Message("Year is invalid it should Between 1800-2499");
    }
}





function StartTime(){ 
    var dif;
    var STIME="";
    var TodayValue = aqDateTime.Today();
    var StringTodayValue = aqConvert.DateTimeToStr(TodayValue);
    var EncodedDate = aqConvert.DateTimeToFormatStr(StringTodayValue,"%d%#B%Y"); 
//    Log.Message(EncodedDate)
    STIME = EncodedDate+" "+getFormattedCurrentTime();
//    Log.Message("Start DATE & TIME :"+EncodedDate +" "+STIME)
    var start = STIME.split(":");
    if(start[1]>0){ 
    dif = Number(start[2]) + Number(start[1]*60);
    }
    if(start[0]>0){ 
    dif = dif + Number(start[0]*60*60);
    }

return STIME;
}
function getFormattedCurrentTime(){
    TodayValue = aqConvert.DateTimeToFormatStr(aqDateTime.Time(), "%H:%M:%S");
    return TodayValue;
}




function Login_Match(Approve_Level,LoginEmp,HRData,comID){ 
login_satuts = true;
aqUtils.Delay(3000, Indicator.Text);;
var UserPasswd = [];
var z =0;
for(var i=0;i<Approve_Level.length;i++){ 
if((Approve_Level[i].indexOf("OpCo")!=-1) && (comID!="")){
Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,comID);
}
// Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,"1710");  //GCD2_Company No- level[0]
if(Approve_Level[i].indexOf("SSC - Biller")==-1){
Approve_Level[i] = Approve_Level[i].replace(/- Billers/g,"- Agency - Biller");
}
if(Approve_Level[i].indexOf("SSC - Billers")!=-1){
Approve_Level[i] = Approve_Level[i].replace(/SSC - Billers/g,"SSC IN -  Biller");
}

var tempLevel = Approve_Level[i].split("*");
ifGotIT = true;
var level = 0;
for(var j=2;j<tempLevel.length;j++){ 
level++;
if((tempLevel[j].indexOf(" (")!=-1) && (tempLevel[j].indexOf(")")!=-1)){
var temp = tempLevel[j].replace(" (","*");
temp = temp.replace(")","");
//Log.Message("temp :"+temp)
var tempSplit = temp.split("*");

for(var k=0;k<LoginEmp.length;k++){
  var A_temp = LoginEmp[k].split("*");
//    Log.Message("tempSplit[0] :"+tempSplit[0]);
//    Log.Message("A_temp[0] :"+A_temp[0]);
//    Log.Message("tempSplit[1] :"+tempSplit[1]);
//    Log.Message("A_temp[1] :"+A_temp[1]);
 if(tempSplit[0]==A_temp[0]){ 
 if(level==1){
    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"0";
    }
    else{ 
    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"1";  
    }
//     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[2];
   Log.Message(UserPasswd[z]);
   z++;
   ifGotIT = false;
   break; 
      
 }else{ 
 if(tempSplit[1]==A_temp[2]){ 
 if(level==1){
    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"0";
    }else{ 
    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"1";
    }
//     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[2];

   Log.Message(UserPasswd[z]);
   z++;
   ifGotIT = false;
   break;     
 }     
 }
    
}
if(!ifGotIT){ 
  break;
}
}
  
if((tempLevel[j].indexOf("SSC -")!=-1) || (tempLevel[j].indexOf("Central Team -")!=-1)){ 
//    Log.Message("tempLevel[j] :"+tempLevel[j]);
   if(tempLevel[j].indexOf("Central Team - Client Management")!=-1){ 
    temp2 = "Central Team - Client Account Management";
  }
  else if(tempLevel[j].indexOf("Central Team - Vendor Management")!=-1){ 
    temp2 = "Central Team - Vendor Account Management";
  }
  else if(tempLevel[j].indexOf("SSC - Expense Cashiers")!=-1){ 
    temp2 = "SSC - Cashier";
  }else{ 
    temp2 = tempLevel[j];
  }
for(var k=0;k<LoginEmp.length;k++){
  var A_temp = LoginEmp[k].split("*");  
 if(temp2==A_temp[1]){ 
 if(level==1){
   UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"0";
   }
   else{ 
   UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"1"; 
   }
//     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
     
   Log.Message(UserPasswd[z]);
   z++;
   ifGotIT = false;
   break;     
 }
}  

if(!ifGotIT){ 
  break;
}
}
  
  
  
if((tempLevel[j].indexOf(" (")==-1) && (tempLevel[j].indexOf(")")==-1) && 
(tempLevel[j].indexOf("SSC -")==-1) && (tempLevel[j].indexOf("Central Team -")==-1)){ 
    
for(var k=0;k<LoginEmp.length;k++){
  var A_temp = LoginEmp[k].split("*");
  if(A_temp[0]==tempLevel[j]){  // Better  to use level[j].indexOf(LoginArrays[k])
  if(level==1){
  UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"0";
  }else{ 
  UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]+"*"+"1";
  }
  Log.Message(UserPasswd[z]);
   z++;
   ifGotIT = false;
   break;     
 }
}
if(!ifGotIT){ 
  break;
}
}
  
if((tempLevel[j].indexOf(" (")!=-1) && (tempLevel[j].indexOf(")")!=-1)){

var temp = tempLevel[j].replace(" (","*");
temp = temp.replace(")","");
//Log.Message("temp :"+temp)
var tempSplit = temp.split("*");

for(var k=0;k<HRData.length;k++){
  var A_temp = HRData[k].split("*");
//    Log.Message("tempSplit[0] :"+tempSplit[0]);
//    Log.Message("A_temp[0] :"+A_temp[0]);
//    Log.Message("tempSplit[1] :"+tempSplit[1]);
//    Log.Message("A_temp[1] :"+A_temp[1]);
 if(tempSplit[1]==A_temp[1]){ 
 if(level==1){
   UserPasswd[z]  = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123"+"*"+"0";
   }else{ 
   UserPasswd[z]  = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123"+"*"+"1";
   }
//     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123";
   Log.Message(UserPasswd[z]);
   z++;
   ifGotIT = false;
   break;     
 }
    
}
if(!ifGotIT){ 
  break;
}
}
 
  
}
if(ifGotIT){ 
  Log.Warning("UserName and Password is Not Matched for Approver and Substitute :"+Approve_Level[i]);
  login_satuts = false;
  break;
}
  
}

return UserPasswd;
}





function DropDownList(value,feild,Address){ 
var checkMark = false;
eval(Sys_Maconomy_Parent).Refresh();
var list = "";
try{
  aqUtils.Delay(100,"Loading Dropdown Values");
  list = eval(Sys_Maconomy_Parent).SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
  }
  catch(e){ 
   Address.Click(); 
   list = eval(Sys_Maconomy_Parent).SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
  }
  var Add_Visible4 = true;
  while(Add_Visible4){
  if(list.isEnabled()){
  Add_Visible4 = false;
      for(var i=0;i<list.getItemCount();i++){ 
        if(list.getItem(i).getText_2(0)!=null){ 
          if(list.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
            list.Keys("[Enter]");
            aqUtils.Delay(1000, "Waiting to find Object");;
            checkMark = true;
            ValidationUtils.verify(true,true,feild+" is selected in Maconomy");
            break;
          }else{
            list.Keys("[Down]");
          }
          
        }else{ 
        Log.Message("i :"+i);
        Log.Message(list.getItem(i).getText_2(0).OleValue.toString().trim());
          list.Keys("[Down]");
        }

      }
  }
  }
  return checkMark;
}


function Credential(excelsheet,sheetname) {
var LoginEmp = [];
  var xlDriver = DDT.ExcelDriver(excelsheet, sheetname, false);
var id =0;
var colsList = [];

 for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){   
   colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
 }
   while (!DDT.CurrentDriver.EOF()) {
   var temp ="";
    for(var idx=0;idx<colsList.length;idx++){  
     if(xlDriver.Value(colsList[idx])!=null){
    temp = temp+xlDriver.Value(colsList[idx]).toString().trim()+"*";
    }
    else{ 
      temp = temp+"*";
    }
    }
//      Log.Message(temp)
   LoginEmp[id]=temp;
   id++;     
   xlDriver.Next();
   }
   DDT.CloseDriver(xlDriver.Name);
   return LoginEmp;
}


//function Rests(uname,pwd){ 
//aqUtils.Delay(5000, Indicator.Text);;
//      Sys.Desktop.KeyDown(0x12); //Alt
//    Sys.Desktop.KeyDown(0x46); //F
//    Sys.Desktop.KeyDown(0x52); //R 
//     Sys.Desktop.KeyUp(0x46); //Alt
//    Sys.Desktop.KeyUp(0x12);     
//     Sys.Desktop.KeyUp(0x52); //R
//aqUtils.Delay(65000, Indicator.Text);;
//     var usernameAddr = eval(Sys_Maconomy_Parent).SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 1);
//    var pwdAddr = eval(Sys_Maconomy_Parent).SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 2);
//    var btnLogin = eval(Sys_Maconomy_Parent).SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Login");
//    usernameAddr.SetFocus();
//    usernameAddr.setText(uname);
//    pwdAddr.setText(pwd);
//    btnLogin.click();
//    
//}


function getExcelData(workBook,sheetName) { 
  excelData =[];  
  
  var i = 0;
  for(var sheet = 0;sheet<sheetName.length;sheet++){
  var colsList = [];
  var xlDriver = DDT.ExcelDriver(workBook, sheetName[sheet], true);
     Log.Message(workBook)
     Log.Message(sheetName[sheet])
  for(var idx=0;idx<DDT.CurrentDriver.ColumnCount;idx++){ 
     
   colsList[idx] = DDT.CurrentDriver.ColumnName(idx);
     
 }
  var data = "";

  while (!DDT.CurrentDriver.EOF()) {
  data = "";
  for(var idx=0;idx<colsList.length;idx++){ 
  try{
     data = data + xlDriver.Value(colsList[idx]).toString().trim() + "*";
     excelData[i] = data;
       
     }
     catch(err)
     {
     data = data +"*";
     excelData[i] = data;
     }
       
     }
//     Log.Message("Excel LENGTH :"+data+"&"+data.length)
    // Log.Message("EXCELDATA :"+excelData[i]);      
     i++;
   xlDriver.Next();
  }
    
  DDT.CloseDriver(xlDriver.Name);
  }
 // Log.Message("completed reading excel data, data length::"+excelData.length);
  return excelData;
  
}


function companyNumber(companyName,wizName,comapany){ 
if(comapany!=""){
  companyName.Click();
  aqUtils.Delay(1000, Indicator.Text);;
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
  aqUtils.Delay(3000, Indicator.Text);;
  ///=========================
ExcelUtils.setExcelName(Project.Path+EnvParams.Opco, "Company info", true);
var cmpName = ExcelUtils.getRowData1("Company Name")
if((cmpName==null)||(cmpName=="")){ 
ValidationUtils.verify(false,true,"Company Name is Not available in ConfigSheet");
}
var tableList = [];
var tl = 0;

  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  do{
  aqUtils.Delay(5000, Indicator.Text);;
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
  tableList[tl] = table.getItem(i).getText_2(1).OleValue.toString().trim()
  tl++;
  }
    }
    var tab = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
    var tabVisible = tab.wEnabled(0,true)
    if(tabVisible){ 
      tab.Click(-1,-1);
    }
    }while(tabVisible)
    var compStatus = false;
    for(var cnt = 0;cnt<tableList.length;cnt++){ 
      if(tableList[cnt]!=cmpName){ 
        do{ 
        Log.Warning("Unwanted Data is available in Maconomy :");
        }while(false)
        Log.Message(tableList[cnt]);
      }else{ 
        compStatus = true;
      }
    }
    
    if(!compStatus)
    ValidationUtils.verify(false,true,"Company Name is Not available in Maconomy");
  ///=========================
  
  var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(comapany.toString().trim());
  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  aqUtils.Delay(5000, Indicator.Text);;
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==comapany.toString().trim()){ 
     var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        cancel.Click();
        aqUtils.Delay(1000, Indicator.Text);;
        companyName.setText("");
        ValidationUtils.verify(false,true,"Company Number is not listed in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
    cancel.Click();
    aqUtils.Delay(1000, Indicator.Text);;
    ValidationUtils.verify(false,true,"Company Number is not listed in Maconomy");
    companyName.setText("");
  } 
  }


}





function DepatmentValidation(Depart,wizName,department){ 
   if(department!=""){
  Depart.Click();
  aqUtils.Delay(1000, Indicator.Text);;
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
  aqUtils.Delay(3000, Indicator.Text);;
  //====================================
var sheet = [];
sheet[0] = "Department";
//sheet[1] = "UDepartment";
 var ExcelData = WorkspaceUtils.getExcelData(Project.Path+EnvParams.Opco,sheet);
var tableList = [];
var tl = 0;

  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  do{
  aqUtils.Delay(5000, Indicator.Text);;
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
  tableList[tl] = table.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+table.getItem(i).getText_2(1).OleValue.toString().trim()+"*";
  Log.Message(tableList[tl])
  tl++;
  }
    }
    var tab = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
    var tabVisible = tab.wEnabled(1,true)
    if(tabVisible){ 
      tab.Click(-1,-1);
    }
    }while(tabVisible)
    
    
    var stat = true;
    for(var exl =0;exl<ExcelData.length;exl++){
        var compStatus = false;
    for(var cnt = 0;cnt<tableList.length;cnt++){       
      if(ExcelData[exl].toLowerCase()==tableList[cnt].toLowerCase()){ 
       compStatus = true;
       break;
      }
      }
      if(!compStatus){ 
        if(stat){
        Log.Warning("Some Expected Department are missing in Maconomy :");
        stat = false;
        }
        var splits = ExcelData[exl].split("*");
        Log.Message(splits[0]+"  "+splits[1]);
      }
    }
    
   var stat = true; 
    for(var cnt = 0;cnt<tableList.length;cnt++){
      var compStatus = false;
    for(var exl =0;exl<ExcelData.length;exl++){
     if(tableList[cnt].toLowerCase()==ExcelData[exl].toLowerCase()){ 
       compStatus = true;
       break;
      }
      }
      if(!compStatus){ 
        if(stat){
        Log.Warning("Some Unwanted Department data is available in Maconomy :");
        stat = false;
        }
        var splits = tableList[cnt].split("*");
        Log.Message(splits[0]+"  "+splits[1]);
      }
    }
    
    var compStatus = false;
    for(var exl =0;exl<ExcelData.length;exl++){
      var splits = ExcelData[exl].split("*");
      if(splits[0]==department.toString().trim()){ 
        compStatus = true;
        break;
      }
      }
  if(!compStatus){
    ValidationUtils.verify(false,true,"Given Department in Datasheet is not available in ConfigPack");
    }
  //====================================
  var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(department.toString().trim());
  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  aqUtils.Delay(5000, Indicator.Text);;
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==department.toString().trim()){ 
     var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        cancel.Click();
        aqUtils.Delay(1000, Indicator.Text);;
        Depart.setText("");
        ValidationUtils.verify(false,true,"Department Number is not listed in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
    cancel.Click();
    aqUtils.Delay(1000, Indicator.Text);;
    ValidationUtils.verify(false,true,"Department Number is not listed in Maconomy");
    Depart.setText("");
  }
        }


}




function BusinessUnitValidation(BussUnit,wizName,buss_unit){ 
    if(buss_unit!=""){
  BussUnit.Click();
  aqUtils.Delay(1000, Indicator.Text);;
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
  aqUtils.Delay(3000, Indicator.Text);;
  //==========================================================
  var sheet = [];
sheet[0] = "Business Unit";
//sheet[1] = "UDepartment";
 var ExcelData = WorkspaceUtils.getExcelData(Project.Path+EnvParams.Opco,sheet);
var tableList = [];
var tl = 0;

  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  do{
  aqUtils.Delay(5000, Indicator.Text);;
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
  tableList[tl] = table.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+table.getItem(i).getText_2(1).OleValue.toString().trim()+"*";
  Log.Message(tableList[tl])
  tl++;
  }
    }
    var tab = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
    var tabVisible = tab.wEnabled(1,true)
    if(tabVisible){ 
      tab.Click(-1,-1);
    }
    }while(tabVisible)
    
    
    var stat = true;
    for(var exl =0;exl<ExcelData.length;exl++){
        var compStatus = false;
    for(var cnt = 0;cnt<tableList.length;cnt++){       
      if(ExcelData[exl].toLowerCase()==tableList[cnt].toLowerCase()){ 
       compStatus = true;
       break;
      }
      }
      if(!compStatus){ 
        if(stat){
        Log.Warning("Some Expected Business Unit are missing in Maconomy :");
        stat = false;
        }
        var splits = ExcelData[exl].split("*");
        Log.Message(splits[0]+"  "+splits[1]);
      }
    }
    
   var stat = true; 
    for(var cnt = 0;cnt<tableList.length;cnt++){
      var compStatus = false;
    for(var exl =0;exl<ExcelData.length;exl++){
     if(tableList[cnt].toLowerCase()==ExcelData[exl].toLowerCase()){ 
       compStatus = true;
       break;
      }
      }
      if(!compStatus){ 
        if(stat){
        Log.Warning("Some Unwanted Business Unit data is available in Maconomy :");
        stat = false;
        }
        var splits = tableList[cnt].split("*");
        Log.Message(splits[0]+"  "+splits[1]);
      }
    }
    
    var compStatus = false;
    for(var exl =0;exl<ExcelData.length;exl++){
      var splits = ExcelData[exl].split("*");
      if(splits[0]==buss_unit.toString().trim()){ 
        compStatus = true;
        break;
      }
      }
  if(!compStatus){
    ValidationUtils.verify(false,true,"Given Business Unit in Datasheet is not available in ConfigPack");
    }
    //==============================================================
  var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(buss_unit.toString().trim());
  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  aqUtils.Delay(5000, Indicator.Text);;
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==buss_unit.toString().trim()){ 
     var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        cancel.Click();
        aqUtils.Delay(1000, Indicator.Text);;
        BussUnit.setText("");
        ValidationUtils.verify(false,true,"Business Unit Number is not listed in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
    cancel.Click();
    aqUtils.Delay(1000, Indicator.Text);;
    ValidationUtils.verify(false,true,"Business Unit Number is not listed in Maconomy");
    BussUnit.setText("");
  } 
            }
 

}






function EmpCategoryValidation(EmpCat,wizName,EmpCategory){ 
   if(EmpCategory!=""){
  EmpCat.Click();
  aqUtils.Delay(1000, Indicator.Text);;
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
  aqUtils.Delay(3000, Indicator.Text);;
  //====================================
var sheet = [];
sheet[0] = "Employee Categories";
//sheet[1] = "UDepartment";
 var ExcelData = WorkspaceUtils.getExcelData(Project.Path+EnvParams.Opco,sheet);
var tableList = [];
var tl = 0;

  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  do{
  aqUtils.Delay(5000, Indicator.Text);;
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
  tableList[tl] = table.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+table.getItem(i).getText_2(1).OleValue.toString().trim()+"*";
  Log.Message(tableList[tl])
  tl++;
  }
    }
    var tab = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
    Sys.HighlightObject(tab);
    var tabVisible = tab.wEnabled(1,true)
    Log.Message(tabVisible);
    if(tabVisible){ 
      tab.Click(-1,-1);
    }
    }while(tabVisible)
    
    
    var stat = true;
    for(var exl =0;exl<ExcelData.length;exl++){
        var compStatus = false;
    for(var cnt = 0;cnt<tableList.length;cnt++){       
      if(ExcelData[exl].toLowerCase()==tableList[cnt].toLowerCase()){ 
       compStatus = true;
       break;
      }
      }
      if(!compStatus){ 
        if(stat){
        Log.Warning("Some Expected Department are missing in Maconomy :");
        stat = false;
        }
        var splits = ExcelData[exl].split("*");
        Log.Message(splits[0]+"  "+splits[1]);
      }
    }
    
   var stat = true; 
    for(var cnt = 0;cnt<tableList.length;cnt++){
      var compStatus = false;
    for(var exl =0;exl<ExcelData.length;exl++){
     if(tableList[cnt].toLowerCase()==ExcelData[exl].toLowerCase()){ 
       compStatus = true;
       break;
      }
      }
      if(!compStatus){ 
        if(stat){
        Log.Warning("Some Unwanted Department data is available in Maconomy :");
        stat = false;
        }
        var splits = tableList[cnt].split("*");
        Log.Message(splits[0]+"  "+splits[1]);
      }
    }
    
    var compStatus = false;
    for(var exl =0;exl<ExcelData.length;exl++){
      var splits = ExcelData[exl].split("*");
      if(splits[0]==EmpCategory.toString().trim()){ 
        compStatus = true;
        break;
      }
      }
  if(!compStatus){
    ValidationUtils.verify(false,true,"Given Department in Datasheet is not available in ConfigPack");
    }
  //====================================
  var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(EmpCategory.toString().trim());
  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  aqUtils.Delay(5000, Indicator.Text);;
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==EmpCategory.toString().trim()){ 
     var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        cancel.Click();
        aqUtils.Delay(1000, Indicator.Text);;
        EmpCat.setText("");
        ValidationUtils.verify(false,true,"Department Number is not listed in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
    cancel.Click();
    aqUtils.Delay(1000, Indicator.Text);;
    ValidationUtils.verify(false,true,"Department Number is not listed in Maconomy");
    EmpCat.setText("");
  }
        }
}


function config_with_Maconomy_Validation(Obj_Address,wizName,value,ExcelData,fieldName){ 
var temp = "";
   if(value!=""){
     Sys.HighlightObject(Obj_Address);
  Obj_Address.Click();
  aqUtils.Delay(2000, wizName);;
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);

var tableList = [];
var tl = 0;
Log.Message(Sys_Maconomy_Parent);
Log.Message(eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).FullName)
//  aqUtils.Delay(5000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  waitForObj(serch);
  serch.Click();
// var Add_Visible0 = true;
//  while(Add_Visible0){
//    if(serch.isEnabled()){
//      serch.HoverMouse();
//      Sys.HighlightObject(serch);
//      serch.Click();
//      Add_Visible0 = false;
//      }
//  }
  
//  Sys.HighlightObject(serch);
//  if(serch.isEnabled())
//  serch.Click();
//  else{ 
//    aqUtils.Delay(3000, Indicator.Text);;
//   serch.Click(); 
//  }
  
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  do{
//  aqUtils.Delay(5000, Indicator.Text);;
//  var Add_Visible0 = true;
//  while(Add_Visible0){
//    if(OK.isEnabled()){
  Sys.HighlightObject(table);
  waitForObj(OK);
        
          var itemCount = table.getItemCount();
//          Log.Message(itemCount)
          if(itemCount>0){ 
          for(var i=0;i<itemCount;i++){
          tableList[tl] = table.getItem(i).getText_2(0).OleValue.toString().trim()+"-"+table.getItem(i).getText_2(1).OleValue.toString().trim();
          tl++;
                          }
                }
//        Add_Visible0 = false;
//      }
//  }
    var tab = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
    var tabVisible = tab.wEnabled(1,true)
    if(tabVisible){ 
      tab.Click(-1,-1);
    }
    }while(tabVisible)
    
    
    var stat = true;
    for(var exl =0;exl<ExcelData.length;exl++){
        var compStatus = false;
    var bb1 = "";
        for(var cnt = 0;cnt<tableList.length;cnt++){
      if(ExcelData[exl].toLowerCase()==tableList[cnt].toLowerCase()){ 
       compStatus = true;
       break;
      }
      }
      if(!compStatus){ 
        if(stat){
        Log.Warning("Some Expected "+fieldName+" are missing in Maconomy :");
        ReportUtils.logStep("WARNING","Some Expected "+fieldName+" are missing in Maconomy :")
        stat = false;
        }
        var splits = []; 
//        ExcelData[exl].split("*");  
//        Log.Message(ExcelData[exl]);
        splits[0] = ExcelData[exl].substring(0, ExcelData[exl].indexOf("-"));
        splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf("-")+1);
        if(splits[0]==value.toString().trim()){ 
        ValidationUtils.verify(false,true,"Given "+fieldName+" in Datasheet is not available in Maconomy");
        }else{
        Log.Message(splits[0]+"  "+splits[1]);
        ReportUtils.logStep("INFO",splits[0]+"  "+splits[1])
        }
      }
    }
    
   var stat = true; 
    for(var cnt = 0;cnt<tableList.length;cnt++){
      var compStatus = false;
    for(var exl =0;exl<ExcelData.length;exl++){
     if(tableList[cnt].toLowerCase()==ExcelData[exl].toLowerCase()){ 
       compStatus = true;
       break;
      }
      }
//      if(!compStatus){ 
//        if(stat){
//        Log.Warning("Some Unwanted Department data is available in Maconomy :");
//        stat = false;
//        }
//        var splits = tableList[cnt].split("*");
//        Log.Message(splits[0]+"  "+splits[1]);
//      }
    }
    
    var compStatus = false;
//    Log.Message(ExcelData.length);
    for(var exl =0;exl<ExcelData.length;exl++){
//      var splits = ExcelData[exl].split("*");
        var splits = []; 
//        ExcelData[exl].split("*");        
        splits[0] = ExcelData[exl].substring(0, ExcelData[exl].indexOf("-"));
//        Log.Message("splits[0] :"+splits[0]);
        splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf("-")+1);
//        Log.Message("value :"+value);
//        Log.Message(splits[0]==value.toString().trim())
      if(splits[0]==value.toString().trim()){ 
        compStatus = true;
        break;
      }
//      Log.Message("splits[0] :"+splits[0] +"  value :"+value);
//      Log.Message(splits[0]==value.toString().trim());
      }
  if(!compStatus){
    ValidationUtils.verify(false,true,"Given "+fieldName+" in Datasheet is not available in ConfigPack");
    }
  //====================================
  var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(value.toString().trim());
//  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
 waitForObj(serch);
 serch.Click();
// var Add_Visible0 = true;
//  while(Add_Visible0){
//    if(serch.isEnabled()){
//      serch.HoverMouse();
//      Sys.HighlightObject(serch);
//      serch.Click();
//      Add_Visible0 = false;
//      }
//  }  


//Sys.HighlightObject(serch);
//   if(serch.isEnabled())
//  serch.Click();
//  else{ 
//    aqUtils.Delay(3000, Indicator.Text);;
//   serch.Click(); 
//  }
//  aqUtils.Delay(4000, Indicator.Text);;
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  waitForObj(OK);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value.toString().trim()){ 
    temp = table.getItem(i).getText_2(1).OleValue.toString().trim();
     var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
     waitForObj(OK);
     OK.Click();
     
//    var Add_Visible0 = true;
//    while(Add_Visible0){
//      if(OK.isEnabled()){
//        OK.HoverMouse();
//        Sys.HighlightObject(OK);
//        OK.Click();
//        Add_Visible0 = false;
//        }
//    } 
    ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
    break;
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){
        var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        waitForObj(cancel);
        cancel.Click();
//          var Add_Visible0 = true;
//          while(Add_Visible0){
//            if(cancel.isEnabled()){
//              cancel.HoverMouse();
//              Sys.HighlightObject(cancel);
//              cancel.Click();
//              Add_Visible0 = false;
//              }
//          } 

//if(cancel.isEnabled()){
//      cancel.HoverMouse();
//      ReportUtils.logStep_Screenshot();
//      cancel.Click();
//      }
//      else{ 
//      aqUtils.Delay(3000, Indicator.Text);;
//      cancel.HoverMouse();
//      ReportUtils.logStep_Screenshot();
//       cancel.Click(); 
//      }
//        aqUtils.Delay(1000, Indicator.Text);;
        Sys.HighlightObject(Obj_Address);
        Obj_Address.setText("");
        ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
    waitForObj(cancel);
    cancel.Click();
//          var Add_Visible0 = true;
//          while(Add_Visible0){
//            if(cancel.isEnabled()){
//              cancel.HoverMouse();
//              Sys.HighlightObject(cancel);
//              cancel.Click();
//              Add_Visible0 = false;
//              }
//          }  
            
//if(cancel.isEnabled()){
//  cancel.HoverMouse();
//  ReportUtils.logStep_Screenshot();
//  cancel.Click();
//  }
//  else{ 
//    aqUtils.Delay(3000, Indicator.Text);;
//  cancel.HoverMouse();
//  ReportUtils.logStep_Screenshot();
//   cancel.Click(); 
//  }
//    aqUtils.Delay(1000, Indicator.Text);;
    ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    Sys.HighlightObject(Obj_Address);
    Obj_Address.setText("");
  }
        }
return temp;
}

function config_with_Maconomy_Validation_Name_column2(Obj_Address,wizName,value,ExcelData,fieldName){ 
var temp = "";
   if(value!=""){
  Obj_Address.Click();
  aqUtils.Delay(1000, wizName);;
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
//  aqUtils.Delay(3000, Indicator.Text);;
  //====================================
//var sheet = [];
//sheet[0] = "value";
//sheet[1] = "UDepartment";
//var ExcelData = ExlArray;
var tableList = [];
var tl = 0;

//  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  waitForObj(serch)
//  Sys.HighlightObject(serch);
  serch.Click();
  
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
//  Sys.HighlightObject(table);
  do{
  waitForObj(table);
  waitForObj(OK);
//  aqUtils.Delay(5000, Indicator.Text);;
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
  tableList[tl] = table.getItem(i).getText_2(0).OleValue.toString().trim()+"-"+table.getItem(i).getText_2(2).OleValue.toString().trim();
//  Log.Message(tableList[tl])
  tl++;
  }
    }
    var tab = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
    var tabVisible = tab.wEnabled(1,true)
    if(tabVisible){ 
      tab.Click(-1,-1);
    }
    }while(tabVisible)
    
    
    var stat = true;
    for(var exl =0;exl<ExcelData.length;exl++){
//    Log.Message(ExcelData[exl])
        var compStatus = false;
        for(var cnt = 0;cnt<tableList.length;cnt++){  
//        Log.Message(tableList[cnt])     
      if(ExcelData[exl].toLowerCase()==tableList[cnt].toLowerCase()){ 
       compStatus = true;
       break;
      }
      }
      if(!compStatus){ 
        if(stat){
        Log.Warning("Some Expected "+fieldName+" are missing in Maconomy :");
        ReportUtils.logStep("WARNING","Some Expected "+fieldName+" are missing in Maconomy :")
        stat = false;
        }
        var splits = []; 
//        ExcelData[exl].split("*");        
        splits[0] = ExcelData[exl].substring(0, ExcelData[exl].indexOf("-"));
        splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf("-")+1);
        Log.Message(splits[0]+"  "+splits[1]);
        ReportUtils.logStep("INFO",splits[0]+"  "+splits[1])
      }
    }
    
   var stat = true; 
    for(var cnt = 0;cnt<tableList.length;cnt++){
      var compStatus = false;
    for(var exl =0;exl<ExcelData.length;exl++){
     if(tableList[cnt].toLowerCase()==ExcelData[exl].toLowerCase()){ 
       compStatus = true;
       break;
      }
      }
//      if(!compStatus){ 
//        if(stat){
//        Log.Warning("Some Unwanted Department data is available in Maconomy :");
//        stat = false;
//        }
//        var splits = tableList[cnt].split("*");
//        Log.Message(splits[0]+"  "+splits[1]);
//      }
    }
    
    var compStatus = false;
//    Log.Message(ExcelData.length);
    for(var exl =0;exl<ExcelData.length;exl++){
//      var splits = ExcelData[exl].split("*");
//Log.Message(ExcelData[exl])
        var splits = []; 
//        ExcelData[exl].split("*");        
        splits[0] = ExcelData[exl].substring(0, ExcelData[exl].indexOf("-"));
//        Log.Message("splits[0] :"+splits[0]);
        splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf("-")+1);
//        Log.Message("value :"+value);
//        Log.Message(splits[0]==value.toString().trim())
      if(splits[0]==value.toString().trim()){ 
        compStatus = true;
        break;
      }
//      Log.Message("splits[0] :"+splits[0] +"  value :"+value);
//      Log.Message(splits[0]==value.toString().trim());
      }
  if(!compStatus){
    ValidationUtils.verify(false,true,"Given "+fieldName+" in Datasheet is not available in ConfigPack");
    }
  //====================================
  var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(value.toString().trim());
//  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
//  Sys.HighlightObject(serch);
  waitForObj(serch);
  serch.Click();
//  aqUtils.Delay(5000, Indicator.Text);;
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  waitForObj(table);
  waitForObj(OK);
//  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value.toString().trim()){ 
    temp = table.getItem(i).getText_2(2).OleValue.toString().trim();
     var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
        OK.HoverMouse();
//        ReportUtils.logStep_Screenshot();
        OK.Click();
        ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        cancel.HoverMouse();
        ReportUtils.logStep_Screenshot();
        cancel.Click();
//        aqUtils.Delay(1000, Indicator.Text);;
        waitForObj(Obj_Address);
        Obj_Address.setText("");
        ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
    cancel.HoverMouse();
    ReportUtils.logStep_Screenshot();
    cancel.Click();
//    aqUtils.Delay(1000, Indicator.Text);;
    waitForObj(Obj_Address);
    ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    Obj_Address.setText("");
  }
        }
return temp;
}


function Config_with_Maconomy_templateValidation_(Obj_Address,wizName,value,ExlArray,Job_Type,comapany,Jobgroup,fieldName){ 
  if(value!=""){
  Obj_Address.Click();
  aqUtils.Delay(1000, Indicator.Text);;
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
  aqUtils.Delay(3000, Indicator.Text);;
  //==============================================
  var JG = "";
  if(Jobgroup.toString().trim()=="Client Billable")
  JG = "RB";
  if(Jobgroup.toString().trim()=="Client Non-Billable")
  JG = "RN";
  if(Jobgroup.toString().trim()=="Internal")
  JG = "RI";
//var sheet = [];
//sheet[0] = "Job Setup";
if(Jobgroup.toString().trim()!="Internal"){
var ExcelData = ExlArray;
var tableList = [];
var TemplateCode = [];
var tl = 0;

  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  do{
  aqUtils.Delay(5000, Indicator.Text);;
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
  tableList[tl] = table.getItem(i).getText_2(1).OleValue.toString().trim();
  TemplateCode[tl] = table.getItem(i).getText_2(0).OleValue.toString().trim();
  Log.Message(TemplateCode[tl]+"-"+tableList[tl])
//  Log.Message(tableList[tl])
  tl++;
  }
    }
    var tab = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
    var tabVisible = tab.wEnabled(0,true)
    if(tabVisible){ 
      tab.Click(-1,-1);
    }
    }while(tabVisible)
    var billable = [];
    var bil = 0;
    var compStatus = false;
//    Log.Message("Job_Type :"+Job_Type)
    for(var exl =0;exl<ExcelData.length;exl++){
//      var splits = ExcelData[exl].split("*");
      var splits = []; 
//        ExcelData[exl].split("*");  
//Log.Message(ExcelData[exl]);
if(ExcelData[exl].indexOf("Business")==-1){     
        splits[0] = ExcelData[exl].substring(0, ExcelData[exl].indexOf("-"));
        splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf("-")+1);
        }
else{ 
if(ExcelData[exl].indexOf("New Business")!=-1){ 
var temp = ExcelData[exl].split("-");
splits[0] = temp[0]+"-"+temp[1];  
splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf(splits[0])+(splits[0].length+1))  
//        splits[0] = ExcelData[exl].substring(0, ExcelData[exl].indexOf("-"));
//        splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf("-")+1);
  }
else{ 
var temp = ExcelData[exl].split("-");
splits[0] = temp[0]+"-"+temp[1]+"-"+temp[2];  
splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf(splits[0])+(splits[0].length+1))   
}
          
        }
//        Log.Message("splits[0] :"+splits[0])
//        Log.Message("splits[1] :"+splits[1])
//-------------------For INDIA ------------------------------
//      if(splits[0]==Job_Type.toString().trim()){ 
//        compStatus = true;
//        if(Jobgroup.toString().trim()!="Client Non-Billable"){
//    if(splits[1]=="Fixed Price - CP")
//    data = comapany+" - "+splits[0]+" - Fixed Price CP";
//    else if(splits[1]=="T&M - CP")
//    data = comapany+" - "+splits[0]+" - T&M CP";
//    if(splits[1]=="Fixed Price - BP")
//    data = comapany+" - "+splits[0]+" - Fixed Price BP";
//    else if(splits[1]=="T&M - BP")
//    data = comapany+" - "+splits[0]+" - T&M BP";
//    else if(splits[1]=="Retainer")
//    data = comapany+" - "+splits[0]+" - RT-CP";
//    else if((splits[1]=="Non Billable")||(splits[1]=="Non - Billable"))
//    data = comapany+" - "+splits[0]+" - Non-Billable";
//    else
//    data = comapany+" - "+splits[0]+" - "+splits[1]; 
//    }else{ 
//      data = comapany+" - "+splits[0]+" - Non-Billable";
//    }
//    billable[bil] = data;
////    Log.Message("Templete are availble in ConfigPack for this JobType :"+data);
//    bil++;
//      }
      
//----------------------------------------------------------------
      if(splits[0]==Job_Type.toString().trim()){ 
        compStatus = true;
        if(Jobgroup.toString().trim()!="Client Non-Billable"){
    if(splits[1]=="Fixed Price - CP")
    data = comapany+"-"+splits[0]+"-Fixed Price CP";
    else if(splits[1]=="T&M - CP")
    data = comapany+"-"+splits[0]+"-T&M CP";
    if(splits[1]=="Fixed Price - BP")
    data = comapany+"-"+splits[0]+"-Fixed Price BP";
    else if(splits[1]=="T&M - BP")
    data = comapany+"-"+splits[0]+"-T&M BP";
    else if(splits[1]=="Retainer")
    data = comapany+"-"+splits[0]+"-RT-CP";
    else if((splits[1]=="Non Billable")||(splits[1]=="Non - Billable"))
    data = comapany+"-"+splits[0]+"-Non-Billable";
    else
    data = comapany+"-"+splits[0]+"-"+splits[1]; 
    }else{ 
      data = comapany+"-"+splits[0]+"-Non-Billable";
    }
    billable[bil] = data;
//    Log.Message("Templete are availble in ConfigPack for this JobType :"+data);
    bil++;
      }
      }
  if(!compStatus){
    if((Jobgroup.toString().trim()=="Client Billable")||(Jobgroup.toString().trim()=="Client Non-Billable"))
//    if(Job_group.toString().trim()!="Internal")
    ValidationUtils.verify(false,true,"Selected JobType doesn't have any Templete in Config Sheet");
    else
    Log.Warning("Selected JobType doesn't have any Templete in Config Sheet")
    } 
    
    
    
    
    
    
    
    
    
    var stat = true;
    for(var exl =0;exl<billable.length;exl++){
        var compStatus = false;
    for(var cnt = 0;cnt<tableList.length;cnt++){ 
//      Log.Message("billable[exl] :"+billable[exl]);
//      Log.Message("tableList[cnt] :"+tableList[cnt]);      
      if(billable[exl].toLowerCase()==tableList[cnt].toLowerCase()){ 
//      Log.Message("billable[exl] :"+billable[exl]);
//      Log.Message("tableList[cnt] :"+tableList[cnt]);
       compStatus = true;
       break;
      }
      }
      if(!compStatus){ 
        if(stat){
        ValidationUtils.verify(false,true,"Some Expected Template are missing in Maconomy compared to ConfigPack:");
//        Log.Warning("Some Expected Template are missing in Maconomy :");
        stat = false;
        }
//        var splits = billable[exl].split("*");
       ReportUtils.logStep("INFO", billable[exl]);
//        Log.Message(billable[exl]);
      }
    }
    
   var stat = true; 
    for(var cnt = 0;cnt<tableList.length;cnt++){
      var compStatus = false;
    for(var exl =0;exl<billable.length;exl++){
//    if(ExcelData[exl].indexOf(JG)==0)       
      if(tableList[cnt].toLowerCase()==billable[exl].toLowerCase()){ 
       compStatus = true;
       break;
      }
      }
      if(!compStatus){ 
        if(stat){
//        ReportUtils.logStep("INFO", "Some Unwanted Templates data are available in Maconomy :");
//        Log.Warning("Some Unwanted Template data is available in Maconomy :");
        stat = false;
        }
//        var splits = tableList[cnt].split("*");
//       ReportUtils.logStep("INFO", tableList[cnt]);
//        Log.Message(tableList[cnt]);
      }
    }
    
    var compStatus = false;
    for(var tl=0;tl<TemplateCode.length;tl++){
//    var TempNo = TemplateCode[tl].substring(TemplateCode[tl].indexOf("_")+1)
      if(TemplateCode[tl].indexOf(value.toString().trim())!=-1){
//      Log.Message("TemplateCode[tl] :"+TemplateCode[tl]);
      Log.Message("value :"+value);
//      value = TemplateCode[tl];
        compStatus = true;
        break;
        }
      }
  if(!compStatus){
    ValidationUtils.verify(false,true,"Given TemplateNo in Datasheet is not available in ConfigPack");
    }
    }
  //==============================================
  var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
//  code.Keys("[Tab]");
//  aqUtils.Delay(2000, Indicator.Text);;
//  var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText("*"+value.toString().trim()+"*");
  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  aqUtils.Delay(5000, Indicator.Text);;
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim().indexOf(value.toString().trim())!=-1){ 
     var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
        OK.Click();
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        cancel.Click();
        aqUtils.Delay(1000, Indicator.Text);;
        Obj_Address.setText("");
        ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
    cancel.Click();
    aqUtils.Delay(1000, Indicator.Text);;
    ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    Obj_Address.setText("");
  }
        }
else{ 
    ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }   
    
}


function Config_with_Maconomy_templateValidation(Obj_Address,wizName,value,ExlArray,Job_Type,comapany,Jobgroup,fieldName){ 
  if(value!=""){
  Obj_Address.Click();
  aqUtils.Delay(1000, Indicator.Text);;
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
//  aqUtils.Delay(5000, Indicator.Text);;
  //==============================================
  var JG = "";
  if(Jobgroup.toString().trim()=="Client Billable")
  JG = "RB";
  if(Jobgroup.toString().trim()=="Client Non-Billable")
  JG = "RN";
  if(Jobgroup.toString().trim()=="Internal")
  JG = "RI";
//var sheet = [];
//sheet[0] = "Job Setup";
if(Jobgroup.toString().trim()!="Internal"){
var ExcelData = ExlArray;
var tableList = [];
var TemplateCode = [];
var tl = 0;

//  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  waitForObj(serch);
  Sys.HighlightObject(serch);
  serch.Click();
//  if(serch.isEnabled())
//  serch.Click();
//  else{ 
//    aqUtils.Delay(3000, Indicator.Text);;
//   serch.Click(); 
//  }
  
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);

  do{
  var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  Sys.HighlightObject(table);
  waitForObj(OK);
//  aqUtils.Delay(5000, Indicator.Text);;
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
  tableList[tl] = table.getItem(i).getText_2(1).OleValue.toString().trim();
  TemplateCode[tl] = table.getItem(i).getText_2(0).OleValue.toString().trim();
//  Log.Message(TemplateCode[tl]+"-"+tableList[tl])
//  Log.Message(tableList[tl])
  tl++;
  }
    }
    var tab = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
    var tabVisible = tab.wEnabled(0,true)
    if(tabVisible){ 
      tab.Click(-1,-1);
    }
    }while(tabVisible)
    var billable = [];
    var bil = 0;
    var compStatus = false;
//    Log.Message("Job_Type :"+Job_Type)
    for(var exl =0;exl<ExcelData.length;exl++){
//      var splits = ExcelData[exl].split("*");
      var splits = []; 
//        ExcelData[exl].split("*");  
//Log.Message(ExcelData[exl]);
if(ExcelData[exl].indexOf("Business")==-1){     
        splits[0] = ExcelData[exl].substring(0, ExcelData[exl].indexOf("-"));
        splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf("-")+1);
//        Log.Message(splits[0])
//        Log.Message(splits[1])
        }
else{ 
if(ExcelData[exl].indexOf("New Business")!=-1){ 
var temp = ExcelData[exl].split("-");
splits[0] = temp[0]+"-"+temp[1];  
splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf(splits[0])+(splits[0].length+1))  
//        splits[0] = ExcelData[exl].substring(0, ExcelData[exl].indexOf("-"));
//        splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf("-")+1);
  }
else{ 
var temp = ExcelData[exl].split("-");
splits[0] = temp[0]+"-"+temp[1]+"-"+temp[2];  
splits[1] = ExcelData[exl].substring(ExcelData[exl].indexOf(splits[0])+(splits[0].length+1))   
}
          
        }

//-------------------For INDIA ------------------------------
      if(splits[0].trim()==Job_Type.toString().trim()){ 
        compStatus = true;
        if(Jobgroup.toString().trim()!="Client Non-Billable"){
    if(splits[1]=="Fixed Price - CP")
    data = comapany+" - "+splits[0]+" - Fixed Price CP";
    else if(splits[1]=="T&M - CP")
    data = comapany+" - "+splits[0]+" - T&M CP";
    else if(splits[1]=="Fixed Price - BP")
    data = comapany+" - "+splits[0]+" - Fixed Price BP";
    else if(splits[1]=="T&M - BP")
    data = comapany+" - "+splits[0]+" - T&M BP";
    else if(splits[1]=="Retainer")
    data = comapany+" - "+splits[0]+" - RT-CP";
    else if((splits[1]=="Non Billable")||(splits[1]=="Non - Billable"))
    data = comapany+" - "+splits[0]+" - Non-Billable";
    else
    data = comapany+" - "+splits[0]+" - "+splits[1]; 
    }else{ 
      data = comapany+" - "+splits[0]+" - Non-Billable";
    }
    billable[bil] = data;
//    Log.Message("Templete are availble in ConfigPack for this JobType :"+data);
    bil++;
      }
      
//----------------------------------------------------------------
//      if(splits[0]==Job_Type.toString().trim()){ 
//        compStatus = true;
//        if(Jobgroup.toString().trim()!="Client Non-Billable"){
//    if(splits[1]=="Fixed Price - CP")
//    data = comapany+"-"+splits[0]+"-Fixed Price CP";
//    else if(splits[1]=="T&M - CP")
//    data = comapany+"-"+splits[0]+"-T&M CP";
//    else if(splits[1]=="Fixed Price - BP")
//    data = comapany+"-"+splits[0]+"-Fixed Price BP";
//    else if(splits[1]=="T&M - BP")
//    data = comapany+"-"+splits[0]+"-T&M BP";
//    else if(splits[1]=="Retainer")
//    data = comapany+"-"+splits[0]+"-RT-CP";
//    else if((splits[1]=="Non Billable")||(splits[1]=="Non - Billable"))
//    data = comapany+"-"+splits[0]+"-Non-Billable";
//    else
//    data = comapany+"-"+splits[0]+"-"+splits[1]; 
//    
//    
//    }else{ 
//      data = comapany+"-"+splits[0]+"-Non-Billable";
//    }
//    billable[bil] = data;
////    Log.Message("Templete are availble in ConfigPack for this JobType :"+data);
//    bil++;
//      }
      }
  if(!compStatus){
    if((Jobgroup.toString().trim()=="Client Billable")||(Jobgroup.toString().trim()=="Client Non-Billable"))
//    if(Job_group.toString().trim()!="Internal")
    ValidationUtils.verify(false,true,"Selected JobType doesn't have any Templete in Config Sheet");
    else
    Log.Warning("Selected JobType doesn't have any Templete in Config Sheet")
    } 
    
    
    
    
    
    
    
    
    
    var stat = true;
    for(var exl =0;exl<billable.length;exl++){
        var compStatus = false;
          Log.Message("billable[exl] :"+billable[exl]);
    for(var cnt = 0;cnt<tableList.length;cnt++){ 

      Log.Message("tableList[cnt] :"+tableList[cnt]);      
      if(tableList[cnt].toLowerCase().indexOf(billable[exl].toLowerCase())!=-1){ 
//      Log.Message("billable[exl] :"+billable[exl]);
//      Log.Message("tableList[cnt] :"+tableList[cnt]);
       compStatus = true;
       break;
      }
      }
//      if(!compStatus){ 
//        if(stat){
//        Log.Message(billable[exl])
//        ValidationUtils.verify(false,true,"Some Expected Template are missing in Maconomy compared to ConfigPack:");
//        stat = false;
//        }
//       ReportUtils.logStep("INFO", billable[exl]);
//      }
    }
    
   var stat = true; 
    for(var cnt = 0;cnt<tableList.length;cnt++){
      var compStatus = false;
    for(var exl =0;exl<billable.length;exl++){
//    if(ExcelData[exl].indexOf(JG)==0)       
      if(tableList[cnt].toLowerCase().indexOf(billable[exl].toLowerCase())!=-1){ 
       compStatus = true;
       break;
      }
      }
      if(!compStatus){ 
        if(stat){
//        ReportUtils.logStep("INFO", "Some Unwanted Templates data are available in Maconomy :");
//        Log.Warning("Some Unwanted Template data is available in Maconomy :");
        stat = false;
        }
//        var splits = tableList[cnt].split("*");
//       ReportUtils.logStep("INFO", tableList[cnt]);
//        Log.Message(tableList[cnt]);
      }
    }
    
    var compStatus = false;
    for(var tl=0;tl<TemplateCode.length;tl++){
//    var TempNo = TemplateCode[tl].substring(TemplateCode[tl].indexOf("_")+1)
      if(TemplateCode[tl].indexOf(value.toString().trim())!=-1){
//      Log.Message("TemplateCode[tl] :"+TemplateCode[tl]);
//      Log.Message("value :"+value);
//      value = TemplateCode[tl];
        compStatus = true;
        break;
        }
      }
  if(!compStatus){
    ValidationUtils.verify(false,true,"Given TemplateNo in Datasheet is not available in ConfigPack");
    }
    }
  //==============================================
  var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  waitForObj(code);
//  code.Keys("[Tab]");
//  aqUtils.Delay(2000, Indicator.Text);;
//  var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText("*"+value.toString().trim()+"*");
//  aqUtils.Delay(3000, Indicator.Text);;
  var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  waitForObj(serch);
  serch.Click();
//   if(serch.isEnabled())
//  serch.Click();
//  else{ 
//    aqUtils.Delay(3000, Indicator.Text);;
//   serch.Click(); 
//  }
//  aqUtils.Delay(4000, Indicator.Text);;
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  waitForObj(OK);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim().indexOf(value.toString().trim())!=-1){ 
     var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
     waitForObj(OK);
     OK.Click();
//      if(OK.isEnabled())
//  OK.Click();
//  else{ 
//    aqUtils.Delay(3000, Indicator.Text);;
//   OK.Click(); 
//  }
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
         break;
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
     waitForObj(cancel);
     cancel.Click();
//        if(cancel.isEnabled())
//  cancel.Click();
//  else{ 
//    aqUtils.Delay(3000, Indicator.Text);;
//   cancel.Click(); 
//  }
//        aqUtils.Delay(1000, Indicator.Text);;
Sys.HighlightObject(Obj_Address);
        Obj_Address.setText("");
        ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
    waitForObj(cancel);
     cancel.Click();
//    if(cancel.isEnabled())
//  cancel.Click();
//  else{ 
//    aqUtils.Delay(3000, Indicator.Text);;
//   cancel.Click(); 
//  }
//    aqUtils.Delay(1000, Indicator.Text);;
    ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    Sys.HighlightObject(Obj_Address);
    Obj_Address.setText("");
  }
        }
else{ 
    ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }   
    
}


function currency(baseCurrency){ 
var NewCurrency = "";
switch(baseCurrency.toLowerCase()) { // if we need to match case sensitive put Uppercase with in switch "baseCurrency.toUpperCase()"
case "indian rupee":{
NewCurrency = "INR"
}
break;

case "chinese yuan renminbi":{
NewCurrency = "CNY"
}
break;

case "us dollar":{
NewCurrency = "USD"
}
break;

case "euro":{
NewCurrency = "EUR"
}
break;

default:{
NewCurrency = ""; 
}
}
return NewCurrency;
}



function SearchByValueTable(ObjectAddrs,popupName,value,fieldName){
var checkmark =  false;
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(1).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
          OK.Click();
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
          cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}


function SearchByValueasset(ObjectAddrs,popupName,value,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, Indicator.Text);
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Transaction Type").OleValue.toString().trim()).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    Sys.HighlightObject(serch);
    if(serch.isEnabled())
  serch.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);
   serch.Click(); 
  }
    aqUtils.Delay(5000, Indicator.Text);
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  if(OK.isEnabled()){
  OK.HoverMouse();
ReportUtils.logStep_Screenshot();
  OK.Click();
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);
    OK.HoverMouse();
ReportUtils.logStep_Screenshot();
   OK.Click(); 
  }
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
if(cancel.isEnabled()){
  cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
  cancel.Click();
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);
      cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
   cancel.Click(); 
  }
          aqUtils.Delay(1000, Indicator.Text);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
if(cancel.isEnabled()){
    cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
  cancel.Click();
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);
      cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
   cancel.Click(); 
  }
      aqUtils.Delay(1000, Indicator.Text);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}

function SearchByValueTableComp(ObjectAddrs,popupName,value,fieldName){
var checkmark =  false;
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim());
          OK.Click();
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
          cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}



function StartwaitTime(){ 
var dif;
var TodayValue = aqDateTime.Today();
var StringTodayValue = aqConvert.DateTimeToStr(TodayValue);
var EncodedDate = aqConvert.DateTimeToFormatStr(StringTodayValue,"%d%#B%Y"); 
var STIME = getFormattedCurrentTime()
Log.Message("Start DATE & TIME for Object Address :"+EncodedDate +" "+STIME)
var start = STIME.split(":");
if(start[1]>0){ 
dif = Number(start[2]) + Number(start[1]*60);
}
if(start[0]>0){ 
dif = dif + Number(start[0]*60*60);
}
return dif;
}

function EndTime(){ 
var dif2;
TodayValue = aqDateTime.Today();
StringTodayValue = aqConvert.DateTimeToStr(TodayValue);
EncodedDate = aqConvert.DateTimeToFormatStr(StringTodayValue,"%d%#B%Y"); 
var ETIME =getFormattedCurrentTime()
Log.Message("End DATE & TIME :"+EncodedDate +" "+ETIME); 
var end = ETIME.split(":");
if(end[1]>0){ 
dif2 = Number( end[2]) + Number(end[1]*60);
}
if(end[0]>0){ 
dif2 = dif2 + Number(end[0]*60*60);
}
return dif2;
}  
//function getFormattedCurrentTime(){
//  TodayValue = aqConvert.DateTimeToFormatStr(aqDateTime.Time(), "%H:%M:%S");
//  return TodayValue;
//}

function waitUntilInvisibleOfObj(ObjAdd)
{
  if (ObjAdd.WaitProperty("VisibleOnScreen", false,400000)){ 
}
else
Log.Error("Object is Visible/Enabled")

}
function waitForObj(ObjAdd){  
  
if (ObjAdd.WaitProperty("Enabled", true, 20000)){ 
 Sys.HighlightObject(ObjAdd);
ObjAdd.HoverMouse(); 
}
else
Log.Error("Object is not Visible/Enabled")
}

/*
 Get Date in MM/DD/YYYY format. 
 Parameter daysFromToday:  0 returns todays date
*/
function getSpecificDate(daysFromToday)
{
  var date = new Date();
  date.setDate(date.getDate() + daysFromToday);
  var dd = date.getDate();
  var mm = date.getMonth()+1; 
  var yyyy = date.getFullYear();
  if(dd<10) {
      dd='0'+dd;
  } 
  if(mm<10) {
      mm='0'+mm;
  } 
  if((EnvParams.Country.toUpperCase()=="INDIA") || (EnvParams.Country.toUpperCase()=="MALAYSIA"))
    date = dd+'-'+mm+'-'+yyyy;
  else
    date = mm+'/'+dd+'/'+yyyy;
  Log.Message(date);
  return date;
}

function getDateFormat(daysFromToday)
{
  var date = new Date();
  date.setDate(date.getDate() + daysFromToday);
  var dd = date.getDate();
  var mm = date.getMonth()+1; 
  var yyyy = date.getFullYear();
  if(dd<10) {
      dd='0'+dd;
  } 
  if(mm<10) {
      mm='0'+mm;
  } 
  date = dd+'/'+mm+'/'+yyyy;
  Log.Message(date);
  return date;
}

function levelMatch(Approve_Level){
var list_A = [];
var list_B = [];
var list_C = [];
var list_D = [];
  		for(var i=0;i<Approve_Level.length;i++){
			var temp = Approve_Level[i].split("*");
			if(i==0){
				if(temp.length>3){
				list_A[i] = temp[0]+"*"+temp[1];
				list_C[i] = temp[0]+"*"+temp[1];
				list_B[i] = temp[2]+"*"+temp[3];
				list_D[i] = temp[2]+"*"+temp[3];
        if(Approve_Level.length==1)
        return list_A;
				}
				else{ 
					list_A[i] = temp[0]+"*"+temp[1];
					list_B[i] = temp[0]+"*"+temp[1];
					list_C[i] = temp[0]+"*"+temp[1];
					list_D[i] = temp[0]+"*"+temp[1];
          if(Approve_Level.length==1)
          return list_A;
				}
			}
      
     	if(i==1){
				var temp1 = list_A[0].toString().split("*");	
				if(!(temp1[0]==temp[0])){
					list_A[1] = temp[0]+"*"+temp[1];
          if(Approve_Level.length==2)
          return list_A;
				}
				temp1 = list_C[0].toString().split("*");
				if(temp.length>3)
				if(!(temp1[0]==temp[2])){
					list_C[1] = temp[2]+"*"+temp[3];
          if(Approve_Level.length==2)
          return list_C;
				}
				temp1 = list_D[0].toString().split("*");	
				if(!(temp1[0]==temp[0])){
					list_D[1] = temp[0]+"*"+temp[1];
          if(Approve_Level.length==2)
          return list_D;
				}
				temp1 = list_B[0].toString().split("*");
				if(temp.length>3)
				if(!(temp1[0]==temp[2])){
					list_B[1] = temp[2]+"*"+temp[3];
          if(Approve_Level.length==2)
          return list_B;
				}
			} 
      
      
      			if(i==2){
	//List A
				
				if(list_A.length==2){
				Log.Message("List A");
				var sts = true;
        for(var z=0;z<list_A.length;z++){
          var temp1 = list_A[z].toString().split("*");
					if(temp1[0]==temp[0]){
						sts = false;
						break;
					}
        }
        
				if(sts){
				list_A[2] = temp[0]+"*"+temp[1];
				}
				else{
				if(temp.length>3){ 
        sts = true;
        for(var z=0;z<list_A.length;z++){
          var temp1 = list_A[z].toString().split("*");
					if(temp1[0]==temp[2]){
						sts = false;
						break;
					}
        }
        
					if(sts){
					list_A[2] = temp[2]+"*"+temp[3];
          Log.Message(temp[2]+"*"+temp[3])
					}
				}
				}

				if(list_A.length==3){
					for(var z=0;z<list_A.length;z++)
						Log.Message(list_A[z]);
					return list_A;
				}
				
				}
        
 //LIST B       
    if(list_B.length==2){
				Log.Message("List B");
				var sts = true;
        for(var z=0;z<list_B.length;z++){
          var temp1 = list_B[z].toString().split("*");
					if(temp1[0]==temp[0]){
						sts = false;
						break;
					}
        }
        
				if(sts){
				list_B[2] = temp[0]+"*"+temp[1];
				}
				else{
				if(temp.length>3){ 
					sts = true;
        for(var z=0;z<list_B.length;z++){
          var temp1 = list_B[z].toString().split("*");
					if(temp1[0]==temp[2]){
						sts = false;
						break;
					}
        }
        
					if(sts)
						list_B[2] = temp[2]+"*"+temp[3];	
				}
				}
				
				if(list_B.length==3){
					for(var z=0;z<list_B.length;z++)
						Log.Message(list_B[z]);
					return list_B;
				}
				
				}
//List C 
    if(list_C.length==2){
				Log.Message("List C");
					var sts = true;
        for(var z=0;z<list_C.length;z++){
          var temp1 = list_C[z].toString().split("*");
					if(temp1[0]==temp[0]){
						sts = false;
						break;
					}
        }
        

				if(sts){
					list_C[2] = temp[0]+"*"+temp[1];
				}
				else{
				if(temp.length>3){ 
					sts = true;
        for(var z=0;z<list_C.length;z++){
          var temp1 = list_C[z].toString().split("*");
					if(temp1[0]==temp[2]){
						sts = false;
						break;
					}
        }

					if(sts)
						list_C[2] = temp[2]+"*"+temp[3];	
				}
				}
				
				if(list_C.length==3){
					for(var z=0;z<list_C.length;z++)
						Log.Message(list_C[z]);
					return list_C;
				}
				}
        
//List D
				if(list_D.length==2){
				Log.Message("List D");
				var sts = true;
        for(var z=0;z<list_D.length;z++){
          var temp1 = list_D[z].toString().split("*");
					if(temp1[0]==temp[0]){
						sts = false;
						break;
					}
        }
        
				if(sts){
					list_D[2] = temp[0]+"*"+temp[1];
				}
				else{
				if(temp.length>3){ 
					sts = true;
        for(var z=0;z<list_D.length;z++){
          var temp1 = list_D[z].toString().split("*");
					if(temp1[0]==temp[2]){
						sts = false;
						break;
					}
        }
        
					if(sts)
						list_D[2] = temp[2]+"*"+temp[3];	
				}
				}
				
				if(list_D.length==3){
					for(var z=0;z<list_D.length;z++)
						Log.Message(list_D[z]);
					return list_D;
				}
				
				}
        
      }
      
      } 

}









function VacationPeriod(ObjectAddrs,popupName,value,Vacation_Calander,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, popupName);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);

aqUtils.Delay(5000, popupName);;
    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
  waitForObj(code);
  code.Click();

    code.setText(Vacation_Calander);
    aqUtils.Delay(3000, popupName);;
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(3000, popupName);;
    
    var period = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
    Sys.HighlightObject(period);
    period.Click();
    period.setText(value);
    
    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    waitForObj(serch);

  serch.Click();
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())

    waitForObj(OK);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){
    for(var i=0;i<itemCount;i++){
      if((table.getItem(i).getText_2(0).OleValue.toString().trim()==Vacation_Calander) && (table.getItem(i).getText_2(1).OleValue.toString().trim()==value)){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  waitForObj(OK);
  OK.Click();

          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
  waitForObj(cancel);
  cancel.Click();

          Sys.HighlightObject(ObjectAddrs);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        waitForObj(cancel);
        cancel.Click();

      Sys.HighlightObject(ObjectAddrs);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    
    return checkmark;
}



function SearchByValue_Emp(ObjectAddrs,popupName,value,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, popupName);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);


    var code = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  waitForObj(code);
  code.Click();

    code.setText(value);

    var serch = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    waitForObj(serch);

  serch.Click();
  var table = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
  var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())

    waitForObj(OK);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  waitForObj(OK);
  OK.Click();

          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
          break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
  waitForObj(cancel);
  cancel.Click();

          Sys.HighlightObject(ObjectAddrs);
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = eval(Sys_Maconomy_Parent).SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        waitForObj(cancel);
        cancel.Click();

      Sys.HighlightObject(ObjectAddrs);
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    
    return checkmark;
}



function savePDF_localDirectory(FileName,pdfname){ 


var SaveTitle = "";
var sFolder = "";
var pdf = Sys.WaitProcess("AcroRd32", 2,120000).Window("AcrobatSDIWindow", "*"+".pdf - Adobe Acrobat Reader DC*", 1).Window("AVL_AVView", "AVFlipContainerView", 2).Window("AVL_AVView", "AVDocumentMainView", 1).Window("AVL_AVView", "AVFlipContainerView", 3).Window("AVL_AVView", "AVSplitterView", 3).Window("AVL_AVView", "AVSplitationPageView", 3).Window("AVL_AVView", "AVSplitterView", 1).Window("AVL_AVView", "AVScrolledPageView", 1).Window("AVL_AVView", "AVScrollView", 1).Window("AVL_AVView", "AVPageView", 5);
   if(Sys.Process("AcroRd32", 2).Window("AcrobatSDIWindow", "*"+".pdf - Adobe Acrobat Reader DC*", 1).WndCaption.indexOf(pdfname)!=-1){
    aqUtils.Delay(6000, Indicator.Text);

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

var filepathforMplValidation =sFolder+SaveTitle+".pdf";

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
aqUtils.Delay(2000, Indicator.Text);
Sys.HighlightObject(pdf);
Sys.Desktop.KeyDown(0x12); //Alt
Sys.Desktop.KeyDown(0x46); //F
Sys.Desktop.KeyDown(0x58); //X 
Sys.Desktop.KeyUp(0x46); //Alt
Sys.Desktop.KeyUp(0x12);     
Sys.Desktop.KeyUp(0x58);
}
ValidationUtils.verify(true,true,"Print "+FileName+" is Clicked and PDF is Saved");
Log.Message("PDF saved location : "+sFolder+SaveTitle+".pdf")
ReportUtils.logStep("INFO","PDF saved location : "+sFolder+SaveTitle+".pdf");

ExcelUtils.setExcelName(workBook,"Data Management", true);
ExcelUtils.WriteExcelSheet(FileName,EnvParams.Opco,"Data Management",sFolder+SaveTitle+".pdf")

return sFolder+SaveTitle+".pdf";
}


function switch_Maconomy(UserName){
  

var Maconomy_ParentAddress="";
var screenFound = false;
for(var mac_Count=0;mac_Count<Sys.ChildCount;mac_Count++){ 
  if(Sys.Child(mac_Count).ProcessName=="Maconomy"){
    for(var Shell_Count=0;Shell_Count<Sys.Child(mac_Count).ChildCount;Shell_Count++){ 
      if((Sys.Child(mac_Count).Child(Shell_Count).Exists)&&(Sys.Child(mac_Count).Child(Shell_Count).IsOpen)&&(Sys.Child(mac_Count).Child(Shell_Count).Name.indexOf("SWTObject")!=-1)&&(Sys.Child(mac_Count).Child(Shell_Count).JavaClassName=="Shell")
      &&(Sys.Child(mac_Count).Child(Shell_Count).WndCaption.toString().trim().indexOf(UserName)!=-1)){ 
        Maconomy_ParentAddress = Sys.Child(mac_Count).Child(Shell_Count).FullName;
        Log.Message(Maconomy_ParentAddress)
        Sys_Maconomy_Parent = Sys.Child(mac_Count);
        Log.Message(Sys_Maconomy_Parent)
        screenFound = true;
        break;
      }
    }
  }

}

if(screenFound)
return Maconomy_ParentAddress;
else
return "Screen Not Found";

}



function Maconomy_index(){
  

var Maconomy_ParentAddress="";
var screenFound = false;
var Index = 0;
for(var mac_Count=0;mac_Count<Sys.ChildCount;mac_Count++){ 
  if(Sys.Child(mac_Count).ProcessName=="Maconomy"){
    Index++;
    screenFound = true;


  }

}

Log.Message(Index)
if(screenFound)
return Index;
else
return "Screen Not Found";

}


//Finding UserName for Approvers in Datasheets
function CredentialLogin(Approve_Level,excelName){ 
  var AppvLevl = [];
  var ApproveInfo = [];
for(var i=0;i<Approve_Level.length;i++){
  var UserN = true;
  var temp="";
  var temp1="";
  var Cred = Approve_Level[i].split("*");
  for(var j=2;j<4;j++){
  temp="";
  if((Cred[j]!="")&&(Cred[j]!=null))
  if((Cred[j].indexOf("IND")==-1)&&(Cred[j].indexOf("SPA")==-1)&&(Cred[j].indexOf("SGP")==-1)&&(Cred[j].indexOf("MYS")==-1)&&(Cred[j].indexOf("FP")==-1)&&(Cred[j].indexOf("CHFP")==-1)&&(Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf(EnvParams.Opco+" ")!=-1)))
  { 
     var sheetName = "Agency Users";
     workBook = Project.Path+excelName;
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.AgencyLogin(Cred[j],EnvParams.Opco);
  }
  else if((Cred[j].indexOf("IND")!=-1)||(Cred[j].indexOf("SPA")!=-1)||(Cred[j].indexOf("SGP")!=-1)||(Cred[j].indexOf("MYS")!=-1)||(Cred[j].indexOf("FP")!=-1)||(Cred[j].indexOf("CHFP")!=-1)||(Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
  { 

    var sheetName = "SSC Users";
    ExcelUtils.setExcelName(workBook, sheetName, true);
    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
  }

  if(temp.length!=0){
    temp1 = temp1+temp+"*"+j+"*";
//  break;
  }
  }
  if((temp1=="")||(temp1==null))
  Log.Error("User Name is Not available for level :"+i);
  Log.Message(temp1)
  AppvLevl[i] = temp1;
}
  ApproveInfo = levelMatch(AppvLevl)
  Log.Message("-----Approvers-------------")
  for(var i=0;i<ApproveInfo.length;i++){
    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+ApproveInfo[i];
    Log.Message(ApproveInfo[i]);
    }

    return ApproveInfo;
}

