//USEUNIT ExcelUtils
//USEUNIT EnvParams
//USEUNIT ValidationUtils
/*Closes workspaces after job completes in maconomy*/

var Language = "";

function closeAllWorkspaces(){
  Sys.Desktop.KeyDown(0x12); //Alt
  Sys.Desktop.KeyDown(0x57); //W
  Sys.Desktop.KeyDown(0x0D); //Enter
  Sys.Desktop.KeyUp(0x12); 
  Sys.Desktop.KeyUp(0x57);
  Sys.Desktop.KeyUp(0x0D);
}

function closeMaconomy(){ 
var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
//  Log.Message("Maconomy is Already in Running")
    Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x58); //X 
    Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
    Sys.Desktop.KeyUp(0x58);
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

function VPWSearchByValue(ObjectAddrs,popupName,value,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);;
//    Log.Message(ObjectAddrs)
//    Log.Message(popupName)
//    Log.Message(value)
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "")
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    Sys.HighlightObject(serch);
    if(serch.isEnabled())
  serch.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
   serch.Click(); 
  }
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2)
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  if(OK.isEnabled()){
  OK.HoverMouse();
ReportUtils.logStep_Screenshot();
  OK.Click();
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
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
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
if(cancel.isEnabled()){
  cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
  cancel.Click();
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
      cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
   cancel.Click(); 
  }
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
if(cancel.isEnabled()){
    cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
  cancel.Click();
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
      cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
   cancel.Click(); 
  }
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}


function SearchByValue(ObjectAddrs,popupName,value,fieldName){ 
var checkmark = false;
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);;
//    Log.Message(ObjectAddrs)
//    Log.Message(popupName)
//    Log.Message(value)
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    Sys.HighlightObject(serch);
    if(serch.isEnabled())
  serch.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
   serch.Click(); 
  }
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  if(OK.isEnabled()){
  OK.HoverMouse();
ReportUtils.logStep_Screenshot();
  OK.Click();
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
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
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
if(cancel.isEnabled()){
  cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
  cancel.Click();
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
      cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
   cancel.Click(); 
  }
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
if(cancel.isEnabled()){
    cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
  cancel.Click();
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
      cancel.HoverMouse();
ReportUtils.logStep_Screenshot();
   cancel.Click(); 
  }
      aqUtils.Delay(1000, Indicator.Text);;
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
    aqUtils.Delay(3000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    var alljob = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Jobs");
    alljob.Click();
    aqUtils.Delay(2000, Indicator.Text);; 
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(1).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
         break;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
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
    var alljob = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Jobs");
    alljob.Click();
    aqUtils.Delay(2000, Indicator.Text);; 
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
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
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         break;
        ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
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
//    Sys.Desktop.KeyDown(0x09);
//    Sys.Desktop.KeyUp(0x09);
//    aqUtils.Delay(1000, Indicator.Text);;
//    var alljob = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Jobs");
//    alljob.Click();
//    aqUtils.Delay(2000, Indicator.Text);; 
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy"); 
         break;
         
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}


function AccessLevel_Add(ObjectAddrs,popupName,value,fieldName){

var checkmark = false; 
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);;
//    Sys.Desktop.KeyDown(0x09);
//    Sys.Desktop.KeyUp(0x09);
//    aqUtils.Delay(1000, Indicator.Text);;
//    var alljob = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Jobs");
//    alljob.Click();
//    aqUtils.Delay(2000, Indicator.Text);; 
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy"); 
         break;
         
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}


function SearchByValues_all_Col_1(ObjectAddrs,popupName,value,fieldName){

var checkmark = false; 
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);;
//    Sys.Desktop.KeyDown(0x09);
//    Sys.Desktop.KeyUp(0x09);
//    aqUtils.Delay(1000, Indicator.Text);;
    var alljob = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Jobs");
    alljob.Click();
    aqUtils.Delay(2000, Indicator.Text);; 
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    Sys.HighlightObject(serch);
    if(serch.isEnabled())
  serch.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
   serch.Click(); 
  }
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         if(OK.isEnabled())
  OK.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
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
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
if(cancel.isEnabled())
  cancel.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
   cancel.Click(); 
  }
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
if(cancel.isEnabled())
  cancel.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
   cancel.Click(); 
  }
      aqUtils.Delay(1000, Indicator.Text);;
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
//    Sys.Desktop.KeyDown(0x09);
//    Sys.Desktop.KeyUp(0x09);
//    aqUtils.Delay(1000, Indicator.Text);;
//    var alljob = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Jobs");
//    alljob.Click();
//    aqUtils.Delay(2000, Indicator.Text);; 
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         break;
        ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
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
    aqUtils.Delay(3000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
//    var alljob = Sys.Process("Maconomy").SWTObject("Shell", "Job").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Jobs");
//    alljob.Click();
//    aqUtils.Delay(2000, Indicator.Text);; 
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(1).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         break;
        ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}








function SearchByValues_all_Col_2(ObjectAddrs,popupName,value,fieldName,all){

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
    var alljob = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", all);
    alljob.Click();
    aqUtils.Delay(2000, Indicator.Text);; 
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(1).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
         break;
        
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}

function SearchByValues_Col_1_all(ObjectAddrs,popupName,value,fieldName,all){

var checkmark = false; 
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);;
//    Sys.Desktop.KeyDown(0x09);
//    Sys.Desktop.KeyUp(0x09);
//    aqUtils.Delay(1000, Indicator.Text);;
    var alljob = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", all);
    alljob.Click();
    aqUtils.Delay(2000, Indicator.Text);; 
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
//    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ")
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
   Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
         OK.Click();
         checkmark = true;
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
         break;
        
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
           cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
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
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
          OK.Click();
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
          cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
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
  aqUtils.Delay(1000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x11);
    Sys.Desktop.KeyDown(0x47);
    Sys.Desktop.KeyUp(0x11);
    Sys.Desktop.KeyUp(0x47);
    aqUtils.Delay(3000, Indicator.Text);;
    Sys.Desktop.KeyDown(0x09);
    Sys.Desktop.KeyUp(0x09);
    aqUtils.Delay(1000, Indicator.Text);;
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(1).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
          OK.Click();
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
          cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
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
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
          OK.Click();
          checkmark = true;
          
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
          cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
    }
    return checkmark;
}






function CalenderDateSelection(ObjectAddrs,value){ 
    var temp = "";
  temp = value.split("/");
  
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





//function Login_Match(Approve_Level,LoginEmp,HRData,comID){ 
//login_satuts = true;
//aqUtils.Delay(3000, Indicator.Text);;
//var UserPasswd = [];
//var z =0;
//for(var i=0;i<Approve_Level.length;i++){ 
//if((Approve_Level[i].indexOf("OpCo")!=-1) && (comID!="")){
//Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,comID);
//}
//// Approve_Level[i] = Approve_Level[i].replace(/OpCo/g,"1710");  //GCD2_Company No- level[0]
//if(Approve_Level[i].indexOf("SSC - Biller")==-1){
//Approve_Level[i] = Approve_Level[i].replace(/- Billers/g,"- Agency - Biller");
//}
//
//var tempLevel = Approve_Level[i].split("*");
//ifGotIT = true;
//for(var j=2;j<tempLevel.length;j++){ 
//
//if((tempLevel[j].indexOf(" (")!=-1) && (tempLevel[j].indexOf(")")!=-1)){
//var temp = tempLevel[j].replace(" (","*");
//temp = temp.replace(")","");
////Log.Message("temp :"+temp)
//var tempSplit = temp.split("*");
//
//for(var k=0;k<LoginEmp.length;k++){
//  var A_temp = LoginEmp[k].split("*");
////    Log.Message("tempSplit[0] :"+tempSplit[0]);
////    Log.Message("A_temp[0] :"+A_temp[0]);
////    Log.Message("tempSplit[1] :"+tempSplit[1]);
////    Log.Message("A_temp[1] :"+A_temp[1]);
// if(tempSplit[0]==A_temp[0]){ 
//    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
////     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[2];
//   Log.Message(UserPasswd[z]);
//   z++;
//   ifGotIT = false;
//   break;     
// }else{ 
// if(tempSplit[1]==A_temp[2]){ 
//    UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
////     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[2];
//   Log.Message(UserPasswd[z]);
//   z++;
//   ifGotIT = false;
//   break;     
// }     
// }
//    
//}
//if(!ifGotIT){ 
//  break;
//}
//}
//  
//if((tempLevel[j].indexOf("SSC -")!=-1) || (tempLevel[j].indexOf("Central Team -")!=-1)){ 
////    Log.Message("tempLevel[j] :"+tempLevel[j]);
//   if(tempLevel[j].indexOf("Central Team - Client Management")!=-1){ 
//    temp2 = "Central Team - Client Account Management";
//  }
//  else if(tempLevel[j].indexOf("Central Team - Vendor Management")!=-1){ 
//    temp2 = "Central Team - Vendor Account Management";
//  }
//  else if(tempLevel[j].indexOf("SSC - Expense Cashiers")!=-1){ 
//    temp2 = "SSC - Cashier";
//  }else{ 
//    temp2 = tempLevel[j];
//  }
//for(var k=0;k<LoginEmp.length;k++){
//  var A_temp = LoginEmp[k].split("*");  
// if(temp2==A_temp[1]){ 
//   UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
////     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3];
//     
//   Log.Message(UserPasswd[z]);
//   z++;
//   ifGotIT = false;
//   break;     
// }
//}  
//
//if(!ifGotIT){ 
//  break;
//}
//}
//  
//  
//  
//if((tempLevel[j].indexOf(" (")==-1) && (tempLevel[j].indexOf(")")==-1) && 
//(tempLevel[j].indexOf("SSC -")==-1) && (tempLevel[j].indexOf("Central Team -")==-1)){ 
//    
//for(var k=0;k<LoginEmp.length;k++){
//  var A_temp = LoginEmp[k].split("*");
//  if(A_temp[0]==tempLevel[j]){  // Better  to use level[j].indexOf(LoginArrays[k])
//  UserPasswd[z] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+A_temp[3]; 
//  Log.Message(UserPasswd[z]);
//   z++;
//   ifGotIT = false;
//   break;     
// }
//}
//if(!ifGotIT){ 
//  break;
//}
//}
//  
//if((tempLevel[j].indexOf(" (")!=-1) && (tempLevel[j].indexOf(")")!=-1)){
//
//var temp = tempLevel[j].replace(" (","*");
//temp = temp.replace(")","");
////Log.Message("temp :"+temp)
//var tempSplit = temp.split("*");
//
//for(var k=0;k<HRData.length;k++){
//  var A_temp = HRData[k].split("*");
////    Log.Message("tempSplit[0] :"+tempSplit[0]);
////    Log.Message("A_temp[0] :"+A_temp[0]);
////    Log.Message("tempSplit[1] :"+tempSplit[1]);
////    Log.Message("A_temp[1] :"+A_temp[1]);
// if(tempSplit[1]==A_temp[1]){ 
//   UserPasswd[z]  = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123";
////     Approve_Level[i] = tempLevel[0]+"*"+tempLevel[1]+"*"+A_temp[0]+"*"+"CORE@WPP123";
//   Log.Message(UserPasswd[z]);
//   z++;
//   ifGotIT = false;
//   break;     
// }
//    
//}
//if(!ifGotIT){ 
//  break;
//}
//}
// 
//  
//}
//if(ifGotIT){ 
//  Log.Warning("UserName and Password is Not Matched for Approver and Substitute :"+Approve_Level[i]);
//  login_satuts = false;
//  break;
//}
//  
//}
//
//return UserPasswd;
//}



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





function DropDownList(value,feild){ 
var checkMark = false;
Sys.Process("Maconomy").Refresh();
  var list = Sys.Process("Maconomy").SWTObject("Shell", "").SWTObject("ScrolledComposite", "").SWTObject("McValuePickerPanel", "").WaitSWTObject("Grid", "", 3,60000); 
  var Add_Visible4 = true;
  while(Add_Visible4){
  if(list.isEnabled()){
  Add_Visible4 = false;
      for(var i=0;i<list.getItemCount();i++){ 
        if(list.getItem(i).getText_2(0)!=null){ 
          if(list.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
            list.Keys("[Enter]");
            aqUtils.Delay(5000, Indicator.Text);;
            checkMark = true;
            ValidationUtils.verify(true,true,feild+" is selected in Maconomy");
            break;
          }else{
//          Log.Message("i :"+i);
//          Log.Message(value+" "+value.length);
          
//        Log.Message(list.getItem(i).getText_2(0).OleValue.toString().trim()+" "+list.getItem(i).getText_2(0).OleValue.toString().trim().length); 
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





//function goToHR(){ 
//var HRData = [];
//aqUtils.Delay(3000, Indicator.Text);;
//  closeAllWorkspaces();
//  aqUtils.Delay(1000, Indicator.Text);
//var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
//    menuBar.DblClick();
//
//if(ImageRepository.ImageSet.HR.Exists()){
//ImageRepository.ImageSet.HR.Click();
//}
//else if(ImageRepository.ImageSet.HR1.Exists()){
//ImageRepository.ImageSet.HR1.Click();
//}
//else if(ImageRepository.ImageSet.HR2.Exists()){
//ImageRepository.ImageSet.HR2.Click();  
//}
////if(ImageRepository.ImageSet.User1.Exists()){
////  ImageRepository.ImageSet.User1.DblClick();// GL
////}
////else if(ImageRepository.ImageSet.User3.Exists()){
////  ImageRepository.ImageSet.User3.DblClick();// GL
////}
////else if(ImageRepository.ImageSet.User2.Exists()){
////  ImageRepository.ImageSet.User2.DblClick();// GL
////}
//
//
//var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
//  var HRitem;
//Log.Message(childCC)
//for(var i=1;i<=childCC;i++){ 
//HRitem = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
//if(HRitem.isVisible()){ 
//HRitem = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
//HRitem.DblClickItem("|Users");
//}
//
//}
//
////var HRitem = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", 5).SWTObject("Tree", "");
////HRitem.DblClickItem("|Users");
//aqUtils.Delay(5000, Indicator.Text);;
////var ActiveUser = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "Active Users");
////ActiveUser.Click();
//var All_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
//All_User.Click();
//aqUtils.Delay(5000, Indicator.Text);;
//var HRTable = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
//var z=0;
//for(var i=0;i<HRTable.getItemCount();i++){ 
//if(HRTable.getItem(i).getText(2)!=""){
//HRData[z] = HRTable.getItem(i).getText_2(0).OleValue.toString().trim()+"*"+HRTable.getItem(i).getText_2(2).OleValue.toString().trim()
////Log.Message(HRData[z]);
//z++;
//
//}
//}
//return HRData;
//}



function goToHR(){ 
Delay(3000);
//  closeAllWorkspaces();

var menuBar = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4)
menuBar.DblClick();
if(ImageRepository.ImageSet.HR.Exists()){
ImageRepository.ImageSet.HR.Click();
}
else if(ImageRepository.ImageSet.HR1.Exists()){
ImageRepository.ImageSet.HR1.Click();
}
else if(ImageRepository.ImageSet.HR2.Exists()){
ImageRepository.ImageSet.HR2.Click();  
}

var mainlist = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").ChildCount;
var main;
for(var id=0;id<mainlist;id++){
main = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
if(main.Child(id).isVisible())
if(main.Child(id).ChildCount==1)
if(main.Child(id).Child(0).Name.indexOf("Composite")!=-1){

var childCC= Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").ChildCount;
  var Client_Managt;
//Log.Message(childCC)
for(var i=1;i<=childCC;i++){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i)
if(Client_Managt.isVisible()){ 
Client_Managt = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("McMaconomyPShelfMenuGui$3", "", 2).SWTObject("PShelf", "").SWTObject("Composite", "", i).SWTObject("Tree", "");
Client_Managt.DblClickItem("|Users");
}
}
}

}


//HRitem.DblClickItem("|Users");
Delay(5000);
var ActiveUser = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.McFilterContainer.Composite.McFilterPanelWidget.Button;
ActiveUser.Click();
//var All_User = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McFilterContainer", "", 1).SWTObject("Composite", "").SWTObject("McFilterPanelWidget", "").SWTObject("Button", "All Users");
//All_User.Click();
Delay(5000);
}

function searchNumber(Eno){
  var temp = "";
Delay(2000);
var HRTable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite2.McClumpSashForm.POApproverList.McWorkspaceSheafGui_McDecoratedPaneGui.Composite.Composite.McFilterPaneWidget.JobsTable.McGrid;
var firstCell = HRTable.SWTObject("McTextWidget", "");
firstCell.Click();
Delay(1000);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
Delay(1000);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x09);
Delay(1000);    
var EmpNumber = HRTable.SWTObject("McValuePickerWidget", "", 2);
EmpNumber.setText("^a[BS]");
EmpNumber.setText(Eno);
Delay(5000);
var z=0;
for(var i=0;i<HRTable.getItemCount();i++){ 
if(HRTable.getItem(i).getText(2).OleValue.toString().trim()==Eno){
temp = HRTable.getItem(i).getText_2(0).OleValue.toString().trim();
//Log.Message(temp);
z++;

}
}
Delay(1000);
Sys.Desktop.KeyDown(0x10);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x10);
Sys.Desktop.KeyUp(0x09);
Delay(1000);
Sys.Desktop.KeyDown(0x10);
Sys.Desktop.KeyDown(0x09);
Sys.Desktop.KeyUp(0x10);
Sys.Desktop.KeyUp(0x09);
Delay(1000); 
return temp;
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


function Rests(uname,pwd){ 
aqUtils.Delay(5000, Indicator.Text);;
      Sys.Desktop.KeyDown(0x12); //Alt
    Sys.Desktop.KeyDown(0x46); //F
    Sys.Desktop.KeyDown(0x52); //R 
     Sys.Desktop.KeyUp(0x46); //Alt
    Sys.Desktop.KeyUp(0x12);     
     Sys.Desktop.KeyUp(0x52); //R
aqUtils.Delay(65000, Indicator.Text);;
     var usernameAddr = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 1);
    var pwdAddr = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 2);
    var btnLogin = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Login");
    usernameAddr.SetFocus();
    usernameAddr.setText(uname);
    pwdAddr.setText(pwd);
    btnLogin.click();
    
  /*  
    aqUtils.Delay(7000, Indicator.Text);;
    var status = true;
    while(status){
    var Name = Sys.Process("Maconomy").SWTObject("Shell", "*").WndCaption.toString().trim();
    if(Name=="Login to Deltek Maconomy"){ 
      status = false;
    usernameAddr = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 1);
    pwdAddr = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 2);
    btnLogin = Sys.Process("Maconomy").SWTObject("Shell", "Login to Deltek Maconomy").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Login");

      aqUtils.Delay(2000, Indicator.Text);;
//      Login_to_Deltek_Maconomy();
      usernameAddr.SetFocus();
      usernameAddr.setText(loginuser);
      pwdAddr.setText(loginpassword);
      btnLogin.click();
      aqUtils.Delay(10000, Indicator.Text);;
      break;
    }
    if(Name=="Server Configuration"){ 
      aqUtils.Delay(2000, Indicator.Text);;
    server_link = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Text", "");
    port_number = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 2)
    company_name = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Text", "", 3)
    chk_box = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Button", "Do not ask me again");
    connectbtn = Sys.Process("Maconomy").SWTObject("Shell", "Server Configuration").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Connect");

//      Server_Configuration();
      server_link.SetFocus();
      server_link.setText(server);
      port_number.SetFocus();
      port_number.setText(port);
      company_name.SetFocus();
      company_name.setText(company);
    
      if(!chk_box.getSelection()){
      chk_box.ClickButton(cbChecked);
      }
      connectbtn.click();
      aqUtils.Delay(5000, Indicator.Text);;
    }
    }
    aqUtils.Delay(10000, Indicator.Text);; 
    
    */  
}


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
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
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
    var tab = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
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
  
  var code = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(comapany.toString().trim());
  aqUtils.Delay(3000, Indicator.Text);;
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  aqUtils.Delay(5000, Indicator.Text);;
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==comapany.toString().trim()){ 
     var OK = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        cancel.Click();
        aqUtils.Delay(1000, Indicator.Text);;
        companyName.setText("");
        ValidationUtils.verify(false,true,"Company Number is not listed in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
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
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
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
    var tab = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
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
  var code = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(department.toString().trim());
  aqUtils.Delay(3000, Indicator.Text);;
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  aqUtils.Delay(5000, Indicator.Text);;
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==department.toString().trim()){ 
     var OK = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        cancel.Click();
        aqUtils.Delay(1000, Indicator.Text);;
        Depart.setText("");
        ValidationUtils.verify(false,true,"Department Number is not listed in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
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
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
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
    var tab = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
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
  var code = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(buss_unit.toString().trim());
  aqUtils.Delay(3000, Indicator.Text);;
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  aqUtils.Delay(5000, Indicator.Text);;
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==buss_unit.toString().trim()){ 
     var OK = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        cancel.Click();
        aqUtils.Delay(1000, Indicator.Text);;
        BussUnit.setText("");
        ValidationUtils.verify(false,true,"Business Unit Number is not listed in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
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
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
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
    var tab = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
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
  var code = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(EmpCategory.toString().trim());
  aqUtils.Delay(3000, Indicator.Text);;
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  aqUtils.Delay(5000, Indicator.Text);;
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==EmpCategory.toString().trim()){ 
     var OK = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
        OK.Click();
          
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        cancel.Click();
        aqUtils.Delay(1000, Indicator.Text);;
        EmpCat.setText("");
        ValidationUtils.verify(false,true,"Department Number is not listed in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
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
  Obj_Address.Click();
  aqUtils.Delay(1000, Indicator.Text);;
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

  aqUtils.Delay(5000, Indicator.Text);;
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  if(serch.isEnabled())
  serch.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
   serch.Click(); 
  }
  
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  do{
  aqUtils.Delay(5000, Indicator.Text);;
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
  tableList[tl] = table.getItem(i).getText_2(0).OleValue.toString().trim()+"-"+table.getItem(i).getText_2(1).OleValue.toString().trim();
//  Log.Message(tableList[tl])
  tl++;
  }
    }
    var tab = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
    var tabVisible = tab.wEnabled(1,true)
    if(tabVisible){ 
      tab.Click(-1,-1);
    }
    }while(tabVisible)
    
    
    var stat = true;
    for(var exl =0;exl<ExcelData.length;exl++){
//    Log.Warning(ExcelData[exl]);
        var compStatus = false;
    var bb1 = "";
        for(var cnt = 0;cnt<tableList.length;cnt++){
//    Log.Message(tableList[cnt])
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
  var code = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(value.toString().trim());
//  aqUtils.Delay(3000, Indicator.Text);;
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
   if(serch.isEnabled())
  serch.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
   serch.Click(); 
  }
  aqUtils.Delay(4000, Indicator.Text);;
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value.toString().trim()){ 
    temp = table.getItem(i).getText_2(1).OleValue.toString().trim();
     var OK = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
  if(OK.isEnabled()){
  OK.HoverMouse();
//  ReportUtils.logStep_Screenshot();
  OK.Click();
  
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
  OK.HoverMouse();
  ReportUtils.logStep_Screenshot();
   OK.Click(); 
  }
        ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
        break;
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
  if(cancel.isEnabled()){
  cancel.HoverMouse();
  ReportUtils.logStep_Screenshot();
  cancel.Click();
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
  cancel.HoverMouse();
  ReportUtils.logStep_Screenshot();
   cancel.Click(); 
  }
        aqUtils.Delay(1000, Indicator.Text);;
        Obj_Address.setText("");
        ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
    if(cancel.isEnabled()){
  cancel.HoverMouse();
  ReportUtils.logStep_Screenshot();
  cancel.Click();
  }
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
  cancel.HoverMouse();
  ReportUtils.logStep_Screenshot();
   cancel.Click(); 
  }
//    aqUtils.Delay(1000, Indicator.Text);;
    ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    Obj_Address.setText("");
  }
        }
return temp;
}


function config_with_Maconomy_Validation_Name_column2(Obj_Address,wizName,value,ExcelData,fieldName){ 
var temp = "";
   if(value!=""){
  Obj_Address.Click();
  aqUtils.Delay(1000, Indicator.Text);;
  Sys.Desktop.KeyDown(0x11);
  Sys.Desktop.KeyDown(0x47);
  Sys.Desktop.KeyUp(0x11);
  Sys.Desktop.KeyUp(0x47);
  aqUtils.Delay(3000, Indicator.Text);;
  //====================================
//var sheet = [];
//sheet[0] = "value";
//sheet[1] = "UDepartment";
//var ExcelData = ExlArray;
var tableList = [];
var tl = 0;

  aqUtils.Delay(3000, Indicator.Text);;
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  do{
  aqUtils.Delay(5000, Indicator.Text);;
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
  tableList[tl] = table.getItem(i).getText_2(0).OleValue.toString().trim()+"-"+table.getItem(i).getText_2(2).OleValue.toString().trim();
//  Log.Message(tableList[tl])
  tl++;
  }
    }
    var tab = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
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
  var code = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText(value.toString().trim());
  aqUtils.Delay(3000, Indicator.Text);;
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  aqUtils.Delay(5000, Indicator.Text);;
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value.toString().trim()){ 
    temp = table.getItem(i).getText_2(2).OleValue.toString().trim();
     var OK = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
        OK.HoverMouse();
//        ReportUtils.logStep_Screenshot();
        OK.Click();
        ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        cancel.HoverMouse();
        ReportUtils.logStep_Screenshot();
        cancel.Click();
        aqUtils.Delay(1000, Indicator.Text);;
        Obj_Address.setText("");
        ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
    cancel.HoverMouse();
    ReportUtils.logStep_Screenshot();
    cancel.Click();
    aqUtils.Delay(1000, Indicator.Text);;
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
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
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
    var tab = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
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
  var code = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
//  code.Keys("[Tab]");
//  aqUtils.Delay(2000, Indicator.Text);;
//  var code = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText("*"+value.toString().trim()+"*");
  aqUtils.Delay(3000, Indicator.Text);;
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  serch.Click();
  aqUtils.Delay(5000, Indicator.Text);;
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim().indexOf(value.toString().trim())!=-1){ 
     var OK = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
        OK.Click();
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        cancel.Click();
        aqUtils.Delay(1000, Indicator.Text);;
        Obj_Address.setText("");
        ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
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
  aqUtils.Delay(5000, Indicator.Text);;
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
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
  if(serch.isEnabled())
  serch.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
   serch.Click(); 
  }
  
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  do{
  aqUtils.Delay(5000, Indicator.Text);;
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
    var tab = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("ToolBar", "", 1);
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
//        Log.Message("splits[0] :"+splits[0])
//        Log.Message("splits[1] :"+splits[1])
//-------------------For INDIA ------------------------------
      if(splits[0]==Job_Type.toString().trim()){ 
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
//          Log.Message("billable[exl] :"+billable[exl]);
    for(var cnt = 0;cnt<tableList.length;cnt++){ 

//      Log.Message("tableList[cnt] :"+tableList[cnt]);      
      if(tableList[cnt].toLowerCase().indexOf(billable[exl].toLowerCase())!=-1){ 
//      Log.Message("billable[exl] :"+billable[exl]);
//      Log.Message("tableList[cnt] :"+tableList[cnt]);
       compStatus = true;
       break;
      }
      }
      if(!compStatus){ 
        if(stat){
        Log.Message(billable[exl])
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
  var code = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
//  code.Keys("[Tab]");
//  aqUtils.Delay(2000, Indicator.Text);;
//  var code = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
  code.setText("*"+value.toString().trim()+"*");
//  aqUtils.Delay(3000, Indicator.Text);;
  var serch = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
  Sys.HighlightObject(serch);
   if(serch.isEnabled())
  serch.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
   serch.Click(); 
  }
  aqUtils.Delay(4000, Indicator.Text);;
  var table = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
  Sys.HighlightObject(table);
  var itemCount = table.getItemCount();
  if(itemCount>0){ 
  for(var i=0;i<itemCount;i++){
    if(table.getItem(i).getText_2(0).OleValue.toString().trim().indexOf(value.toString().trim())!=-1){ 
     var OK = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
      if(OK.isEnabled())
  OK.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
   OK.Click(); 
  }
         ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");  
         break;
    }
    else{ 
      Sys.Desktop.KeyDown(0x28);
      Sys.Desktop.KeyUp(0x28);
      if(i==itemCount-1){ 
        var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
        if(cancel.isEnabled())
  cancel.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
   cancel.Click(); 
  }
//        aqUtils.Delay(1000, Indicator.Text);;
        Obj_Address.setText("");
        ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
      }
    }
      
    }
  }
  else { 
    var cancel = Sys.Process("Maconomy").SWTObject("Shell", wizName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
    if(cancel.isEnabled())
  cancel.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);;
   cancel.Click(); 
  }
//    aqUtils.Delay(1000, Indicator.Text);;
    ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
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
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(1).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
          OK.Click();
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
          cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }
      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim())
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
    var code = Sys.Process("Maconomy").SWTObject("Shell", "Transaction Type").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McValuePickerWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Search").OleValue.toString().trim()+" ");
    Sys.HighlightObject(serch);
    if(serch.isEnabled())
  serch.Click();
  else{ 
    aqUtils.Delay(3000, Indicator.Text);
   serch.Click(); 
  }
    aqUtils.Delay(5000, Indicator.Text);
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "OK").OleValue.toString().trim())
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
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
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
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Language, "Cancel").OleValue.toString().trim());
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
    var code = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "");
    code.setText(value);
    aqUtils.Delay(3000, Indicator.Text);;
    var serch = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 1).SWTObject("Button", "Search ");
    Sys.HighlightObject(serch);
    serch.Click();
    aqUtils.Delay(5000, Indicator.Text);;
    var table = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 2).SWTObject("McGrid", "", 2);
    Sys.HighlightObject(table);
    var itemCount = table.getItemCount();
    if(itemCount>0){ 
    for(var i=0;i<itemCount;i++){
      if(table.getItem(i).getText_2(0).OleValue.toString().trim()==value){ 
       var OK = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "OK");
          OK.Click();
          checkmark = true;
          ValidationUtils.verify(true,true,fieldName+" is listed and  Selected in Maconomy");
      }
      else{ 
        Sys.Desktop.KeyDown(0x28);
        Sys.Desktop.KeyUp(0x28);
        if(i==itemCount-1){ 
          var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
          cancel.Click();
          aqUtils.Delay(1000, Indicator.Text);;
          ObjectAddrs.setText("");
          ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
        }
      }      
      }
    }
    else { 
      var cancel = Sys.Process("Maconomy").SWTObject("Shell", popupName).SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Button", "Cancel");
      cancel.Click();
      aqUtils.Delay(1000, Indicator.Text);;
      ObjectAddrs.setText("");
      ValidationUtils.verify(false,true,fieldName+" is not listed  in Maconomy");
    }
    return checkmark;
}



function StartTime(){ 
var dif;
var TodayValue = aqDateTime.Today();
var StringTodayValue = aqConvert.DateTimeToStr(TodayValue);
var EncodedDate = aqConvert.DateTimeToFormatStr(StringTodayValue,"%d%#B%Y"); 
var STIME = getFormattedCurrentTime()
Log.Message("Start DATE & TIME :"+EncodedDate +" "+STIME)
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

function getFormattedCurrentTime(){
  TodayValue = aqConvert.DateTimeToFormatStr(aqDateTime.Time(), "%H:%M:%S");
  return TodayValue;
}