﻿
function nm(){ 
  
//var  PropArray = new Array("JavaClassName", "Index");
//var  ValuesArray = new Array("McDatePickerWidget", "2");
//
//  p = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Senior Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1)
//  obj = p.FindAll(PropArray, ValuesArray, 1000);
//  var JobStart_Date=null;
//  Log.Message(obj.length)
//  
//  if(obj.length>=3){
//      if(obj[0].isVisible){
//        Sys.HighlightObject(obj[0]);
//        JobStart_Date = obj[0];        
//      }
//  
//  }
//Sys.HighlightObject(JobStart_Date);


//  PropArray = new Array("JavaClassName", "Index");
//  ValuesArray = new Array("SingleToolItemControl", "8");
//  p = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Senior Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//  obj = p.FindAll(PropArray, ValuesArray, 1000);
//  
//for (let i_count = 0; i_count < obj.length; i_count++){ 
//  if(obj[i_count].toolTipText=="Submit"){
//  Sys.HighlightObject(obj[i_count]);
//  Submit = obj[i_count];
//  break;
// }
//}
//
//Log.Message(Submit.FullName)
//Sys.HighlightObject(Submit);


//var approve_bar ;
//PropArray = new Array("JavaClassName", "Index");
//ValuesArray = new Array("TabControl", "1");
//  p = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Senior Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//      var approv_bar = p.FindAll(PropArray, ValuesArray, 1000);
//      for (let j_count = 0; j_count < approv_bar.length; j_count++){ 
//        if(approv_bar[j_count].Visible){ 
//          approve_bar = approv_bar[j_count];
//          Log.Message(approve_bar.FullName);
//          break;
//        }
//        
//        }
//      
//Sys.HighlightObject(approve_bar);



      PropArray = new Array("JavaClassName", "Index","Visible");
  ValuesArray = new Array("McGroupWidget", "1", "true");
  p = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  var newQuote = "";
for (let i_count = 0; i_count < obj.length; i_count++){ 
if((obj[i_count].Exists) && (obj[i_count].Parent.JavaClassName=="Composite") && (obj[i_count].Parent.Index==1))
newQuote = obj[i_count];
}
Sys.HighlightObject(newQuote)
newQuote = newQuote.SWTObject("Composite", "", 7).SWTObject("McTextWidget", "", 2)
Sys.HighlightObject(newQuote)
//
////newQuote = newQuote.Parent.Name;
//Log.Message(newQuote.Parent.Name)
//newQuote = newQuote.SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
//Sys.HighlightObject(newQuote);
//newQuote = newQuote.getText().OleValue.toString().trim();;    

//  PropArray = new Array("JavaClassName", "Visible");
//  ValuesArray = new Array("McPaneGui$10", "true");
//p = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
////obj = p.FindAll("JavaClassName", "McPaneGui$10", 1000);
//obj = p.FindAll(PropArray, ValuesArray, 1000);
//  var Page = "";
//for (let i_count = 0; i_count < obj.length; i_count++){ 
//if(obj[i_count].Exists)
//Page = obj[i_count];
//}
//Sys.HighlightObject(Page)
    

//  PropArray = new Array("JavaClassName", "Index","Visible");
//  ValuesArray = new Array("McGroupWidget", "2", "true");
//  p = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//  obj = p.FindAll(PropArray, ValuesArray, 1000);
//  Log.Message(obj.length)
//  var McGroupWidget = "";
//for (let i_count = 0; i_count < obj.length; i_count++){ 
//if((obj[i_count].Exists) && (obj[i_count].Parent.JavaClassName=="Composite") && (obj[i_count].Parent.Index==3))
//McGroupWidget = obj[i_count];
//}
//Sys.HighlightObject(McGroupWidget)
//
//Log.Message(McGroupWidget.Parent.Name)
//var submittedby = McGroupWidget.SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
//Sys.HighlightObject(submittedby);
//
//var approvedby = McGroupWidget.SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
//Sys.HighlightObject(approvedby);



// var show_budget;   
//      PropArray = new Array("JavaClassName", "Index","Visible");
//  ValuesArray = new Array("McGroupWidget", "1", "true");
//  p = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//  obj = p.FindAll(PropArray, ValuesArray, 1000);
//  Log.Message(obj.length)
//for (let i_count = 0; i_count < obj.length; i_count++){ 
//if((obj[i_count].Exists) && (obj[i_count].ChildCount>=8))
//show_budget = obj[i_count];
//}
//    Sys.HighlightObject(show_budget);


//var ApprovalTableBar ;
//  PropArray = new Array("JavaClassName", "Index","Visible");
//  ValuesArray = new Array("TabControl", "1","true");
//  p = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//  obj = p.FindAll(PropArray, ValuesArray, 1000);
//  Log.Message(obj.length)
//  for (let i_count = 0; i_count < obj.length; i_count++){ 
//  if((obj[i_count].Exists) && (obj[i_count].Parent.JavaClassName=="PTabItemPanel") && (obj[i_count].Parent.Index==1)){
//    ApprovalTableBar = obj[i_count]
//    break;      
//  }
//}
//Sys.HighlightObject(ApprovalTableBar);

//  PropArray = new Array("JavaClassName", "Visible");
//  ValuesArray = new Array("TabControl", "true");
//  p = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//  obj = p.FindAll(PropArray, ValuesArray, 1000);
//  var Information = ""
//for (let i_count = 0; i_count < obj.length; i_count++){ 
//  if(obj[i_count].text=="Information"){
//  Sys.HighlightObject(obj[i_count]);
//  Information = obj[i_count];
//  break;
// }
//}
//
//Log.Message(Information.FullName)
//Sys.HighlightObject(Information);


//var approve_bar ;
//  PropArray = new Array("JavaClassName", "Index","ChildCount");
//  ValuesArray = new Array("PTabItemPanel", "3","1");
//  p = Sys.Process("Maconomy", 3).SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//  obj = p.FindAll(PropArray, ValuesArray, 1000);
//  Log.Message(obj.length)
//  for (let i_count = 0; i_count < obj.length; i_count++){ 
//  if((obj[i_count].Exists)&&(obj[i_count].isVisible())){
//    approve_bar = obj[i_count].SWTObject("TabControl", "");
//    break;      
//  }
//}
//Sys.HighlightObject(approve_bar);
//Log.Message("quteNumber :"+quteNumber)



  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("SingleToolItemControl", "true");
  p = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var convertToOrder = "";
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].toolTipText=="Create Time Sheet"){
  Sys.HighlightObject(obj[i_count]);
  convertToOrder = obj[i_count];
  break;
 }
}

Log.Message(convertToOrder.FullName)
Sys.HighlightObject(convertToOrder);

}


function DemoFUn (){ 
  
var Maconomy_ParentAddress = Sys.Process("Maconomy", 2).SWTObject("Shell", "Deltek Maconomy - 1221 Management (TST)")
//var approve_bar ;
//  PropArray = new Array("JavaClassName", "Index", "Visible");
//  ValuesArray = new Array("Label", "1", "true");
//  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//  obj = p.FindAll(PropArray, ValuesArray, 1000);
//  Log.Message(obj.length)
//  for (let i_count = 0; i_count < obj.length; i_count++){ 
//  if(obj[i_count].Exists){
//    approve_bar = obj[i_count];
//    break;      
//  }
//}
//Sys.HighlightObject(approve_bar);



//  PropArray = new Array("JavaClassName", "Visible");
//  ValuesArray = new Array("SingleToolItemControl", "true");
//  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//  obj = p.FindAll(PropArray, ValuesArray, 1000);
//  var convertToOrder = "";
//for (let i_count = 0; i_count < obj.length; i_count++){ 
//  if(obj[i_count].toolTipText=="Create Time Sheet"){
//  Sys.HighlightObject(obj[i_count]);
//  convertToOrder = obj[i_count];
//  break;
// }
//}
//
//Log.Message(convertToOrder.FullName)
//Sys.HighlightObject(convertToOrder);



//  var ObjAddress ;
//  PropArray = new Array("JavaClassName", "Index", "Visible");
//  ValuesArray = new Array("McDatePickerWidget", "1", "true");
//  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//  obj = p.FindAll(PropArray, ValuesArray, 1000);
//  for (let i_count = 0; i_count < obj.length; i_count++){ 
//  if(obj[i_count].Exist){
//    ObjAddress = obj[i_count];
//    break;      
//  }
//}
//Sys.HighlightObject(ObjAddress.SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("McTextWidget", "", 3));

//
//
  var ObjAddress ;

  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll("JavaClassName", "McGrid", 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Exist){
    ObjAddress = obj[i_count];
    break;      
  }
}
Sys.HighlightObject(ObjAddress);

//     var Allaprovetab ;
//  PropArray = new Array("JavaClassName", "Index","ChildCount","Visible");
//  ValuesArray = new Array("PTabItemPanel", "3","1",true);
//  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//  obj = p.FindAll(PropArray, ValuesArray, 1000);
//  Log.Message(obj.length)
//  let objHeight = 1000;
//  for (let i_count = 0; i_count < obj.length; i_count++){ 
//  if((obj[i_count].Exists)&&(obj[i_count].Parent.Left>0)){
//    if(objHeight>obj[i_count].Parent.Height)
//    Allaprovetab = obj[i_count];  
//  }
//}
//Allaprovetab = Allaprovetab.SWTObject("TabControl", "");  
//Log.Message(Allaprovetab.length)
//Sys.HighlightObject(Allaprovetab);
}




function HK(){
  
var Maconomy_ParentAddress = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)");
//var ApvPerson = getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"PTabItemPanel",1,"Composite",1);
Sys.HighlightObject(Maconomy_ParentAddress);

var groupWidget = getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,"McGroupWidget",2,"Composite",2);
var Excul = groupWidget.SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
Sys.HighlightObject(Excul);
var Incul = groupWidget.SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
Sys.HighlightObject(Incul);

}


//Finding Object in selected screen with JavaClassName and Index Property
function getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,JClassName,Obj_Index){ 
  var ObjAddress ;
  PropArray = new Array("JavaClassName", "Index", "Visible");
  ValuesArray = new Array(JClassName, Obj_Index, "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Exists){
    ObjAddress = obj[i_count];
    break;      
  }
}
if(ObjAddress!=null)
Sys.HighlightObject(ObjAddress);
else
ObjAddress==null
return ObjAddress;
}

//Finding Object in selected screen with JavaClassName and Index Property with Parent Index
function getObjectAddress_JavaClasssName_and_Index_withParentIndex(Maconomy_ParentAddress,JClassName,Obj_Index,Parent_Index){ 
  var ObjAddress ;
  PropArray = new Array("JavaClassName", "Index", "Visible");
  ValuesArray = new Array(JClassName, Obj_Index, "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Parent.Index == Parent_Index){
    ObjAddress = obj[i_count];
    break;      
  }
}
if(ObjAddress!=null)
Sys.HighlightObject(ObjAddress);
else
ObjAddress==null
return ObjAddress;
}


function waitUntil_MaconomyScreen_loaded_Completely(){ 
var count = 0;
do{
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ }
  aqUtils.Delay(5000, Indicator.Text);
  if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    break;
  }else{ 
    count++;
  }
}while(count<5)
}


//Finding Object in selected screen with JavaClassName and Index Property with Parent Index
function getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,JClassName,Obj_Index,Parent_JavaClass,Parent_Index){ 
  var ObjAddress ;
  PropArray = new Array("JavaClassName", "Index", "Visible");
  ValuesArray = new Array(JClassName, Obj_Index, "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length);
  Log.Message(obj.FullName)
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if((obj[i_count].Parent.Index == Parent_Index) && (obj[i_count].Parent.JavaClassName == Parent_JavaClass)){
    ObjAddress = obj[i_count];
    break;      
  }
}
if(ObjAddress!=null)
Sys.HighlightObject(ObjAddress);
else
ObjAddress==null
return ObjAddress;
}



function Hug(){ 
var Excl_Tax = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
Sys.HighlightObject(Excl_Tax);
var grandTotal = Aliases.Maconomy.Group3.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite6.Composite2.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite2.McGroupWidget.SWTObject("Composite", "", 2).SWTObject("McTextWidget", "", 2);
Sys.HighlightObject(grandTotal);
}