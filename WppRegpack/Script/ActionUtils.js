//USEUNIT ReportUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils





//Click on the element and clear the text before entering
//comment
function ClearAndEnterKeys(control, keys)
{
control.HoverMouse();
control.Click();
control.Keys("[Home]");
control.Keys("![End]");
control.Keys(keys);
}

//Performs click action and enters the text on the element
function ClickAndEnterText(element, field, value)
{
if(element.Exists&&element.Visible)
{
Log.Message("Performing click action on the '"+field+"'");
element.Click();
Log.Message("Entering the value '"+value+"' on the '"+field+"' field");
element.SetText("");
element.Keys(value);
}
else
{
Log.Error("Element is not found or accessible "+field);
}
}

//Performs click action on the element
function ClickAction(element, field)
{
if(element.Exists)
{
element.SetFocus();
Log.Message("Performing click action on the '"+field+"'");
element.Click();
}
else
{
Log.Error("Element is not found or accessible Object "+field);
}
}

function simpleDelay(){
  Delay(1000);
}

function mediumDelay(){
  Delay(5000);
}

function complexDelay(){
  Delay(10000);
}

//Wait for Page to Load
function WaitForWindowLoad(ObjectType)
{
var eleObject = ObjectType;

for(i=0;i<=30;i++)
{
if ((eleObject.Exists && eleObject.VisibleOnScreen))
{
Log.Message("Window was found.");
break;
}
else
{
Log.Warning("Waiting for the page to load...");
Delay(500, "Waiting for Page to load....");
}
}
}


/*
Find element with single and multi property
var obj = {
menuItems  : "Name;VBObject(\"trvSysAdmin\")?Text;Roles",  
Menu_Role : "Text?Access Archive;Width?74?25"
};
where,
//Name - Object type, 
//VBObject(\"trvSysAdmin\") - Object Value
*/
function findElementMultiProp(page, locator, depth=100){

var eleProperties = locator.split(";");
var element;
var key = [];
var value = [];
for(i=0; i<eleProperties.length; i++){
if(eleProperties[i].split("?").length ==3)
{
depth = eleProperties[i].split("?")[2];
}
key.push(eleProperties[i].split("?")[0]) ;
value.push(eleProperties[i].split("?")[1]) 
}
element = page.Find(key,value,depth);
return element;
}


function SelectListViewItem(listName,listitemName)
{
  var listViewObj=ObjectSelection.GetObjectByName(listName);
  var count=listViewObj.Items.Count;
  for (i=0;i<count;i++)
  {
    listViewObj.Items.Item(i).set_Selected(false);
    if(listViewObj.Items.Item(i).Text.OleValue == listitemName)
    {
      listViewObj.Items.Item(i).set_Selected(true);
    }    
  }
}

function SelectItem(dropdownName,itemName)
{
  try
  {
     var selectedItemName = null;
     for(i = 0; i< dropdownName.Items.Count;)
     {
              
      selectedItemName = dropdownName.Items.Item(i).Text;
       if(selectedItemName == itemName)
       {
         selectedItemName.set_Selected("true");        
       }
       i++;       
     }
  }
  catch (e)
  {
    throw(e);
  }
}

function GetSelectedText(dropdownName)
{
  try
  {
     var  itemObj =null, selectedItemName = null;
     selectedItemName = dropdownName.Text.OleValue;
     return selectedItemName;  
  }
  catch (e)
  {
    throw(e);
  }
}

//TextTobeSelected should be string , ddElement should be object of type element
function selectItemInDropDown(ddElement, TextTobeSelected){
ddElement.ClickItem(TextTobeSelected);
}

//elementTobeCollpased should be string strts with |, if if contains multiple element then it must be like ("|Server Databases|OQTEST2")
function collapseTree(element, elementTobeCollpased){
element.Collapse(elementTobeCollpased);
}

//elementTobeExpand should be string strts with |, if if contains multiple element then it must be like ("|Server Databases|OQTEST2")
function ExpandTree(element, elementTobeExpand){
element.Collapse(elementTobeExpand);
}

//element is the checkbox element
function selectCheckBox(element){
element.ClickButton(cbChecked);
}

//element is the checkbox element
function unSelectCheckBox(element){
element.ClickButton(cbUnchecked);
}

function GetDate(dateControlName)
{
  try
  {
     var startDate = null;
     startDate = dateControlName.wDate;
     Log.Message(startDate);
     return startDate
  }
  catch (e)
  {
    Log.Message(e.discription);
  }
}


function addPicture(imageName)
{
  var w = Sys.Desktop.Picture();
  Log.Picture(w, imageName);
}

/*
function Click(btnName)
{
  try
  {
    ObjectSelection.GetObjectByName(btnName).Click();
  }
  catch(e)
  {
    Log.Error(e);
  }   
}
*/



/**
  *  This function Navigates to workspace
  */
function Moving_intoWorkspace(Maconomy_ParentAddress,WorkSpace){
  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("McMaconomyPShelfMenuGui$3", "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Exists){
    tree_Object = obj[i_count]
    break;      
  }
}

  PropArray = new Array("JavaClassName", "Visible");
  ValuesArray = new Array("Tree", "true");
  obj = tree_Object.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Exists){
    tree_Object = obj[i_count]
    break;      
  }
}
Sys.HighlightObject(tree_Object);
tree_Object.ClickItem("|"+WorkSpace);
ReportUtils.logStep_Screenshot();
tree_Object.DblClickItem("|"+WorkSpace);

ReportUtils.logStep("INFO", "Moving into "+WorkSpace+" Workspace");
TextUtils.writeLog("Moving into "+WorkSpace+" Workspace");
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
function getObjectAddress_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,JClassName,Obj_Index,Parent_ChildCount){ 
  var ObjAddress ;
  PropArray = new Array("JavaClassName", "Index", "Visible");
  ValuesArray = new Array(JClassName, Obj_Index, "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Parent.ChildCount == Parent_ChildCount){
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
function getObjectAddress_forSlidingPanel_JavaClasssName_and_Index_withParent(Maconomy_ParentAddress,JClassName,Obj_Index,Parent_JavaClass,Parent_Index){ 
  var ObjAddress ;
  PropArray = new Array("JavaClassName", "Index", "Visible");
  ValuesArray = new Array(JClassName, Obj_Index, "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
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
//Finding Object in selected screen with JavaClassName and Index Property with Parent JavaClassName
function getObjectAddress_JavaClasssName_and_Index_withParentClassName(Maconomy_ParentAddress,JClassName,Obj_Index,Parent_JavaClassName){ 
  var ObjAddress ;
  PropArray = new Array("JavaClassName", "Index", "Visible");
  ValuesArray = new Array(JClassName, Obj_Index, "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Parent.JavaClassName == Parent_JavaClassName){
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

//Finding Object in selected screen with JavaClassName and Index Property with Parent JavaClassName
function getObjectAddress_JavaClasssName_and_Index_withChildCount(Maconomy_ParentAddress,JClassName,Obj_Index,ObjChildCount){ 
  var ObjAddress ;
  PropArray = new Array("JavaClassName", "Index", "Visible");
  ValuesArray = new Array(JClassName, Obj_Index, "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].ChildCount == ObjChildCount){
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

//Finding Object in selected screen with Single Property Check
function getObjectAddress_withSingleProperty_Check(Maconomy_ParentAddress,JClassName){ 
  var ObjAddress ;

  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll("JavaClassName", JClassName, 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Visible==true){
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


//Finding Object in selected screen with JavaClassName and Index Property
function getObjectAddress_JavaClasssName(Maconomy_ParentAddress,JClassName,tooltipText){ 
  var ObjAddress ;
  PropArray = new Array("JavaClassName",  "Visible");
  ValuesArray = new Array(JClassName, "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].toolTipText==tooltipText){
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



//Finding Object in selected screen with JavaClassName and Index Property
function getObjectAddress_JavaClasssName_conatinsTooltip(Maconomy_ParentAddress,JClassName,ele_Text){ 
  var ObjAddress ;
  PropArray = new Array("JavaClassName",  "Visible");
  ValuesArray = new Array(JClassName, "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
    if(obj[i_count].toolTipText!=null)
  if(obj[i_count].toolTipText.OleValue.toString().trim().indexOf(ele_Text)!=-1){
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


//Finding Object in selected screen with JavaClassName and Index Property
function getObjectAddress_JavaClasssName_Index_conatinsTooltip(Maconomy_ParentAddress,JClassName,index,Text){ 
  var ObjAddress ;
  PropArray = new Array("JavaClassName", "Index", "Visible");
  ValuesArray = new Array(JClassName, index,"true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
    if(obj[i_count].toolTipText!=null)
  if(obj[i_count].toolTipText.OleValue.toString().trim().indexOf(Text)!=-1){
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


function getObjectAddress_JavaClasssName_withTabText(Maconomy_ParentAddress,JClassName,Text){ 
  var ObjAddress ;
  PropArray = new Array("JavaClassName",  "Visible");
  ValuesArray = new Array(JClassName, "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].text==Text){
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

function getObjectAddress_JavaClasssName_Index_withTabText(Maconomy_ParentAddress,JClassName,oIndex,Text){ 
  var ObjAddress ;
  PropArray = new Array("JavaClassName","Index", "Visible");
  ValuesArray = new Array(JClassName,oIndex, "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].text==Text){
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


function DoubleClick_with_Screenshot(Workspace_Client){ 
Workspace_Client.HoverMouse();
ReportUtils.logStep_Screenshot("");
Workspace_Client.DblClick();
}


function Select_timesheet_from_workspace(){ 

if(ImageRepository.ImageSet.TimeExpense.Exists()){
 ImageRepository.ImageSet.TimeExpense.Click();
}
else if(ImageRepository.ImageSet.TimeExpense1.Exists()){
ImageRepository.ImageSet.TimeExpense1.Click();
}
else{
ImageRepository.ImageSet.TimeExpense2.Click();
}

}


function Select_AccountPayable_from_workspace(){ 
if(ImageRepository.ImageSet.AccountPayable.Exists()){
ImageRepository.ImageSet.AccountPayable.Click();// GL
}
else if(ImageRepository.ImageSet.AccountPayable2.Exists()){
ImageRepository.ImageSet.AccountPayable2.Click();
}
else{
ImageRepository.ImageSet.AccountPayable2.Click();
}

}
function Select_Jobs_from_workspace(){ 


if(ImageRepository.ImageSet3.Jobs.Exists()){
 ImageRepository.ImageSet3.Jobs.Click();
}
else if(ImageRepository.ImageSet.Job.Exists()){
ImageRepository.ImageSet.Job.Click();
}
else{
ImageRepository.ImageSet.Jobs1.Click();
}


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



// Finding Created Budget from To-Do's List
function ToDos_Selection(Maconomy_ParentAddress, apr_Level, lvl, MainAprover, MainApprover_byType, Substitute, Substitute_byType){ 
  TextUtils.writeLog("Loged into Level "+apr_Level+" Approver login"); 
  
  var toDo = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 5).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("TabControl", "", 4);
  toDo.HoverMouse();
  ReportUtils.logStep_Screenshot();
  toDo.DBlClick();
  TextUtils.writeLog("Entering into To-Dos List");
  
  aqUtils.Delay(3000, Indicator.Text);
  //To Maximaize the window
  Sys.Desktop.KeyDown(0x12);
  Sys.Desktop.KeyDown(0x20);
  Sys.Desktop.KeyUp(0x12);
  Sys.Desktop.KeyUp(0x20);
  Sys.Desktop.KeyDown(0x58);
  Sys.Desktop.KeyUp(0x58);  
  aqUtils.Delay(1000, Indicator.Text);

  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("SingleToolItemControl", "1");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var refresh;
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].toolTipText=="Refresh ToDo's"){
  Sys.HighlightObject(obj[i_count]);
  refresh = obj[i_count];
  break;
  }
}
Log.Message(refresh.FullName)
Sys.HighlightObject(refresh)
refresh.Click();
aqUtils.Delay(15000, Indicator.Text);
if(ImageRepository.ImageSet.ToDos_Icon.Exists()){ 
  
}


  PropArray = new Array("JavaClassName", "Index");
  ValuesArray = new Array("Tree", "1");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  var Client_Managt;
for (let i_count = 0; i_count < obj.length; i_count++){ 
  if(obj[i_count].Visible){
  Sys.HighlightObject(obj[i_count]);
  Client_Managt = obj[i_count];
  break;
  }
}
Log.Message(Client_Managt.FullName)
Sys.HighlightObject(Client_Managt)


Log.Message(lvl)
var listPass = true;
if(lvl==3){
  
Log.Message("Substitute :"+Substitute);
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
  Log.Message(temp1.length)
  Log.Message(temp.indexOf(Substitute+" (")!=-1)
  Log.Message(temp1.length>=2)
if((temp.indexOf(Substitute+" (")!=-1)&&(temp1.length>=2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into "+Substitute+" from To-Dos List");
listPass = false; 
  }
  
}



if(listPass && Substitute_byType!=null){

for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if((temp.indexOf(Substitute_byType+" (")!=-1)&&(temp1.length>=2)){ 
Client_Managt.ClickItem("|"+temp);   
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp);  
TextUtils.writeLog("Entering into "+Substitute_byType+" from To-Dos List");
listPass = false; 
  }
  }
}

}


if(lvl==2){
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if(temp.indexOf(MainAprover+" (")!=-1){  
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into "+MainAprover+" from To-Dos List");
var listPass = false;   
break;
  }
}  

if(listPass && MainApprover_byType!=null){
for(var j=0;j<Client_Managt.getItemCount();j++){ 
  var temp = Client_Managt.getItem(j).getText().OleValue.toString().trim();
  var temp1 = temp.split("(");
if(temp.indexOf(MainApprover_byType+" (")!=-1){ 
Client_Managt.ClickItem("|"+temp);    
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|"+temp); 
TextUtils.writeLog("Entering into "+MainApprover_byType+" from To-Dos List");
var listPass = false; 
break;  
  }
} 
  }
}
 

}
