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