//USEUNIT ReportUtils
function todo(lvl){ 
aqUtils.Delay(5000, Indicator.Text);
Client_Managt = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite2.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.ToDoList;
 
if(lvl==3){
Client_Managt.ClickItem("|Approve Purchase Order by Type (Substitute) (*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|Approve Purchase Order by Type (Substitute) (*)");
}
if(lvl==2){
Client_Managt.ClickItem("|Approve Purchase Order by Type(*)");
ReportUtils.logStep_Screenshot(); 
Client_Managt.DblClickItem("|Approve Purchase Order by Type(*)");
}
}

function ApprovalStatus(){ 
var POapproval = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.TabFolderPanel.TabControl2;
POapproval.HoverMouse();
ReportUtils.logStep_Screenshot();
POapproval.Click();
aqUtils.Delay(3000, Indicator.Text);;
var approvertable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite4.Composite.PTabFolder.Composite2.McClumpSashForm.Composite.Composite.McTableWidget.McGrid
ReportUtils.logStep_Screenshot();
}

