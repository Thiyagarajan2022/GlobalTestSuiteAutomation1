//USEUNIT ReportUtils
//USEUNIT TextUtils
//USEUNIT WorkspaceUtils

function Employeenumber(SelectionBilling,EmpNo,B_Estimatelines){ 
  for(var t=0;t<SelectionBilling.getItemCount();t++){ 
      for(var g = 0;g<B_Estimatelines.length;g++){
        var temp = B_Estimatelines[g].split("*");
        if(SelectionBilling.getItem(t).getText_2(0).OleValue.toString().trim().indexOf(temp[0])==0){
          if(SelectionBilling.getItem(t).getText_2(0).OleValue.toString().trim().indexOf("T")==0){
      aqUtils.Delay(100, Indicator.Text);
      SelectionBilling.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
      var Employee = Aliases.Maconomy.InvoicingFromBudget.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget2;
      WorkspaceUtils.waitForObj(Employee);
      Employee.Click();
//      if((Employee.getText()=="")||(Employee.getText()==null)){ 
      if((EmpNo!="")&&(EmpNo!=null)){
      Employee.HoverMouse();
      Employee.Click();
      WorkspaceUtils.SearchByValue(Employee,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
      
var SaveStat = true;
      var Save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2;
      Log.Message(Save.FullName)
      WorkspaceUtils.waitForObj(Save);
      for(var i=0;i<Save.ChildCount;i++){
        Log.Message(Save.Child(i).Name)
        Log.Message(Save.Child(i).toolTipText)
        if((Save.Child(i).isVisible())&&(Save.Child(i).toolTipText==JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Save Invoice Selection Line").OleValue.toString().trim())){
          Save = Save.Child(i);
          WorkspaceUtils.waitForObj(Save);
          ReportUtils.logStep_Screenshot("");
          Save.Click();
          SaveStat = false;
          break;
        }
        
      } 
      
      }
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
    }else{ 
    ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
    }

      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(1000, Indicator.Text);
    }
      var entries = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
      WorkspaceUtils.waitForObj(entries);
      entries.Click();
      aqUtils.Delay(100, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }
      var entries = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.TabControl2;
      WorkspaceUtils.waitForObj(entries);
      entries.Click();
      aqUtils.Delay(100, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      } 
      var add = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
      WorkspaceUtils.waitForObj(add);
      add.Click();
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      } 
      var EntryGrid = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
      Sys.HighlightObject(EntryGrid);
      var Emp = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
      WorkspaceUtils.waitForObj(Emp);
      Emp.Click();
      WorkspaceUtils.SearchByValue(Emp,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Employee").OleValue.toString().trim(),EmpNo,"Employee Number");
      Sys.Desktop.KeyDown(0x09); // Press Ctrl
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyDown(0x09); 
                
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      Sys.Desktop.KeyUp(0x09);
                
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      var Qty = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget3;
      Qty.setText(temp[2]);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      var billable = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget3
      billable.setText(temp[3]);
      aqUtils.Delay(4000, Indicator.Text);
      var save = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
//      var save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
      WorkspaceUtils.waitForObj(save);
      save.Click();
      aqUtils.Delay(100, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }
      aqUtils.Delay(4000, Indicator.Text);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);
      Sys.Desktop.KeyDown(0x09);
      aqUtils.Delay(1000, Indicator.Text);
      Sys.Desktop.KeyUp(0x09);

      aqUtils.Delay(100, Indicator.Text);
      var allocate = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
      WorkspaceUtils.waitForObj(allocate);
      allocate.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,WorkspaceUtils.Language, "Allocate").OleValue.toString().trim());
      aqUtils.Delay(1000, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }else{ 
      ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
      }
//      var save = Aliases.Maconomy.GlobalVendor.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite3.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
      var save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
      WorkspaceUtils.waitForObj(save);
      save.Click();
      aqUtils.Delay(100, Indicator.Text);
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }
      ImageRepository.ImageSet.Close_Down.Click();
      }
      }
      SelectionBilling.Keys("[Down]");
  }

}