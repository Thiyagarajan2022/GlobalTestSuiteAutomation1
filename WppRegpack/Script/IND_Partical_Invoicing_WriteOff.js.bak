//USEUNIT ReportUtils
//USEUNIT TextUtils
//USEUNIT WorkspaceUtils
//USEUNIT Partial_invoicing_WriteOff

function WriteOff(writeOff_Line,SelectionBilling){ 
  for(var i=0;i<writeOff_Line.length;i++)
  Log.Message(writeOff_Line[i])
  ImageRepository.ImageSet.Maximize1.Click();
    for(var t=0;t<SelectionBilling.getItemCount();t++){ 
      var match = true;
      for(var z=0;z<writeOff_Line.length;z++){
        var temp = writeOff_Line[z].split("*");
    if(SelectionBilling.getItem(t).getText_2(1).OleValue.toString().trim().indexOf(temp[0])!=-1){ 
      if(match){
      var entries = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.PTabItemPanel.TabControl;
      WorkspaceUtils.waitForObj(entries);
      entries.Click();
      aqUtils.Delay(100, Indicator.Text);
      }
  match = false;
            if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
            }else{ 
            ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
            }
            
                var add = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
                WorkspaceUtils.waitForObj(add);
                add.Click();
                var EmpNo = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
                WorkspaceUtils.waitForObj(EmpNo);
                EmpNo.Click();
                var EmployeeNum = Partial_invoicing_WriteOff.EmpNo;
                if(temp[0].indexOf("T")==0){
                if((EmployeeNum!="")&&(EmployeeNum!=null)){ 
                WorkspaceUtils.SearchByValue(EmpNo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Partial_invoicing_WriteOff.Language, "Employee").OleValue.toString().trim(),EmployeeNum,"Employee Number :"+EmployeeNum);
                }
                }
                Sys.Desktop.KeyDown(0x09); // Press Ctrl
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyDown(0x09);
//                var workCode = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
//                workCode.HoverMouse();
//                workCode.Click();
//                WorkspaceUtils.SearchByValue(workCode,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Partial_invoicing_WriteOff.Language, "Work Code").OleValue.toString().trim(),temp[0],"Work Code :"+temp[0]);
                
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyUp(0x09);
//                var desc = Aliases.Maconomy.JobInvoiceAllocation_wip.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget2;
//                WorkspaceUtils.waitForObj(desc);
//                desc.Click();
//                desc.setText(split_text[1])
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);

                var qty = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
                WorkspaceUtils.waitForObj(qty);
                qty.Click();
                qty.setText(temp[2])
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                var unitprice = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McTextWidget;
                WorkspaceUtils.waitForObj(unitprice);
                unitprice.Click();
                unitprice.setText(temp[3])
//                Sys.Desktop.KeyDown(0x09);
//                aqUtils.Delay(1000, Indicator.Text);
//                Sys.Desktop.KeyUp(0x09);
                var save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
Log.Message(temp[5])
Log.Message(temp[5]=="Write-Off");

                if(temp[5]=="Write-Off"){
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);

                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var allocate = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
                WorkspaceUtils.waitForObj(allocate);
                if(temp[5]=="Write-Off"){
                allocate.Keys("Write Off");
                }else{ 
                allocate.Keys("Invoice");
                }
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }

}//Write-off
                }
                }
              ImageRepository.ImageSet.Close_Down.Click();
              Sys.HighlightObject(SelectionBilling);
      
      SelectionBilling.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
      var hsn = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget2;
      hsn.Click();
      WorkspaceUtils.SearchByValue(hsn,"Local Specification 8",temp[4],"HSN");
      if(temp[0].indexOf("T")==0){
      hsn.Keys("[Tab]");
      var Employee = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget2;
      WorkspaceUtils.waitForObj(Employee);
      Employee.Click();
      if((EmployeeNum!="")&&(EmployeeNum!=null)){ 
      Employee.HoverMouse();
      Employee.Click();
      WorkspaceUtils.SearchByValue(Employee,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Partial_invoicing_WriteOff.Language, "Employee").OleValue.toString().trim(),EmployeeNum,"Employee Number :"+EmployeeNum);
      }
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }
      
      var Save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.TabFolderPanel.Composite2;
      WorkspaceUtils.waitForObj(Save);
      for(var i=0;i<Save.ChildCount;i++){ 
        if((Save.Child(i).isVisible())&&(Save.Child(i).toolTipText=="Save Invoice Selection Line")){
          Save = Save.Child(i);
          ReportUtils.logStep_Screenshot("");
          Save.Click();
          break;
        }
      }
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
      aqUtils.Delay(100, Indicator.Text);
      if(temp[0].indexOf("T")==0){
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      }
     
      SelectionBilling.Keys("[Down]");
    

  }
  
  var closeSelection = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
  WorkspaceUtils.waitForObj(closeSelection);
  closeSelection.Click();
  
}