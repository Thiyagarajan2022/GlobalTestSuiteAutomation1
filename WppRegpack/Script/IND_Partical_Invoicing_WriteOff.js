//USEUNIT ReportUtils
//USEUNIT TextUtils
//USEUNIT WorkspaceUtils
//USEUNIT Partial_invoicing_WriteOff

function WriteOff(writeOff_Line,SelectionBilling){
  var seperate = [];
  var uniqueData = [];
  var match = true;
  
    ImageRepository.ImageSet.Maximize1.Click();
    
    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
      
    }
    
    
    for(var t=0;t<SelectionBilling.getItemCount();t++){ 
    var WriteOff_Match = false;
    Log.Message(SelectionBilling.getItem(t).getText_2(1).OleValue.toString().trim())
    for(var z=0;z<writeOff_Line.length;z++){
    var temp = writeOff_Line[z].split("*");
    if(SelectionBilling.getItem(t).getText_2(1).OleValue.toString().trim().indexOf(temp[0])!=-1){
      Log.Message(temp[0])
    WriteOff_Match = true;
    match = true;
    }
    }
    
    if(WriteOff_Match){ 
      
      if(match){
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
      }
        match = false;
      if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
      }
      
      var add = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
      WorkspaceUtils.waitForObj(add);
                
      var EntryGrid = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
      Sys.HighlightObject(EntryGrid);
      for(var s=0;s<EntryGrid.getItemCount()-1;s++){ 
      var EntStatus = true;
      
                var EmpNo = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
                WorkspaceUtils.waitForObj(EmpNo);
                EmpNo.Click();
                var EmployeeNum = Partial_invoicing_WriteOff.EmpNo;
                if(temp[0].indexOf("T")==0){
                if((EmployeeNum!="")&&(EmployeeNum!=null)){ 
                WorkspaceUtils.SearchByValue(EmpNo,JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Partial_invoicing_WriteOff.Language, "Employee").OleValue.toString().trim(),EmployeeNum,"Employee Number :"+EmployeeNum);
                }
                }
                
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }
                
                for(var Tab_Move =0;Tab_Move<8;Tab_Move++){ 
                Sys.Desktop.KeyDown(0x09); 
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                }
  /*
                Sys.Desktop.KeyDown(0x09); // Press Ctrl
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyDown(0x09); 
                
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyUp(0x09);
                
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyDown(0x09);
                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
*/
                aqUtils.Delay(1000, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }
                var allocate = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
                WorkspaceUtils.waitForObj(allocate);
                allocate.Keys(JavaClasses.MLT.MultiLingualTranslator.GetTransText(Project.Path,Partial_invoicing_WriteOff.Language, "Write Off").OleValue.toString().trim());
                aqUtils.Delay(1000, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }
                
                var save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }
       
      for(var Tab_Move =0;Tab_Move<8;Tab_Move++){          
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(1000, Indicator.Text);
      }

/*
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
*/
      
      if(EntStatus){
      var EntryGrid = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
      Log.Message(s)
      Log.Message(EntryGrid.getItemCount())
      Log.Message(EntryGrid.getItemCount()-1)
      if(s<EntryGrid.getItemCount()-2){
      EntryGrid.Keys("[Down]");
      }        
      }  
      
      if(s==EntryGrid.getItemCount()-2){
      break;
      }
      
      }
      
//Closing Entries Tab ( Bottom Window )
    ImageRepository.ImageSet.Close_Down.Click();
Sys.HighlightObject(SelectionBilling);

      
      SelectionBilling.Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]");
      var hsn = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget2;
      hsn.Click();
      //WorkspaceUtils.SearchByValue(hsn,"Local Specification 8",temp[2],"HSN");
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

      for(var Tab_Move =0;Tab_Move<8;Tab_Move++){ 
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(1000, Indicator.Text);
      }
      
/*
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
      
      */

      if(temp[0].indexOf("T")==0){
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
      }
  
      
    }
//    else{ 
      aqUtils.Delay(100, Indicator.Text);
      SelectionBilling.Keys("[Down]");  
      
//    }
    
    }
    
      var closeSelection = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
      WorkspaceUtils.waitForObj(closeSelection);
      closeSelection.Click();
    }
    
