//USEUNIT ReportUtils
//USEUNIT TextUtils
//USEUNIT WorkspaceUtils
//USEUNIT Partial_invoicing_WriteOff

function WriteOff(writeOff_Line,SelectionBilling){
  var seperate = [];
  var uniqueData = [];
  for(var i=0;i<writeOff_Line.length;i++){
    var temp = writeOff_Line[i].split("*");
    seperate[i] = temp[0];
  }
  uniqueData = Array.from(new Set(seperate));
  for(var i=0;i<uniqueData.length;i++){
  Log.Message(uniqueData[i])
  }
  
  
    ImageRepository.ImageSet.Maximize1.Click();
    for(var t=0;t<SelectionBilling.getItemCount();t++){ 
    var match = true;
    var W_Off = [];
    var w =0;
    W_Off[0] = "***";
    var InvoiceData = [];
    var i=0;
    Log.Message(SelectionBilling.getItem(t).getText_2(1).OleValue.toString().trim())
      for(var z=0;z<writeOff_Line.length;z++){
        var temp = writeOff_Line[z].split("*");
    if(SelectionBilling.getItem(t).getText_2(1).OleValue.toString().trim().indexOf(temp[0])!=-1){
      InvoiceData[i] = writeOff_Line[z];
      Log.Message(InvoiceData[i])
      i++;
    }
    }
    if(InvoiceData.length>0){
      for(var z=0;z<InvoiceData.length;z++){
        var temp = InvoiceData[z].split("*");
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
    
      }else{ 
      ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
      }   
      var tot = parseFloat(temp[1]);
      tot = parseFloat(tot.toFixed(2));
      
                var add = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl;
                WorkspaceUtils.waitForObj(add);
                
                
                var EntryGrid = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
                Sys.HighlightObject(EntryGrid);
                for(var s=0;s<EntryGrid.getItemCount()-1;s++){ 
                  var EntStatus = false;
                for(var y=0;y<InvoiceData.length;y++){
                var Etemp = InvoiceData[y].split("*");
                Log.Message(temp[0]+" : "+Etemp[0])
                Log.Message(temp[0] == Etemp[0])
                Log.Message(Etemp[5])
                Log.Message(Etemp[5]== "Write Off")
                if((temp[0] == Etemp[0])&&(Etemp[5]== "Write Off")){ 
                  var EQanty = EntryGrid.getItem(s).getText_2(3).OleValue.toString().trim().replace(/,/g, '');
                  var EUnitP = EntryGrid.getItem(s).getText_2(6).OleValue.toString().trim().replace(/,/g, '');
Log.Message(Etemp[2])
if(Etemp[2]!=""){
                  var WQty = parseFloat(Etemp[2]);
                  WQty = parseFloat(WQty.toFixed(2));
                  var WUnitP = parseFloat(Etemp[3]);
                  WUnitP = parseFloat(WUnitP.toFixed(2));
//                  Log.Message(EQanty+"  :"+WQty)
//                  Log.Message(EQanty==WQty)
//                  Log.Message(EUnitP+"  :"+WUnitP)
//                  Log.Message(EUnitP==WUnitP)

                  if((EQanty==WQty)&&(EUnitP==WUnitP)){ 
                    W_Off[w] = InvoiceData[y];
                    w++;
                    aqUtils.Delay(1000, Indicator.Text);
                    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                    }
                    entries.Click();
                    EntStatus = true;
//      /*
                var EmpNo = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
                WorkspaceUtils.waitForObj(EmpNo);
                EmpNo.Click();
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

                aqUtils.Delay(100, Indicator.Text);
                var allocate = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
                WorkspaceUtils.waitForObj(allocate);
                allocate.Keys("Write Off");
                aqUtils.Delay(1000, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                aqUtils.Delay(100, Indicator.Text);
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
      aqUtils.Delay(100, Indicator.Text);
      
//      */
      
//                    if(EntStatus){
//                    Log.Message(s)
//                    Log.Message(EntryGrid.getItemCount())
//                    Log.Message(EntryGrid.getItemCount()-1)
//                    if(s<EntryGrid.getItemCount()-2){
//                    EntryGrid.Keys("[Down]");
//                    }
//                   }       
break;
      }
      else{ 


      }
//---------Today Chnages------------------
}else{ 
  
                    W_Off[w] = InvoiceData[y];
                    w++;
                    aqUtils.Delay(1000, Indicator.Text);
                    if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                    }
                    entries.Click();
                    EntStatus = true;
//      /*
                var EmpNo = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McValuePickerWidget;
                WorkspaceUtils.waitForObj(EmpNo);
                EmpNo.Click();
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

                aqUtils.Delay(100, Indicator.Text);
                var allocate = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid.McPopupPickerWidget;
                WorkspaceUtils.waitForObj(allocate);
                allocate.Keys("Write Off");
                aqUtils.Delay(1000, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }
                var save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl3;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                aqUtils.Delay(100, Indicator.Text);
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
      aqUtils.Delay(100, Indicator.Text);
      
//      */
      
//                    if(EntStatus){
//                    Log.Message(s)
//                    Log.Message(EntryGrid.getItemCount())
//                    Log.Message(EntryGrid.getItemCount()-1)
//                    if(s<EntryGrid.getItemCount()-2){
//                    EntryGrid.Keys("[Down]");
//                    }
//                   }       
break;



}
//---------------------------------
                }  
               } 
                    if(EntStatus){
                    Log.Message(s)
                    Log.Message(EntryGrid.getItemCount())
                    Log.Message(EntryGrid.getItemCount()-1)
                    if(s<EntryGrid.getItemCount()-2){
                    EntryGrid.Keys("[Down]");
                    }
//                    break;         
                   }              
  }
                
//===========
                var m_writeOff = false;
                Log.Message(W_Off.length)
                Log.Message(m_writeOff)
                 for(var zy=0;zy<W_Off.length;zy++){
                   Log.Message("==Checking in List =========")
                   Log.Message(W_Off[zy] +" : "+InvoiceData[z])
                   Log.Message(W_Off[zy]==InvoiceData[z])
                   if(W_Off[zy]==InvoiceData[z]){ 
                     m_writeOff = true;
                   }
                   }
                   Log.Message(m_writeOff)
                if(!m_writeOff){
                
                
Log.Message("New Line")
Log.Message(InvoiceData[z])
                

                add.Click();
                aqUtils.Delay(100, Indicator.Text);
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
                  
                }
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

                aqUtils.Delay(1000, Indicator.Text);
                Sys.Desktop.KeyUp(0x09);
                Sys.Desktop.KeyUp(0x09);

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
                var save = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite2.Composite.PTabFolder.TabFolderPanel.Composite.SingleToolItemControl2;
                WorkspaceUtils.waitForObj(save);
                save.Click();
                if(ImageRepository.ImageSet.Tab_Icon.Exists()){ 
    
                }else{ 
                ValidationUtils.verify(true,false,"Maconomy is loading continously......")  
                }

// Write Off or Carry Forward or Invocie
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
                if(temp[5]=="Write Off"){
                allocate.Keys("Write Off");
                }else if(temp[5]=="Carry Forward"){
                allocate.Keys("Carry Forward");
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
      Sys.Desktop.KeyDown(0x10)
      Sys.Desktop.KeyDown(0x09)
      Sys.Desktop.KeyUp(0x10)
      Sys.Desktop.KeyUp(0x09)
      aqUtils.Delay(100, Indicator.Text);
//    break;
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
     
}
SelectionBilling.Keys("[Down]");      
    }
    
  var closeSelection = Aliases.Maconomy.Invoicing_WriteOff.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabItemPanel2.TabControl;
  WorkspaceUtils.waitForObj(closeSelection);
  closeSelection.Click();
}










