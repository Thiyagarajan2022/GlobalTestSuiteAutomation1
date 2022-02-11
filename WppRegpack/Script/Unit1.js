//USEUNIT ExcelUtils
//USEUNIT WorkspaceUtils
function exlApprove() {
	var Approve_Level = [];
	var ApproveInfo = [];
	var level =0;

		 
//		Approve_Level[0] = "1307*1307200440*1307 Senior Finance (13079510)*1307 Management (13079507)";
//		Approve_Level[1] = "1307*1307100030*1307 Senior Finance (13079510)*muthu*";
//		Approve_Level[2] = "1307*1307100030*1307 Senior Finance (13079510)*1307 Billers*";
//		CredentialLogin();

		Approve_Level[0] = "1307 Senior Finance*2*1307 Management*3*";
		Approve_Level[1] = "1307 Biller*2*1307 Senior Finance*3";
		Approve_Level[2] = "1307 Management*2*1307 Biller*3*";
		
		Approve_Level[0] = "1307 Senior Finance*2*1307 Management*3*";
		Approve_Level[1] = "1307 Senior Finance*3*";
		Approve_Level[2] = "1307 Biller*2*1307 Management*3*";

		var uniqueMaconomyData = [];
		var ss =[];
    var s=0;
		for(var i=0;i<3;i++){ 
			var temp = Approve_Level[i].split("*");
			if(temp.length==2){
				Log.Message("   "+temp[0]);
				ss[s]=temp[0];
        s++;
        Log.Message("ss :"+temp[0]);
			}
		}
    for(var i=0;i<ss.length;i++){
      Log.Message(ss[i]);
      }
    uniqueMaconomyData = Array.from(new Set(ss));
		
		for(var i=0;i<ss.length;i++){ 
			for(var j=i+1;j<ss.length;j++){ 
				if(ss[i]==ss[j]){ 
					Log.Message("Two levels has same Approver Logins and Substitute is not there in Opco Approver or SSC Approvers List");
				}
			}
		}
		
		for(var i=0;i<Approve_Level.length;i++){
			var temp = Approve_Level[i].split("*");
			if(temp.length!=2){
			for(var j=0;j<ss.length;j++){
				if(Approve_Level[i].indexOf(ss[j].toString())!=-1){ 
					if(temp[0]==ss[j]){
						Approve_Level[i] = temp[2]+"*"+temp[3]+"*";	
					}
					if(temp[2]==ss[j]){
						Approve_Level[i] = temp[0]+"*"+temp[1]+"*";	
					}
				}
			}
		}
		}
		
		for(var i=0;i<Approve_Level.length;i++){ 
			var temp = Approve_Level[i].split("*");
			for(var j=i+1;j<Approve_Level.length;j++){ 
				var temp1 = Approve_Level[j].split("*");
			if((temp.length==2)&&(temp1.length==2)&&(temp[0]==temp1[0])){
				Log.Message("Two levels has same Approver Logins and Main Approver (or) Substitute is not there in Approvers List");
			}
			}
		}
		
	
		var ApproveInfo = Approve_Level;
		for(var i=0;i<ApproveInfo.length;i++){
			var temp = ApproveInfo[i].split("*");
      Log.Message(temp.length)
			if(temp.length>3)
				for(var j=i+1;j<ApproveInfo.length;j++){
//			if(ApproveInfo[j]==temp[0]){ 
				var temp1 = ApproveInfo[j].split("*");
				if(temp1.length>3){
          Log.Message("temp[0] :"+temp[0]);
          Log.Message("temp[2] :"+temp[2]);
          Log.Message("temp1[0] :"+temp1[0]);
          Log.Message("temp1[2] :"+temp1[2]);
				if((temp[0]==temp1[0]) && (temp[2]==temp1[2])){ 
					ApproveInfo[i] = temp[0]+"*"+temp[1]+"*";
					ApproveInfo[j] = temp1[2]+"*"+temp[3]+"*";
					Log.Message("ApproveInfo["+i+"] = "+temp[0]+"*"+temp[1]+"*");
          Log.Message("ApproveInfo["+j+"] = "+temp1[2]+"*"+temp[3]+"*");
				}
				if((temp[0]==temp1[2]) && (temp[2]==temp1[0])){ 
					ApproveInfo[i] = temp[0]+"*"+temp[1]+"*";
					ApproveInfo[j] = temp1[0]+"*"+temp[1]+"*";
					Log.Message("ApproveInfo["+i+"] = "+temp[0]+"*"+temp[1]+"*");
          Log.Message("ApproveInfo["+j+"] = "+temp1[0]+"*"+temp[1]+"*");
				}
				
				if((temp[0]==temp1[2])&&(!(temp[2]==temp1[0]))){ 
					ApproveInfo[i] = temp[0]+"*"+temp[1]+"*";
					ApproveInfo[j] = temp1[0]+"*"+temp[1]+"*";
					Log.Message("ApproveInfo["+i+"] = "+temp[0]+"*"+temp[1]+"*");
          Log.Message("ApproveInfo["+j+"] = "+temp1[0]+"*"+temp[1]+"*");          
				}
				if((temp[0]==temp1[0])&&(!(temp[2]==temp1[2]))){ 
					ApproveInfo[i] = temp[0]+"*"+temp[1]+"*";
					ApproveInfo[j] = temp1[2]+"*"+temp[3]+"*";
					Log.Message("ApproveInfo["+i+"] = "+temp[0]+"*"+temp[1]+"*");
          Log.Message("ApproveInfo["+j+"] = "+temp1[2]+"*"+temp[3]+"*");          
				}
				
				}

//			}
			
		}
		
				}
		Log.Message("Output :");
		for(var i=0;i<ApproveInfo.length;i++){
			Log.Message(ApproveInfo[i]);
		}
		
		
		Log.Message("Final Output :");
		for(var i=0;i<Approve_Level.length;i++){
			Log.Message(Approve_Level[i]);
		}
		}
		
	


//	public static void CredentialLogin1() throws IOException{
//		String[] Crd = new String[10];
//		int z =0;
//		
//		
//		for(int i=level;i<Approve_Level.length;i++){	
//		  String[] Cred = Approve_Level[i].split("\\*");
//		  for(int j=2;j<4;j++){
//		  if((Cred[j]!="")&&(Cred[j]!=null)){
//			  Crd[z] = Cred[j]; 
//			  Log.Message(Crd[z]);
//			  z++;
//		  }
//		  }
//		}
//		
////		Set<String> uniq = new LinkedHashSet<String>();
////		Set<String> dups = new LinkedHashSet<String>();
////		for(String s:Crd){ 
////			uniq.add(s);
////		}
////		Log.Message();
////		Log.Message();
////		Iterator<String> itr = uniq.iterator();
////		while(itr.hasNext()){ 
////			Log.Message(itr.next());
////		}
//		
//	for(int i=level;i<Approve_Level.length;i++){
//
//	  String temp="";
//	  String[] Cred = Approve_Level[i].split("\\*");
//	  Log.Message(" ");
//	  Log.Message(" ");
//	  Log.Message(" A");
//	  for(int j=2;j<4;j++){
//	  if((Cred[j]!="")&&(Cred[j]!=null))
//	  if((Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf("1307 ")!=-1)))
//	  { 
//
//		  if((Cred[j].indexOf("(")!=-1)&&(Cred[j].indexOf(")")!=-1))
//			  temp = temp+Cred[j].substring(0,Cred[j].indexOf("(")-1);
//	  }
//	  
//	  }
//	  Crd[z] = temp;
//	  Log.Message(temp);
//	  z++;
//	}
//
//	}
//
//
//public static void CredentialLogin() throws IOException{
//	String[] Crd = new String[10];
//	int z =0;
//for(int i=level;i<Approve_Level.length;i++){
//  boolean UserN = true;
//
//  String temp1="";
//  String[] Cred = Approve_Level[i].split("\\*");
//  Log.Message(" ");
//  Log.Message(" ");
//  Log.Message(" ");
//  for(int j=2;j<4;j++){
//	  String temp="";
//  if((Cred[j]!="")&&(Cred[j]!=null))
//  if((Cred[j].indexOf("SSC - ")==-1)&&(Cred[j].indexOf("Central Team - Client Management")==-1) &&(Cred[j].indexOf("Central Team - Vendor Management")==-1) && ((Cred[j].indexOf("OpCo - ")!=-1) || (Cred[j].indexOf("1307 ")!=-1)))
//  { 
//
//     String sheetName = "Agency Users";
//     String workBook = "D:\\GlobalTestPack\\GlobalTestPack\\WppRegPack\\TestResource\\Regression\\DS_CHN_REGRESSION.xlsx";
//    temp = AgencyLogin(Cred[j],"1307");
////    Log.Message(Cred[j] +" 1307");
//  }
////  else if((Cred[j].indexOf("SSC - ")!=-1)||(Cred[j].indexOf("Central Team - Vendor Management")!=-1) ||(Cred[j].indexOf("Central Team - Client Management")!=-1))
////  { 
////
////    var sheetName = "SSC Users";
////    ExcelUtils.setExcelName(workBook, sheetName, true);
////    temp = ExcelUtils.SSCLogin(Cred[j],"Username");
////  }
//
////  Log.Message("D :"+temp);
//  if(temp.length()!=0){
//	    temp1 = temp1+temp+"*"+j+"*";
////	    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;
////	    Log.Message(Cred[0]+"*"+Cred[1]+"*"+temp);
//	  }
//  }
//  Log.Message(temp1);
//  if(temp1.length()!=0){
//  Crd[z] =temp1;
//  z++;
//  }else{ 
//	  Log.Message("Login is not Found");
//  }
//  
//  
//  
//  
////  if(temp.length()!=0){
//////	    ApproveInfo[i] = Cred[0]+"*"+Cred[1]+"*"+temp;
//////	    Log.Message(Cred[0]+"*"+Cred[1]+"*"+temp);
////	  }
//}
//Set<String> uniq = new LinkedHashSet<String>();
//Set<String> dups = new LinkedHashSet<String>();
//for(int j=0;j<Crd.length;j++){ 
//	if(Crd[j]!=null){
//	String [] temp2 = Crd[j].split("\\*");
//	for(int k=0;k<temp2.length;k=k+2){ 
//	uniq.add(temp2[k]);
////	Log.Message(temp2[k]);
//	}
//	}
//}
//Log.Message();
//Log.Message();
//Iterator<String> itr = uniq.iterator();
//while(itr.hasNext()){ 
//	Log.Message(itr.next());
//}
//
//}
//
//
//
//
//
//
//
//public static String AgencyLogin(String rowidentifier, String column) throws IOException
//{
//	String sheetName = "Agency Users";
//    String workBook = "D:\\GlobalTestPack\\GlobalTestPack\\WppRegPack\\TestResource\\Regression\\DS_CHN_REGRESSION.xlsx";
//	FileInputStream fis = new FileInputStream(new File(workBook));
//	XSSFWorkbook outWorkbook = new XSSFWorkbook(fis);
//    XSSFSheet spreadsheet = outWorkbook.getSheet(sheetName); // Sheet Name
//      
////var xlDriver = DDT.ExcelDriver(excelName,sheet,true);
//int id =0;
//String temp ="";
//if(rowidentifier.indexOf("OpCo -")!=-1){ 
//  rowidentifier = rowidentifier.replaceAll("OpCo -",column);
//  }
//if(rowidentifier.indexOf("Billers")!=-1)
//    rowidentifier = rowidentifier.replace("Billers","Biller");
//  
//if((rowidentifier.indexOf("(")!=-1)&&(rowidentifier.indexOf(")")!=-1))
//    rowidentifier = rowidentifier.substring(0,rowidentifier.indexOf("(")-1);
////   Log.Message("rowidentifier :"+rowidentifier);
//int noOfrow = spreadsheet.getLastRowNum();
//XSSFRow row = spreadsheet.getRow(0);
//CellStyle style = outWorkbook.createCellStyle();
//int ColumnCount = row.getLastCellNum();
//
//for(int idCol =0;idCol<ColumnCount;idCol++){
//	XSSFCell cell = row.getCell(idCol);
//	if (cell.getCellType() == cell.CELL_TYPE_NUMERIC) {
//	if(cell.getNumericCellValue()==Integer.parseInt(column)){
////		Log.Message("Numeric :"+cell.getNumericCellValue());
//		for(int idRow =1;idRow<=noOfrow;idRow++){
//		row = spreadsheet.getRow(idRow);
//		cell = row.getCell(idCol);
////		Log.Message("String :"+cell.getStringCellValue());
//		if(cell.getStringCellValue().equalsIgnoreCase(rowidentifier)){
////		Log.Message("String :"+cell.getStringCellValue());
//			temp = cell.getStringCellValue();
//	}
//		
//		}
//	}
//	}
//	
////	if (cell.getCellType() == cell.CELL_TYPE_STRING) {
////	if(cell.getStringCellValue().equalsIgnoreCase(column)){
////		Log.Message("String :"+cell.getStringCellValue());
////	}
////	}
//}
//
//
////var Col = "";
////for(var i=0;i<DDT.CurrentDriver.ColumnCount;i++){ 
////  if(DDT.CurrentDriver.ColumnName(i).toString().trim().indexOf(column)!=-1)
////  Col = DDT.CurrentDriver.ColumnName(i).toString().trim();
////}
//
////     while (!DDT.CurrentDriver.EOF()) {
////       if(xlDriver.Value(Col).toString().trim().indexOf(rowidentifier.toString().trim())!=-1){
////        try{
////         temp = temp+xlDriver.Value(Col).toString().trim();
////         }
////        catch(e){
////        temp = "";
////        }
////      break;
////      }
////
////    xlDriver.Next();
////     }
////     DDT.CloseDriver(xlDriver.Name);
//     return temp;
//
//}


function vv(){ 
  var Email = "1707_TestEmployeeAutomation@gmail.com";
  
  var Eml_split1 = Email.substring(0,Email.indexOf("@"));
var Eml_split2 = Email.substring(Email.indexOf("@"));
Eml_split1 = Eml_split1 +" "+StartTime();
Eml_split1 = Eml_split1.replace(/[_: ]/g,"");
Email = Eml_split1+Eml_split2;
Log.Message(Email);
//Email_1.setText(Email); 
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
var add = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").Child(kk).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("TabFolderPanel", "", 1).SWTObject("Composite", "", 1).SWTObject("SingleToolItemControl", "", 4);;;
Sys.HighlightObject(add);
}

//.Enabled
function bn(){ 

var table = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 Senior Finance (TSTAUTO)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2);
Sys.HighlightObject(table);
Log.Message(table.getItemCount())
var name = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 Senior Finance (TSTAUTO)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McTableWidget", "", 3).SWTObject("McGrid", "", 2).SWTObject("McTextWidget", "", 2);
name.setText("Auto*");

var result = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1707 Senior Finance (TSTAUTO)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 2).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McWorkspaceSheafGui$McDecoratedPaneGui", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McFilterPaneWidget", "").SWTObject("McPagingWidget", "", 2).SWTObject("Label", "*");
var i=0;

while((result.getText().OleValue.toString().trim().indexOf("results")==-1)&&(i!=60)){ 
  aqUtils.Delay(100);
  i++;
  result.Refresh();
}

Log.Message(table.getItemCount())
}

function bv(){ 
//  var workBook = "C:\\Users\\674087\\Documents\\BAU\\GlobalTestSuiteAutomation\\WppRegpack\\TestResource\\Regression\\DS_IND_REGRESSION.xlsx";
//  var sheetName = "JobCreation";
//  var jd = "ValidateJobtype_CB(Client Billable)";
//  var Opco = "1707";
//ExcelUtils.setExcelName(workBook, sheetName, true);
//FullName = ExcelUtils.getRowDatas("ValidateJobtype_CB(Client Billable)","1707");
////FullName = ExcelUtils.getRowDatas("ValidateJobtype_CB(Client billable)","1707");
//Log.Message(FullName);
//
//var excelData = [];
//      if(FullName.indexOf(",")!=-1){
//     excelData =  FullName.split(",");
//     }else if(FullName.length>0){ 
//      excelData[0] = FullName;
//     }
//     
//
// for(var i=0;i<excelData.length;i++)
// Log.Message(excelData[i]);

 
 
 var ApvPerson = Aliases.Maconomy.EmployeeAndUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.McTextWidget;
var ApvPerson = Aliases.Maconomy.EmployeeAndUser.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.PTabFolder.Composite.McClumpSashForm.Composite.McClumpSashForm.Composite.Composite.McPaneGui_10.Composite.Composite.Composite.McGroupWidget.Composite.Composite.Composite.McTextWidget
//var ApvPerson = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 4).SWTObject("Composite", "", 2).SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("McPaneGui$10", "").SWTObject("Composite", "").SWTObject("Composite", "", 1).SWTObject("Composite", "", 1).SWTObject("McGroupWidget", "", 1).SWTObject("Composite", "", 2).SWTObject("Composite", "", 1).SWTObject("McTextWidget", "", 2);
//WorkspaceUtils.waitForObj(ApvPerson);
ApvPerson.Click();
var loginPer = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - *").WndCaption;
    loginPer = loginPer.substring(loginPer.indexOf(" - ")+3);
    var i=0;
    
Log.Message((ApvPerson.getText().OleValue.toString().trim().indexOf("approved")==-1)&&(ApvPerson.getText().OleValue.toString().trim().indexOf("Approved")==-1))
Log.Message((ApvPerson.getText().OleValue.toString().trim().indexOf("By You")==-1)&&(ApvPerson.getText().OleValue.toString().trim().indexOf("by you")==-1)&&(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)==-1))
while ((ApvPerson.getText().OleValue.toString().trim().indexOf("approved")==-1)&&(ApvPerson.getText().OleValue.toString().trim().indexOf("Approved")==-1)&&(ApvPerson.getText().OleValue.toString().trim().indexOf("By You")==-1)&&(ApvPerson.getText().OleValue.toString().trim().indexOf("by you")==-1)&&(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)==-1)&&(i!=600))
{
  aqUtils.Delay(100);
  i++;
  ApvPerson.Refresh();  
}
  if((ApvPerson.getText().OleValue.toString().trim().indexOf("By You")==-1)||(ApvPerson.getText().OleValue.toString().trim().indexOf("by you")==-1)||(ApvPerson.getText().OleValue.toString().trim().indexOf(loginPer)==-1)){
    Log.Checkpoint("Pass")
//  ValidationUtils.verify(true,true,"Created Employee and User is Approved by :"+loginPer)
//  TextUtils.writeLog("Created Employee and User is Rejected by :"+loginPer); 
  }else{ 
    Log.Checkpoint("Fail")
//  TextUtils.writeLog("Created Employee and User is Rejected by :"+loginPer+ "But its Not Reflected"); 
//  ValidationUtils.verify(true,false,"Created Employee and User is Approved by :"+loginPer+ "But its Not Reflected")
  }
}

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



//      PropArray = new Array("JavaClassName", "Index","Visible");
//  ValuesArray = new Array("McGroupWidget", "1", "true");
//  p = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
//  obj = p.FindAll(PropArray, ValuesArray, 1000);
//  Log.Message(obj.length)
//  var newQuote = "";
//for (let i_count = 0; i_count < obj.length; i_count++){ 
//if((obj[i_count].Exists) && (obj[i_count].Parent.JavaClassName=="Composite") && (obj[i_count].Parent.Index==2))
//newQuote = obj[i_count];
//}
//Sys.HighlightObject(newQuote)
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


function waitUntil_MaconomyScreen_loaded_Completely(){ 
//var lneaprovetab = Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
//Sys.HighlightObject(lneaprovetab);
//
//var APGrid =     Aliases.Maconomy.Shell.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite.Composite5.Composite.PTabFolder.Composite.McClumpSashForm.Composite.Composite.McTableWidget.McGrid;
//Sys.HighlightObject(APGrid);
//var lneaprovetab = getObjectAddress_JavaClasssName_and_Index(Maconomy_ParentAddress,"McGrid",1);
//Sys.HighlightObject(lneaprovetab);

var Maconomy_ParentAddress = Sys.Process("Maconomy").SWTObject("Shell", "Deltek Maconomy - 1221 Finance (TST)")

      PropArray = new Array("JavaClassName", "Index","Visible");
  ValuesArray = new Array("SingleToolItemControl",3, "true");
  p = eval(Maconomy_ParentAddress).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "");
  obj = p.FindAll(PropArray, ValuesArray, 1000);
  Log.Message(obj.length)
  var quteNumber = "";
for (let i_count = 0; i_count < obj.length; i_count++){ 
  Log.Message(obj[i_count].toolTipText)
if(obj[i_count].toolTipText.OleValue.toString().trim().indexOf("Save Invoice Line")!=-1){
quteNumber = obj[i_count];
}
}
Sys.HighlightObject(quteNumber)

Log.Message(quteNumber.FullName);

}
