package MLT;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.XSSFCell;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class MultiLingualTranslator{
	
	private static String PjtPath="C:\\Users\\274374\\Desktop\\MultiLingualTranslator";
	private static String Lang;
	private static String CvtWord;
	private static CharSequence str;

	
	public static void main(String args[]) throws SQLException, IOException {
	
		
	//	AA();
		// String a  = GetTransText("C:\\Users\\snaray01\\eclipse-workspace\\Ebox", "German", "Customers");
		// System.out.print(a);
		 String Text = "";
		 if(args.length == 2)
		 {
		 Text = GetTransText(args[0], args[1]);
		System.out.println(Text);
		 }
//		 
//		 else if(args.length == 4)
//		 {
//		//Comparator(String ProjectPath, String Language, String[] Arr1, String[] Arr2)
//		//	ComparatorNonArray(args[0], args[1], args[2],args[3]);
//			
//		 System.out.println("Comparision over.. Check Reports..");
//		 }
		
		 
	}


	public static String GetTransText(String Language, String ConvertWord)
			throws SQLException, IOException {
		String ConvertedText = "";
		PjtPath = "C:\\Users\\274374\\Desktop";
		// Lang = Language;
		CvtWord = ConvertWord;
		File dir = new File(PjtPath + "\\MultiLingualTranslator");

		if (!dir.exists()) {
			GetMultiLingualTranslatorFolder(dir);
		}
		PjtPath = PjtPath + "\\MultiLingualTranslator";

		dir = new File(PjtPath + "\\node_modules");

		if (!dir.exists()) {
			GetNodeModules();
		}

		dir = new File(PjtPath + "\\LanguageDB.accdb");

		if (!dir.exists()) {
			// GetLanguageDB(dir);
			System.out.println("LanguageDB not exists.. Kindly place the required DB in the ProjectPath");
			System.exit(0);
		}
		
		
		Lang = ConnectDB(" Select Code from [LanguageCode] where Language = '" + Language + "'");
		// SELECT Code FROM LanguageCode where Language='German'

		// System.out.println(Lang);
		if (Language != "English") {
			String PresentText = ConnectDB("Select LangText from [" + Language + "] where EngText = '" + CvtWord + "'");

			if (PresentText == "") {
			//	MakeTranslateJs(CvtWord, Lang);
				ConvertedText = RunTranslateJS(Lang,CvtWord);
				ConnectDB("Insert into [" + Language + "](Flag,EngText,LangText,LangText1) values (0,'" + CvtWord + "','"
						+ ConvertedText + "','')");
			} else {

				// PresentText = ConnectDB("Select LangText from "+Language+" where EngText =
				// '"+CvtWord + "'" );
				
				if(ConnectDB("Select Flag from [" + Language + "] where EngText = '" + CvtWord + "'").equals("0"))
				{
				ConvertedText = PresentText;
				}
				else
				{
					ConvertedText =	ConnectDB("Select LangText1 from [" + Language + "] where EngText = '" + CvtWord + "'");
				}
			}
		} else {
			//MakeTranslateJs(CvtWord, Lang);
			ConvertedText = RunTranslateJS(Lang,CvtWord);
		}

		return ConvertedText.toString();
	}
	
	
	public static String AAA1(String ProjectPath, String Language, String ConvertWord)
			throws SQLException, IOException {
		String ConvertedText = "";
		PjtPath = ProjectPath;
		// Lang = Language;
		CvtWord = ConvertWord;
		File dir = new File(PjtPath + "\\MultiLingualTranslator");

		if (!dir.exists()) {
			GetMultiLingualTranslatorFolder(dir);
		}
		PjtPath = PjtPath + "\\MultiLingualTranslator";

		dir = new File(PjtPath + "\\node_modules\\google-translate-api");

		if (!dir.exists()) {
			GetNodeModules();
		}

		dir = new File(PjtPath + "\\LanguageDB.accdb");

		if (!dir.exists()) {
			// GetLanguageDB(dir);
			System.out.println("LanguageDB not exists.. Kindly place the required DB in the ProjectPath");
			System.exit(0);
		}
		
		
		Lang = ConnectDB(" Select Code from [LanguageCode] where Language = '" + Language + "'");
		// SELECT Code FROM LanguageCode where Language='German'

		// System.out.println(Lang);
		if (Language != "English") {
			String PresentText = ConnectDB("Select LangText from [" + Language + "] where EngText = '" + CvtWord + "'");

			if (PresentText == "") {
				//MakeTranslateJs(CvtWord, Lang);
				ConvertedText = RunTranslateJS(Lang,CvtWord);
				ConnectDB("Insert into [" + Language + "](Flag,EngText,LangText,LangText1) values (0,'" + CvtWord + "','"
						+ ConvertedText + "','')");
			} else {

				// PresentText = ConnectDB("Select LangText from "+Language+" where EngText =
				// '"+CvtWord + "'" );
				
				if(ConnectDB("Select Flag from [" + Language + "] where EngText = '" + CvtWord + "'").equals("0"))
				{
				ConvertedText = PresentText;
				}
				else
				{
					ConvertedText =	ConnectDB("Select LangText1 from [" + Language + "] where EngText = '" + CvtWord + "'");
				}
			}
		} else {
			//MakeTranslateJs(CvtWord, Lang);
			ConvertedText = RunTranslateJS(Lang,CvtWord);
		}
		
		//ConvertedText = ConvertedText + " Semma ";

		return ConvertedText;
	}
	
	
	public static String GetTransTextNonEnglishAlphabets(String ProjectPath, String Language, String ConvertWordByteforArr)
			throws SQLException, IOException {
		
		String ConvertWord = "";
		if(ConvertWordByteforArr.contains(","))
		{
			String[] ConvertWordByteArr = ConvertWordByteforArr.split(",");
			int[] ConvertWordintByteArr = new int[ConvertWordByteArr.length] ;
			for(int z = 0; z < ConvertWordByteArr.length; z++)
			{				
				ConvertWordintByteArr[z] = Integer.parseInt(ConvertWordByteArr[z]);
				ConvertWord = ConvertWord + fromCharCode(ConvertWordintByteArr[z]);
			}
			
		}
		else
		{
			int[] ConvertWordintByteArr = new int[0] ;
			ConvertWordintByteArr[0] = Integer.parseInt(ConvertWordByteforArr);
			ConvertWord = fromCharCode(ConvertWordintByteArr[0]);
		}
		
		
		String ConvertedText = "";
		//String ConvertWord = "";		
		//ConvertWord = new String(ConvertWordArr);
		PjtPath = ProjectPath;
		// Lang = Language;
		CvtWord = ConvertWord;
		File dir = new File(PjtPath + "\\MultiLingualTranslator");

		if (!dir.exists()) {
			GetMultiLingualTranslatorFolder(dir);
		}
		PjtPath = PjtPath + "\\MultiLingualTranslator";

		dir = new File(PjtPath + "\\node_modules\\google-translate-api");

		if (!dir.exists()) {
			GetNodeModules();
		}

		dir = new File(PjtPath + "\\LanguageDB.accdb");

		if (!dir.exists()) {
			// GetLanguageDB(dir);
			System.out.println("LanguageDB not exists.. Kindly place the required DB in the ProjectPath");
			System.exit(0);
		}
		
		
		Lang = ConnectDB(" Select Code from [LanguageCode] where Language = '" + Language + "'");
		// SELECT Code FROM LanguageCode where Language='German'

		// System.out.println(Lang);
		if (Language != "English") {
			String PresentText = ConnectDB("Select LangText from [" + Language + "] where EngText = '" + CvtWord + "'");

			if (PresentText == "") {
			//	MakeTranslateJs(CvtWord, Lang);
				ConvertedText = RunTranslateJS(Lang,CvtWord);
				ConnectDB("Insert into [" + Language + "](Flag,EngText,LangText,LangText1) values (0,'" + CvtWord + "','"
						+ ConvertedText + "','')");
			} else {

				// PresentText = ConnectDB("Select LangText from "+Language+" where EngText =
				// '"+CvtWord + "'" );
				
				if(ConnectDB("Select Flag from [" + Language + "] where EngText = '" + CvtWord + "'").equals("0"))
				{
				ConvertedText = PresentText;
				}
				else
				{
					ConvertedText =	ConnectDB("Select LangText1 from [" + Language + "] where EngText = '" + CvtWord + "'");
				}
			}
		} else {
		//	MakeTranslateJs(CvtWord, Lang);
			ConvertedText = RunTranslateJS(Lang,CvtWord);
		}

	//	ConvertedTextArr = ConvertedText.getBytes();
		
	
		String ConvertedTextByteforArr = fromByteArr(ConvertedText);
		
		return ConvertedTextByteforArr;
	}

	private static void GetMultiLingualTranslatorFolder(File dir) {
		dir.mkdir();
	}

	private static void GetLanguageDB(File dir) {
		try {
			dir.createNewFile();
		} catch (IOException e) {

			e.printStackTrace();
		}
	}

	private static void GetNodeModules() {
		try {
			System.out.println("am coming here");
			ProcessBuilder builder = new ProcessBuilder("cmd.exe", "/c",
					"cd  \"" + PjtPath + "\"&& npm install");
			builder.redirectErrorStream(true);
			Process p = builder.start();
			BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream()));
			String line;
			while (true) {
				line = r.readLine();
				//if (line.contains("'node' is not recognized")) {
				//	System.out.println(
				//			"Kindly install NodeJs from https://nodeJS.org  and if already installed kindly add to environmental variables ");
				//	System.exit(0);
			//	}
				if (line == null) {
					break;
				}

			}
		} catch (IOException e) {

			e.printStackTrace();
		}

	}

	private static void MakeTranslateJs(String Word, String Lang) {
		try {
			File dir = new File(PjtPath + "\\Translate.js");

			if (dir.exists()) {
				dir.delete();
				// System.out.println("Yes deleting!");
			}
			dir.createNewFile();

			String str = "const translate = require('google-translate-api');\r\n";
			str += "translate('" + Word + "', {to: '" + Lang + "'}).then(res => {\r\n";
			str += "  console.log(res.text);\r\n";
			// str+=" console.log(res.from.language.iso);\r\n";
			str += "}).catch(err => {\r\n";
			str += "   console.error(err);\r\n";
			str += "});\r\n";

			String fileName = PjtPath + "\\Translate.js";
			BufferedWriter writer = new BufferedWriter(new FileWriter(fileName));
			writer.write(str);

			writer.close();
		} catch (IOException e) {

			e.printStackTrace();
		}
	}

	private static String ConnectDB(String strQuery) throws SQLException {
		System.out.println(strQuery);
		
	
		Connection conn = DriverManager.getConnection("jdbc:ucanaccess://" + "C:\\Users\\274374\\Desktop\\MultiLingualTranslator\\LanguageDB.accdb");
		Statement s = conn.createStatement();
		String ret = "";
		// ResultSet rs = s.executeQuery("SELECT * FROM [TestTable]");
		// System.out.println(strQuery);
		if (!strQuery.contains("Insert into ")) {
			// System.out.println("In");
			ResultSet rs = s.executeQuery(strQuery);

			while (rs.next()) {
				// System.out.println(rs.getString(1));
				ret = rs.getString(1);
			}
			// System.out.println("Out");
			s.execute(strQuery);

		}
		return ret;
	}

	private static String RunTranslateJS(String lang2, String cvtWord2) throws IOException {
		String line = "";

		ProcessBuilder builder = new ProcessBuilder("cmd.exe", "/c", "cd \"" + PjtPath + "\" && node Translate.js "+ lang2 +" " +cvtWord2);
		System.out.println(builder.command());
		builder.redirectErrorStream(true);
		
	//		System.out.println(builder.);
		
		Process p = builder.start();
		BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream()));

		while (true) {
			line = r.readLine();			
			if (line == null) {
				break;
			}
			// System.out.println(line);
			return line;
		}
		System.out.println(line);
		return line;
	}
	
	public static void ComparatorNonArray(String ProjectPath, String Language, String Arr1, String Arr2)
			throws SQLException, IOException {
		
		String[] Array1 = new String[10];
		String[] Array2 = new String[10];
		if(Arr1.contains(","))
		{
		Array1 = Arr1.split(",");
		}
		else
		{
		Array1 = new String[1];
		Array1[0] = Arr1;
		}
		
		if(Arr2.contains(","))
		{
		Array2 = Arr2.split(",");
		}
		else
		{
		Array2 = new String[1];
		Array2[0] = Arr2;
		}
		
		
		Comparator(ProjectPath,Language, Array1, Array2);
		
	}
	

	public static void Comparator(String ProjectPath, String Language, String[] Arr1, String[] Arr2)
			throws SQLException, IOException {
		//dd/MMM/yyyy 
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("ddmmyyyy-HHmmss");  
		   LocalDateTime now = LocalDateTime.now(); 
		   String LogName = "Comparator_Report_" + dtf.format(now);
		   
		   String StartTag = "<!DOCTYPE html>\r\n" + 
		   		"<html>\r\n" + 
		   		"<head>\r\n" + 
		   		"<style>\r\n" + 
		   		"table {\r\n" + 
		   		"    border-collapse: collapse;\r\n" + 
		   		"    width: 100%;\r\n" + 
		   		"}\r\n" + 
		   		"\r\n" + 
		   		"th, td {\r\n" + 
		   		"    text-align: left;\r\n" + 
		   		"    padding: 8px;\r\n" + 
		   		"}\r\n" + 
		   		"\r\n" + 
		   		"tr:nth-child(even) {background-color: #F0F8FF;}\r\n" + 
		   		"th {background-color: #FFDAB9;}\r\n" + 
		   		"</style>\r\n" + 
		   		"</head>\r\n" + 
		   		"<body>\r\n" + 
		   		"<center><h1>Comparator Report</h1></center>\r\n" + 
		   		"<center><h2>Comparator Graph</h2></center>\r\n" + 
		   		"<center><div id=\"piechart\"></div></center>\r\n" + 
		   		"<center><h2>Comparator Result</h2></center>"+
		   		"<table border ='2'>\r\n" + 
		   		"  <tr>\r\n" + 	
		   		"    <th>Match Result</th>\r\n" + 
		   		"    <th>Text</th>\r\n" + 
		   		"    <th>Expected</th>\r\n" + 
		   		"    <th>Actual</th>\r\n" + 
		   		"    <th>Near Translate</th>\r\n" +
		   		"  </tr>";
		   	
		   		
		   		
		   		
		   
		   
		String EnterText = "";
		PjtPath = ProjectPath + "\\MultiLingualTranslator";
		WriteLog(StartTag,LogName,true);
		int mat =0,nmat =0;
		
		if (Arr1.length == Arr2.length) {
			
			for (int k = 0; k < Arr1.length; k++) {
				
				if ((GetTransText( Language, Arr1[k]).trim()).equals(Arr2[k].trim())) {
					//EnterText = "Label Matched -" + Arr1[k] + "      Expected: "
					//		+ GetTransText(ProjectPath, Language, Arr1[k]) + "      Actual :" + Arr2[k];
					EnterText = "<td>Label Matched</td><td>" + Arr1[k] + "</td><td>"+ GetTransText( Language, Arr1[k]) + "</td><td>" + Arr2[k] + "</td><td></td>";
					//System.out.println(EnterText);
					mat++;
					WriteLog(EnterText,LogName,false);

				} else {
					EnterText = "<td><font color='red'>Label Not Matched</font></td><td>" + Arr1[k] + "</td><td>"+ GetTransText( Language, Arr1[k]) 
					+ "</td><td>" + Arr2[k] 
							+ "</td><td>"+ GetTransText( "English", Arr2[k]) +"</td>";
					//EnterText = "Label Not Matching -" + Arr1[k] + "      Expected: "
					//		+ GetTransText(ProjectPath, Language, Arr1[k]) + "      Actual :" + Arr2[k]
					//		+ "     Near Translate : " + GetTransText(ProjectPath, "English", Arr2[k]);
					//System.out.println(EnterText);
					nmat++;
					WriteLog(EnterText,LogName,false);
				}
			}

			
			} else {
			EnterText = "Length MisMatch -" + Arr1.length + " " + Arr2.length;
			System.out.println(EnterText);
			System.exit(0);
			//WriteLog(EnterText,LogName,false);
		}
		String EndTag = "</tr> \r\n" + 
				"</table>\r\n" +
				"<script type=\"text/javascript\" src=\"https://www.gstatic.com/charts/loader.js\"></script>\r\n" + 
				"\r\n" + 
				"<script type=\"text/javascript\">\r\n" + 
				"// Load google charts\r\n" + 
				"google.charts.load('current', {'packages':['corechart']});\r\n" + 
				"google.charts.setOnLoadCallback(drawChart);\r\n" + 
				"\r\n" + 
				"// Draw the chart and set the chart values\r\n" + 
				"function drawChart() {\r\n" + 
				"  var data = google.visualization.arrayToDataTable([\r\n" + 
				"  ['Result', 'Hours per Day'],\r\n" + 
				"  ['Matched',"+ mat+"],\r\n" + 
				"  ['Not Matched'," +nmat+"],\r\n" + 
				"\r\n" + 
				"]);\r\n" + 
				"\r\n" + 
				"  // Optional; add a title and set the width and height of the chart\r\n" + 
				"  var options = {'title':'', 'width':550, 'height':400};\r\n" + 
				"\r\n" + 
				"  // Display the chart inside the <div> element with id=\"piechart\"\r\n" + 
				"  var chart = new google.visualization.PieChart(document.getElementById('piechart'));\r\n" + 
				"  chart.draw(data, options);\r\n" + 
				"}\r\n" + 
				"</script>" +
				"</body> \r\n" + 
				"</html>\r\n" ; 
		WriteLog(EndTag,LogName,true);
		}
		
	

	private static void WriteLog(String Text,String LogName,boolean EndStartTag) throws IOException {
		File dir = new File(PjtPath + "\\Reports");

		if (!dir.exists()) {
			dir.mkdir();
		}

		dir = new File(PjtPath + "\\Reports\\"+LogName+".html");

		if (!dir.exists()) {
			dir.createNewFile();
		}

		FileWriter fw = new FileWriter(dir, true);

		BufferedWriter bw = new BufferedWriter(fw);
		PrintWriter out = new PrintWriter(bw);
		//out.write("\n");
		if(EndStartTag)
		{
		
		out.println(Text);	
		}
		else
		{
			out.println("<tr>");
			out.println(Text);
			out.println("</tr>");
		}
		
		out.close();
	}
	
	
  
  private static void AA() throws SQLException
  {
	  
	  Connection conn=DriverManager.getConnection(
		        "jdbc:ucanaccess://C:\\Users\\snaray01\\Documents\\Database1.accdb");
		Statement s = conn.createStatement();
		ResultSet rs = s.executeQuery("SELECT * FROM [TestTable]");
		while (rs.next()) {
		    System.out.println(rs.getString(1));
		}
  }
  

//  private static String ReadData(int RowNo, int ColNo) throws IOException
//	 {
//		//WebDriver driver;
//		//WebDriverWait wait;
//		XSSFWorkbook workbook;
//		XSSFSheet sheet;
//		XSSFCell cell;
//		 // Import excel sheet.
//		 File src=new File(PjtPath + "\\MultiLingualTranslator\\NonEngScript.xlsx");
//		 
//		 // Load the file.
//		 FileInputStream finput = new FileInputStream(src);
//		 
//		 // Load he workbook.
//		workbook = new XSSFWorkbook(finput);
//		 
//	     // Load the sheet in which data is stored.
//		 sheet= workbook.getSheet("NonEnglishText");
//		 
//	
//			
//			 cell = sheet.getRow(RowNo).getCell(ColNo);
//			 cell.setCellType(Cell.CELL_TYPE_STRING);
//			
//			 
//			   		
//	  //      }
//		 return cell.getStringCellValue();
//	  }
	    
  
//	private static void WriteData(int RowNo, int ColNo, String Data) throws IOException
//	 {
//		//create an object of Workbook and pass the FileInputStream object into it to create a pipeline between the sheet and eclipse.
//				FileInputStream fis = new FileInputStream(PjtPath + "\\NonEngScript.xlsx");
//				XSSFWorkbook workbook = new XSSFWorkbook(fis);
//				
//				XSSFSheet sheet = workbook.getSheet("NonEnglishText");
//				
//		               Row row = sheet.getRow(RowNo);
//		               
//		               Cell cell = row.getCell(ColNo);
//		       	
//		       		cell.setCellType(cell.CELL_TYPE_STRING);
//				
//				cell.setCellValue(Data);
//				FileOutputStream fos = new FileOutputStream(PjtPath + "\\NonEngScript.xlsx");
//				workbook.write(fos);
//				fos.close();
//				//System.out.println("END OF WRITING DATA IN EXCEL");
//	 }
	
//	private static void WriteinExcelTest(String ProjectPath, String Language,String WriteWord) throws IOException, SQLException
//	{
//		
//		
//		System.out.println(WriteWord);
//		String WriteWord1 =new String(WriteWord);
//		WriteWord1 = WriteWord1.toString();
//		String TranslatedWord = GetTransText(ProjectPath, Language, WriteWord1);
//		
//		 PjtPath = ProjectPath + "\\MultiLingualTranslator";
//		 WriteData(1, 0, WriteWord);
//		WriteData(1, 1, TranslatedWord);
//		System.out.println(TranslatedWord);
//	}
	 
	private static String fromCharCode(int...codePoints) {
	    StringBuilder builder = new StringBuilder(codePoints.length);
	    for (int codePoint :codePoints) {
	    	//System.out.println(builder);
	        builder.append(Character.toChars(codePoint));
	    }
	    return builder.toString();
	}
	
	private static String fromByteArr(String ConvertedWord)
	{
	   // String str = "Ìì¿Õ×Ô¶¯»¯";
	    int i,k;
	    int n = ConvertedWord.length();
	  
	    String bytes[] = new String[n * 2];
	    String bytes1[] = new String[n];
	    String bytes2[] = new String[n];

	    String ConvertedWordByteforArr = "";
    for(i = 0;  i < (n); i++) {
        int char1 = Character.codePointAt(ConvertedWord, i);

        bytes1[i] = "" + (char1 >>> 8);

        bytes2[i] = "" + (char1 & 0xFF);

	}
    int f = 0,s = 0;
    for(k = 0 ; k < (n*2); k++)
    {
    
    	if((k == 0 || k % 2 == 0))
    	{
    		bytes[k] = bytes1[f];
    		f++;
    	}
    	else 
    	{
    		bytes[k] = bytes2[s];
    		s++;
    	}
    	//89,41,122,122,129,234,82,168,83,22
    	//System.out.println(bytes[k]);
    	if(k == 0)
    	{
    		ConvertedWordByteforArr = bytes[k];
    	}
    	else
    	{
    	ConvertedWordByteforArr = ConvertedWordByteforArr + "," +bytes[k];
    	}
    	
    	
    }
    return ConvertedWordByteforArr;
	}
	
	
	
	public static String GetTransTextNonEnglishAlphabets_test(String ProjectPath, String Language, String ConvertWordByteforArr)
			throws SQLException, IOException {
		
		String ConvertWord = "";
		int k=0;
		if(ConvertWordByteforArr.contains(","))
		{
			String[] ConvertWordByteArr = ConvertWordByteforArr.split(",");
			int[] ConvertWordintByteArr = new int[ConvertWordByteArr.length+1] ;
			for(int z = 0; z < ConvertWordByteArr.length; z++)
			{				
				ConvertWordintByteArr[z] = Integer.parseInt(ConvertWordByteArr[z]);
				k = ((ConvertWordintByteArr[z++] & 0xff) << 8) | (ConvertWordintByteArr[z++] & 0xff);
				ConvertWord = ConvertWord + fromCharCode(k);
			}
			
		}
		else
		{
			int[] ConvertWordintByteArr = new int[0] ;
			ConvertWordintByteArr[0] = Integer.parseInt(ConvertWordByteforArr);
			ConvertWord = fromCharCode(ConvertWordintByteArr[0]);
		}
		
		System.out.println("AAA  -----" +ConvertWord);
		String ConvertedText = "";
		//String ConvertWord = "";		
		//ConvertWord = new String(ConvertWordArr);
		PjtPath = ProjectPath;
		// Lang = Language;
		CvtWord = ConvertWord;
		File dir = new File(PjtPath + "\\MultiLingualTranslator");

		if (!dir.exists()) {
			GetMultiLingualTranslatorFolder(dir);
		}
		PjtPath = PjtPath + "\\MultiLingualTranslator";

		dir = new File(PjtPath + "\\node_modules\\google-translate-api");

		if (!dir.exists()) {
			GetNodeModules();
		}

		dir = new File(PjtPath + "\\LanguageDB.accdb");

		if (!dir.exists()) {
			// GetLanguageDB(dir);
			System.out.println("LanguageDB not exists.. Kindly place the required DB in the ProjectPath");
			System.exit(0);
		}
		
		
		Lang = ConnectDB(" Select Code from [LanguageCode] where Language = '" + Language + "'");
		// SELECT Code FROM LanguageCode where Language='German'

		// System.out.println(Lang);
		if (Language != "English") {
			String PresentText = ConnectDB("Select LangText from [" + Language + "] where EngText = '" + CvtWord + "'");

			if (PresentText == "") {
			//	MakeTranslateJs(CvtWord, Lang);
				ConvertedText = RunTranslateJS(Lang,CvtWord);
				ConnectDB("Insert into [" + Language + "](Flag,EngText,LangText,LangText1) values (0,'" + CvtWord + "','"
						+ ConvertedText + "','')");
			} else {

				// PresentText = ConnectDB("Select LangText from "+Language+" where EngText =
				// '"+CvtWord + "'" );
				
				if(ConnectDB("Select Flag from [" + Language + "] where EngText = '" + CvtWord + "'").equals("0"))
				{
				ConvertedText = PresentText;
				}
				else
				{
					ConvertedText =	ConnectDB("Select LangText1 from [" + Language + "] where EngText = '" + CvtWord + "'");
				}
			}
		} else {
		//	MakeTranslateJs(CvtWord, Lang);
			ConvertedText = RunTranslateJS(Lang,CvtWord);
		}
		System.out.println("BBB  -----" +ConvertedText);
	//	ConvertedTextArr = ConvertedText.getBytes();
		
	
		String ConvertedTextByteforArr = fromByteArr(ConvertedText);
		System.out.println("CCC  -----" +ConvertedTextByteforArr);
		
		return ConvertedTextByteforArr;
	}
 }



