//USEUNIT ExcelUtils
//USEUNIT ValidationUtils
//USEUNIT WorkspaceUtils
//USEUNIT EnvParams
//USEUNIT ReportUtils
//USEUNIT PdfUtils

function ExecuteScript(Texts,convTOLang){ 
  var ConvertedText = GetTransText(Texts,convTOLang);
  Log.Message(ConvertedText);
}

function GetTransText(TransText,convTOLang){ 
  var path = ProjectSuite.Path;
  var Lang = convTOLang;
  var TransWord = JavaClasses.MLT.MultiLingualTranslator.GetTransText(path,Lang, TransText);
  return TransWord;

}

function TransText(){ 
  ExecuteScript("Dieu","English");
//fr
//French
}