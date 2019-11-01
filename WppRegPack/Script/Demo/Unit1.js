function ExecuteScript(){ 
  var ConvertedText = GetTransText("功能");
//  var EngStr = "New,Functions,Clients";
//  var LangArr = "Neu,Funktions,Kunden";
//  var LangArr = "新,功能,客户端";
//  ComparatorNonArray(EngStr, LangArr);
 Log.Message(ConvertedText);
// Log.Message(aqEnvironment.LanguageForNonUnicodePrograms);
}

function GetTransText(TransText){ 
  var path = ProjectSuite.Path;
//  var Lang = ProjectSuite.Variables.LanguageController;
//  var Lang = "Tamil";
  var Lang = "English";
//  var Lang = "Chinese (Simplified)";
  //var TranWord = JavaClasses.MLT.MultiLingualTranslator.GetTransText(path, Lang, TransText);
  var TransWord = JavaClasses.MLT.MultiLingualTranslator.GetTransText(path,Lang, TransText);
  return TransWord;

}

function ComparatorNonArray(Arr1, Arr2) { 
  var path = ProjectSuite.Path;
//  var Lang = ProjectSuite.Variables.LanguageController;
//  var Lang = "Tamil";
  var Lang = "English";
//  var Lang = "Chinese (simplified)";
  JavaClasses.MLT.MultiLingualTranslator.ComparatorNonArray(path, Lang,Arr1,Arr2);
  Log.Message("Comparison Over... Check Reports");
}