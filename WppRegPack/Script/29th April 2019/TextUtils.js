
var str;

function ReadWholeFile(AFileName)
{
var s = aqFile.ReadWholeTextFile(AFileName, aqFile.ctANSI);
Log.Message("File entire contents:");
Log.Message(s);
}

function GetUserValue(element)
{
var UserConfigPath = Project.Path + "\\" + "UserConfig.config";
var File, line;
   
File = aqFile.OpenTextFile(UserConfigPath, aqFile.faRead, aqFile.ctANSI);
File.Cursor = 0;


while(! File.IsEndOfFile()){
line = File.ReadLine(); 

key = line.substring(0, line.indexOf("=")); 
   

 if(key.toLowerCase().trim()==element.toLowerCase().trim()) //aqString.Contains(InputString, SubString, StartPosition, CaseSensitive)
{
value = line.substring(line.indexOf("=") + 1);
value = value.trim();

}
}

File.Close();
return value;
}

function GetProjectValue(element)
{
var UserConfigPath = Project.Path + "\\" + "ProjectConfig.config";
var File, line;
   
File = aqFile.OpenTextFile(UserConfigPath, aqFile.faRead, aqFile.ctANSI);
File.Cursor = 0;


while(! File.IsEndOfFile()){
line = File.ReadLine(); 

key = line.substring(0, line.indexOf("=")); 
   

 if(key.toLowerCase().trim()==element.toLowerCase().trim()) //aqString.Contains(InputString, SubString, StartPosition, CaseSensitive)
{
value = line.substring(line.indexOf("=") + 1);
value = value.trim();

}
}

File.Close();
return value;
}

function Run()
{

//Log.Message(AFileName);
//ReadWholeFile(AFileName);
val = GetValue("Browser");
Log.Message(val)
}