//USEUNIT ReportUtils
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
var value = "";
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

function readDetails(UserConfigPath,element)
{
//var UserConfigPath = Project.Path + "\\" + "ProjectConfig.config";
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


function writeDetails(notepadPath,element,Data){
      if(!aqFile.Exists(notepadPath)){
      aqFile.Create(notepadPath);
      }
   
//Log.Message(notepadPath);
oFile = aqFile.OpenTextFile(notepadPath, aqFile.faRead, aqFile.ctANSI);
oFile.Cursor = 0;
var NotePadlines = [];
var i=0;
var status = true;
while(! oFile.IsEndOfFile()){
line = oFile.ReadLine(); 

key = line.substring(0, line.indexOf("=")); 


 if(key.toLowerCase().trim()==element.toLowerCase().trim()) //aqString.Contains(InputString, SubString, StartPosition, CaseSensitive)
{
  NotePadlines[i] = element+"= "+Data;
  i++;
  status =false;
}
else{ 
  NotePadlines[i] = line;
  i++;
}

}
if(status){ 
   NotePadlines[i] = NotePadlines[i] = element+"= "+Data;;
  i++; 
}

oFile.Close();

  oFile = aqFile.OpenTextFile(notepadPath, aqFile.faWrite, aqFile.ctANSI, true);
  for(var i=0;i<NotePadlines.length;i++)
  oFile.WriteLine(NotePadlines[i]);

oFile.Close();
}

function writeLog(Data){
      if(!aqFile.Exists(ReportUtils.file_path+"\\TestLog.txt")){
      aqFile.Create(ReportUtils.file_path+"\\TestLog.txt");
      }
   
//Log.Message(notepadPath);
oFile = aqFile.OpenTextFile(ReportUtils.file_path+"\\TestLog.txt", aqFile.faRead, aqFile.ctANSI);
oFile.Cursor = 0;
var NotePadlines = [];
var i=0;
var status = true;
while(! oFile.IsEndOfFile()){
line = oFile.ReadLine(); 

  NotePadlines[i] = line;
  i++;

}
NotePadlines[i] = Data;
oFile.Close();

  oFile = aqFile.OpenTextFile(ReportUtils.file_path+"\\TestLog.txt", aqFile.faWrite, aqFile.ctANSI, true);
  for(var i=0;i<NotePadlines.length;i++)
  oFile.WriteLine(NotePadlines[i]);

oFile.Close();
}