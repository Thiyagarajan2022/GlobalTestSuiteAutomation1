function mplValidate(){ 
 var docObj = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3("C:\\WppRegression_v12.50\\WppRegPack\\MPLReports\\Job Detail_1707200113.pdf");
// docObj = loadDocument("C:\\WppRegression_v12.50\\WppRegPack\\MPLReports\\Job Detail_1707200084.pdf");
 var textobj;
  try{
  var obj = JavaClasses.org_apache_pdfbox_util.PDFTextStripper.newInstance();
  textobj = obj.getText_2(docObj);
  Log.Message(textobj)
  var pictureObj = convertPageToPicture(docObj,0,"C:\\Users\\674087\\Pictures\\pdfimage1.png");
//  Regions.Compare("C:\\Users\\674087\\Pictures\\pdfimage.png", "C:\\Users\\674087\\Pictures\\pdfimage1.png");
  }catch(objEx){
    Log.Error("Exception while getting text from document::"+objEx);
  }
  verify(textobj.contains("Job: 1707200113"),true,"Job Number is available");
}

function verify(actual, expected, message){
if(actual == expected){
Log.Checkpoint(message);
}
else{  
Log.Error(message);
}
}

function convertPageToPicture(docObj, pageIndex, fileName)
{
  var pageObj, imgBuffer, imgFile, imgFormat, pictureObj;
  // Get the desired page
  pageObj = getPage(docObj, pageIndex);

  // Convert the page to image data
  imgBuffer = pageObj.convertToImage();

  // Create a new file to save
  imgFile = JavaClasses.java_io.File.newInstance(fileName);

  // Get the image format from the name
  imgFormat = aqString.SubString(fileName, aqString.GetLength(fileName)-3, 3);

  // Save the image to the created file
  JavaClasses.javax_imageio.ImageIO.write(imgBuffer, imgFormat, imgFile);

  // Create a Picture object
  pictureObj = Utils.Picture;

  // Load the image as a picture
  pictureObj.LoadFromFile(fileName);

  // Return the picture object
  return pictureObj; 
}

function getPage(docObj, pageIndex)
{
  var pageArray, pageObj;

  // Obtain a collection of the pages 
  pageArray = docObj.getDocumentCatalog().getAllPages();

  // Obtain the specified page
  pageObj =  pageArray.get(pageIndex);

  // Return the result
  return pageObj;
} 


function compareDocsAsImg()
{
//  var pdfFile_1 = "C:\\WppRegression_v12.50\\WppRegPack\\MPLReports\\Job Detail_1707200084.pdf";
  var pdfFile_1 = "C:\\WppRegression_v12.50\\WppRegPack\\MPLReports\\Job Detail_1707200113.pdf";
  var pdfFile_2 = "C:\\WppRegression_v12.50\\WppRegPack\\MPLReports\\Job Detail_1707200113 - Copy.pdf";
  var maskImg = "";
  var imgFile_1, imgFile_2, docObj_1, docObj_2, totalPages_1, totalPages_2;

  // Specify the fully-qualified name of the temporary image files
  imgFile_1 = "C:\\Temp\\page_doc_1.png";
  imgFile_2 = "C:\\Temp\\page_doc_2.png";

  // Load specified documents
  docObj_1 = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(pdfFile_1);
  docObj_2 = JavaClasses.org_apache_pdfbox_pdmodel.PDDocument.load_3(pdfFile_2);

  // Get the total number of pages in both documents
  totalPages_1 = docObj_1.getNumberOfPages();
  totalPages_2 = docObj_2.getNumberOfPages();

  // Check whether the documents contain the same number of the pages
  if (totalPages_1 != totalPages_2)
  {
    Log.Message("The documents contain different number of pages.");
  } else
  {
    for (i = 0; i < totalPages_1; i++)
    {
      // Call a routine that converts the specified page to an image
      var obj = JavaClasses.org_apache_pdfbox_util.PDFTextStripper.newInstance();
      textobj = obj.getText_2(docObj_1);
      Log.Message(textobj)
      pic_1 = convertPageToPicture(docObj_1, i, imgFile_1);
      var obj = JavaClasses.org_apache_pdfbox_util.PDFTextStripper.newInstance();
      textobj = obj.getText_2(docObj_1);
      Log.Message(textobj)
      pic_2 = convertPageToPicture(docObj_2, i, imgFile_2);

      // Compare two images
      if (!pic_1.Compare(pic_2))
      {
        // If the images are different...
        // Post image differences to the log
        Log.Picture(pic_1.Difference(pic_2));

        // Post a warning message
        Log.Warning("Pages " + aqConvert.IntToStr(i+1) + " are different. Documents are different.");

        // Break the loop
//        break;
      } else
      {
        // Post a message that the pages are equal
        Log.Message("Pages " + aqConvert.IntToStr(i+1) + " are equal.")
      }
      
      // Delete the temporary image files
      aqFile.Delete(imgFile_1);
      aqFile.Delete(imgFile_2);
    }
  }
}



function Test()
{

  
  var w1 = Sys.Process("*").SWTObject("Shell", "*").SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "").SWTObject("Composite", "", 3).SWTObject("Composite", "", 1).SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "").SWTObject("Composite", "", 7).SWTObject("Composite", "").SWTObject("PTabFolder", "").SWTObject("Composite", "", 3).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("McClumpSashForm", "").SWTObject("Composite", "", 1).SWTObject("Composite", "");
  Regions.AddPicture(w1, "MyAppWnd");

  var w2 = Sys.Desktop.ActiveWindowtestList();

  if (! Regions.Compare("MyAppWnd", w2))
    Log.Error("The compared regions are not identical.", w2.Name);

}