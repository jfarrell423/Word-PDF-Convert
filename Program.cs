using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Word;

// ******************************************************************************* 
// **** Program to convert a word document to a pdf document 
// **** 1) This needs to Be A Microsoft Core Solution
// **** Jerry Farrell 
// *******************************************************************************
// **** 2) Added the Following COM Reference
// **** Microsoft Office 16.0 Object Library
// *******************************************************************************
// **** 3) The Following NuGet Package Needs to be Added to The Solution
// **** Microsoft.Office.Interop.Word
// *******************************************************************************

var wordApp = new Microsoft.Office.Interop.Word.Application();
var directoryPath = Directory.GetCurrentDirectory();
foreach (string filePath in Directory.GetFiles(directoryPath, "*.docx", SearchOption.AllDirectories)
    .Concat(Directory.GetFiles(directoryPath, "*.doc", SearchOption.AllDirectories)))
{
    var doc = wordApp.Documents.Open(filePath);
    var pdfFilePath = Path.ChangeExtension(filePath, ".pdf");
    doc.SaveAs2(pdfFilePath, WdSaveFormat.wdFormatPDF);
    doc.Close();

    // Rename the PDF file to match the original file name
    File.Move(pdfFilePath, Path.Combine(Path.GetDirectoryName(pdfFilePath), Path.GetFileNameWithoutExtension(filePath) + ".pdf"));
}
wordApp.Quit();
