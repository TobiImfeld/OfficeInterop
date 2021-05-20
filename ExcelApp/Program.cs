using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace ExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {
            object sigID = "{4157e14126de03994c7e41c6d36d8ea7}";

            var excelApp = new Application();
            excelApp.Visible = true;

            Workbook excelFile = excelApp.Workbooks.Open("C:\\Temp\\Test1.xlsx");
            var signatureSet = excelFile.Signatures;
            Signature objSignature = signatureSet.AddNonVisibleSignature(sigID);



            //Signature objSignature = signatureSet.AddSignatureLine(sigID);
            //objSignature.Setup.SuggestedSigner = "docSigner";
            //objSignature.Setup.SuggestedSignerEmail = "abc@xyz.com";
            //objSignature.Setup.ShowSignDate = true;


            // See how it works for word: https://www.codeproject.com/Questions/1233135/Digital-signature-on-documents
        }
    }
}
