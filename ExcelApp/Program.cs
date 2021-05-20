using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace ExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {
            object sigID = "{7973591c-a24c-5814-1954-1dd7667f7ddc}";

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
