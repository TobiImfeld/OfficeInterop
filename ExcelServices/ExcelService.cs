using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;

namespace ExcelServices
{
    public class ExcelService : IExcelService
    {
        private ICertificateStoreService certificateStoreService;
        private Application excelApp = null;
        private Workbooks books = null;
        private Workbook book = null;

        public ExcelService(ICertificateStoreService certificateStoreService)
        {
            this.certificateStoreService = certificateStoreService;
        }

        public void AddDigitalSignature(string certName)
        {
            var cert = this.certificateStoreService.GetCertificateFromStore(certName);

            if (cert == null)
            {
                Console.WriteLine($"AddDigitalSignature(): Certificate CN={certName}' not found.");
            }
            else
            {
                try
                {
                    excelApp = new Application();
                    books = excelApp.Workbooks;
                    book = books.Open("C:\\Temp\\Test1.xlsx");

                    excelApp.DisplayAlerts = false;
                    excelApp.Visible = true;

                    var signatureSet = book.Signatures;
                    Signature objSignature = signatureSet.AddNonVisibleSignature(cert);
                    //var signed = objSignature.IsSigned;
                    //Console.WriteLine($"Is file signed: {signed}");

                    //excelFile.SaveAs("C:\\Temp\\Test1.xlsx");

                    //Vermutung: Wahrscheinlich muss die Signierung zuerst komplett abgearbeitet sein, Signierungsfenster geschlossen, keine Zugriffe mehr auf File.
                    //ûnd erst dann kann das Excel gespeichert werden.
                    //excelFile.Save();
                    //excelFile.SaveCopyAs("C:\\Temp\\Test1.xlsx");
                    //book.SaveAs(@"C:\Temp\Test1_" + DateTime.Now.Millisecond + ".xlsx"); //Zeile hier verhindert Save-Error, aber signaturen werden ungültig durch bearbeitung!
                    //book.SaveCopyAs("C:\\Temp\\Test1.xlsx");
                    book.Save();
                    book.Close();
                    excelApp.Quit();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    this.disposeComObjects();
                    Console.ReadLine();
                }
                finally
                {
                    this.disposeComObjects();
                }
            }
        }

        private void disposeComObjects()
        {
            if (book != null) Marshal.ReleaseComObject(book);
            if (books != null) Marshal.ReleaseComObject(books);
            if (excelApp != null) Marshal.ReleaseComObject(excelApp);
        }
    }
}
