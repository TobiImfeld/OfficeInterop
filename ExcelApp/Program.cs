using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Security.Cryptography.X509Certificates;

namespace ExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {
            object sigID = "{7973591c-a24c-5814-1954-1dd7667f7ddc}";

            var excelApp = new Application();
            excelApp.Visible = true;

            // Get the certifcate to use to encrypt the key.
            X509Certificate2 cert = GetCertificateFromStore("CN=TobiOfficeCert");
            if (cert == null)
            {
                Console.WriteLine("Certificate ' CN=TobiOfficeCert' not found.");
                //Console.ReadLine();
            }
            else
            {
                Console.WriteLine("Certificate ' CN=TobiOfficeCert' found.");
                //Console.ReadLine();
            }

            Workbook excelFile = excelApp.Workbooks.Open("C:\\Temp\\Test1.xlsx");
            
            var signatureSet = excelFile.Signatures;
            Signature objSignature = signatureSet.AddNonVisibleSignature(cert);
            var signed = objSignature.IsSigned;
            Console.WriteLine($"Is file signed: {signed}");

            excelFile.SaveAs("C:\\Temp\\Test1.xlsx");
            excelFile.Close();


            //oWord.Activate();

            //SignatureSet signatureSet = oWord.ActiveDocument.Signatures;
            //// signatureSet.ShowSignaturesPane = false;
            //Signature objSignature = signatureSet.AddSignatureLine(sigID);
            //objSignature.Setup.SuggestedSigner = "docSigner";
            //objSignature.Setup.SuggestedSignerEmail = "abc@xyz.com";
            //objSignature.Setup.ShowSignDate = true;
            ////  dynamic shape = objSignature.SignatureLineShape;


            // See how it works for word: https://www.codeproject.com/Questions/1233135/Digital-signature-on-documents
            // See pfx cert for excel: https://stackoverflow.com/questions/55159759/how-to-signed-excel-with-my-private-pfx-key
        }

        private static X509Certificate2 GetCertificateFromStore(string certName)
        {
            //See link: https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.x509certificates.x509certificate2?view=net-5.0#examples
            // Get the certificate store for the current user.
            X509Store store = new X509Store(StoreLocation.CurrentUser);
            try
            {
                store.Open(OpenFlags.ReadOnly);

                // Place all certificates in an X509Certificate2Collection object.
                X509Certificate2Collection certCollection = store.Certificates;
                // If using a certificate with a trusted root you do not need to FindByTimeValid, instead:
                // currentCerts.Find(X509FindType.FindBySubjectDistinguishedName, certName, true);
                //X509Certificate2Collection currentCerts = certCollection.Find(X509FindType.FindByTimeValid, DateTime.Now, false);
                X509Certificate2Collection currentCerts = certCollection.Find(X509FindType.FindByIssuerName, "TobiOfficeCert", false);
                X509Certificate2Collection signingCert = currentCerts.Find(X509FindType.FindByIssuerName, "TobiOfficeCert", false);
                if (signingCert.Count == 0)
                    return null;
                // Return the first certificate in the collection, has the right name and is current.
                return signingCert[0];
            }
            finally
            {
                store.Close();
            }
        }
    }
}
