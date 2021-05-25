using System;
using System.Security.Cryptography.X509Certificates;

namespace ExcelServices
{
    public class CertificateStoreService : ICertificateStoreService
    {
        private X509Store store;
        private X509Certificate2 signingCert;

        public X509Certificate2 GetCertificateFromStore(string certName)
        {
            store = new X509Store(StoreLocation.CurrentUser);

            try
            {
                store.Open(OpenFlags.ReadOnly);

                // Place all certificates in an X509Certificate2Collection object.
                X509Certificate2Collection certCollection = store.Certificates;
                // If using a certificate with a trusted root you do not need to FindByTimeValid, instead:
                // currentCerts.Find(X509FindType.FindBySubjectDistinguishedName, certName, true);
                //X509Certificate2Collection currentCerts = certCollection.Find(X509FindType.FindByTimeValid, DateTime.Now, false);
                X509Certificate2Collection currentCerts = certCollection.Find(X509FindType.FindByIssuerName, "TobiOfficeCert", false);
                X509Certificate2Collection signingCerts = currentCerts.Find(X509FindType.FindByIssuerName, "TobiOfficeCert", false);

                if (signingCerts.Count == 0)
                {
                    return null;
                }
                else
                {
                    signingCert = signingCerts[0];
                    return signingCert;
                }    
            }
            catch(Exception ex)
            {
                Console.WriteLine($"GetCertificateFromStore(): {certName} not found! Exception: {ex}");
                return null;
            }
            finally
            {
                store.Close();
            }
        }
    }
}
