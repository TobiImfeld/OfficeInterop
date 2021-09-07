using Logging;
using System;
using System.Security.Cryptography.X509Certificates;

namespace ExcelServices
{
    public class CertificateStoreService : ICertificateStoreService
    {
        private X509Store store;
        private X509Certificate2 signingCert;
        private readonly ILogger logger;

        public CertificateStoreService(ILoggerFactory loggerFactory)
        {
            this.logger = loggerFactory.Create<CertificateStoreService>();
        }

        public X509Certificate2 GetCertificateFromStore(string certName)
        {
            this.store = new X509Store(StoreLocation.CurrentUser);
            this.logger.Info($"Open certificate store for: {StoreLocation.CurrentUser.ToString()}");

            try
            {
                store.Open(OpenFlags.ReadOnly);

                X509Certificate2Collection certCollection = store.Certificates;
                X509Certificate2Collection currentCerts = certCollection.Find(X509FindType.FindByIssuerName, certName, false);
                X509Certificate2Collection signingCerts = currentCerts.Find(X509FindType.FindByIssuerName, certName, false);

                if (signingCerts.Count == 0)
                {
                    this.logger.Error($"Certificate: {certName} not found! Certificates list has: {signingCerts.Count} inserts");
                    Console.WriteLine($"Certificate: {certName} not found!");
                    return null;
                }
                else
                {
                    signingCert = signingCerts[0];
                    this.logger.Debug($"Found certificate: {signingCert} Certificates list has: {signingCerts.Count} inserts");
                    return signingCert;
                }
            }
            catch (Exception ex)
            {
                this.logger.Error(ex);
                return null;
            }
            finally
            {
                this.logger.Debug($"Close certificate store");
                store.Close();
            }
        }

        public X509Certificate2 GetCertificateWithoutPrivateKeyFromStore()
        {
            this.store = new X509Store(StoreName.Root, StoreLocation.CurrentUser);
            this.logger.Info($"Open certificate store: {StoreName.Root.ToString()} for: {StoreLocation.CurrentUser.ToString()}");

            try
            {
                store.Open(OpenFlags.ReadOnly);
                X509Certificate2Collection certCollection = store.Certificates;

                foreach (var cert in certCollection)
                {
                    if (!cert.HasPrivateKey)
                    {
                        this.logger.Info($"Certificate {cert.Issuer} without private key found");
                        return cert;
                    }
                }

                this.logger.Info($"None certificate found without private key!");
                return null;
            }
            catch (Exception ex)
            {
                this.logger.Error(ex);
                return null;
            }
            finally
            {
                this.logger.Debug($"Close certificate store");
                store.Close();
            }
        }
    }
}
