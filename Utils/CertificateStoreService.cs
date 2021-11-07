using Logging;
using System;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;

namespace Common
{
    public class CertificateStoreService : ICertificateStoreService
    {
        private X509Store store;
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
                    this.CertErrorMessage(certName, signingCerts.Count);
                    return null;
                }
                else
                {
                    string sPattern = @"CN=";

                    foreach (var cert in signingCerts)
                    {
                        var issuerName = Regex.Replace(cert.Issuer, sPattern, string.Empty);

                        if (issuerName.Equals(certName))
                        {
                            this.logger.Debug($"Found certificate: {cert} Certificates list has: {signingCerts.Count} inserts");
                            return cert;
                        }
                    }

                    this.CertErrorMessage(certName, signingCerts.Count);
                    return null;
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

        private void CertErrorMessage(string certName, int count)
        {
            this.logger.Error($"Certificate: {certName} not found! Certificates list has: {count} inserts");
            Console.WriteLine($"Certificate: {certName} not found!");
        }
    }
}
