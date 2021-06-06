﻿using Logging;
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
                    this.logger.Error($"{certName} not found! Certificates list has: {signingCerts.Count} inserts");
                    return null;
                }
                else
                {
                    signingCert = signingCerts[0];
                    this.logger.Debug($"Found certificate: {signingCert}");
                    return signingCert;
                }    
            }
            catch(Exception ex)
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
