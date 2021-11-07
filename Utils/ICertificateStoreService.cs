using System.Security.Cryptography.X509Certificates;

namespace Common
{
    public interface ICertificateStoreService
    {
        X509Certificate2 GetCertificateFromStore(string certName);
        X509Certificate2 GetCertificateWithoutPrivateKeyFromStore();
    }
}
