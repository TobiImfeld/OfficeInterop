using System.Security.Cryptography.X509Certificates;

namespace ExcelServices
{
    public interface ICertificateStoreService
    {
        X509Certificate2 GetCertificateFromStore(string certName);
    }
}
