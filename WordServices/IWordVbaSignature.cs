using System.Security.Cryptography.X509Certificates;

namespace WordServices
{
    public interface IWordVbaSignatureService
    {
        X509Certificate2 GetSignatureFromZipPackage(string targetDirectory);
    }
}
