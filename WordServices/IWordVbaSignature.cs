using System.Security.Cryptography.X509Certificates;

namespace WordServices
{
    public interface IWordVbaSignatureService
    {
        void AddDigitalSignature(string file, X509Certificate2 cert);
    }
}

