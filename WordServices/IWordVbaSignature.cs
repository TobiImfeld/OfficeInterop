using System.Security.Cryptography.X509Certificates;

namespace WordServices
{
    public enum eSyskind
    {
        Win16 = 0,
        Win32 = 1,
        Macintosh = 2,
        Win64 = 3
    }

    public interface IWordVbaSignatureService
    {
        void AddDigitalSignature(string file, X509Certificate2 cert);
    }
}

