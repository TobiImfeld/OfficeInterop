using System.Linq;
using System.IO.Packaging;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Pkcs;
using Logging;
using Common;

namespace WordServices
{
    public class WordVbaSignatureService : IWordVbaSignatureService
    {
        private readonly ILogger logger;
        private readonly IFileService fileService;
        private const string schemaRelVbaSignature = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";

        public WordVbaSignatureService(ILoggerFactory loggerFactory, IFileService fileService)
        {
            this.logger = loggerFactory.Create<WordVbaSignatureService>();
            this.fileService = fileService;
        }

        public X509Certificate2 GetSignatureFromZipPackage(string targetDirectory)
        {
            X509Certificate2 certificate = null;
            this.logger.Debug(targetDirectory);
            certificate = this.GetSignature(targetDirectory);
            return certificate;
        }

        private X509Certificate2 GetSignature(string targetDirectory)
        {
            using (ZipPackage appx = Package.Open(targetDirectory, FileMode.Open, FileAccess.Read) as ZipPackage)
            {
                var name = "/word/vbaProject.bin";
                PackagePartCollection packagePartCollection = appx.GetParts();
                var vbaProjectPart = packagePartCollection.FirstOrDefault(u => u.Uri.Equals(name));

                return this.ReadSignature((ZipPackagePart)vbaProjectPart);
            }
        }

        private X509Certificate2 ReadSignature(ZipPackagePart vbaProjectPart)
        {
            X509Certificate2 certificate = null;
            SignedCms verifier = null;

            if (vbaProjectPart == null)
            {
                return null;
            }

            var rel = vbaProjectPart.GetRelationshipsByType(schemaRelVbaSignature).FirstOrDefault();

            if (rel != null)
            {
                var uri = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                var part = vbaProjectPart.Package.GetPart(uri);

                var stream = part.GetStream();
                BinaryReader br = new BinaryReader(stream);
                uint cbSignature = br.ReadUInt32();
                uint signatureOffset = br.ReadUInt32();
                uint cbSigningCertStore = br.ReadUInt32();
                uint certStoreOffset = br.ReadUInt32();
                uint cbProjectName = br.ReadUInt32();
                uint projectNameOffset = br.ReadUInt32();
                uint fTimestamp = br.ReadUInt32();
                uint cbTimestampUrl = br.ReadUInt32();
                uint timestampUrlOffset = br.ReadUInt32();
                byte[] signature = br.ReadBytes((int)cbSignature);
                uint version = br.ReadUInt32();
                uint fileType = br.ReadUInt32();

                uint id = br.ReadUInt32();
                while (id != 0)
                {
                    uint encodingType = br.ReadUInt32();
                    uint length = br.ReadUInt32();
                    if (length > 0)
                    {
                        byte[] value = br.ReadBytes((int)length);
                        switch (id)
                        {
                            case 0x20:
                                certificate = new X509Certificate2(value);
                                break;
                            default:
                                break;
                        }
                    }

                    id = br.ReadUInt32();
                }

                uint endel1 = br.ReadUInt32();
                uint endel2 = br.ReadUInt32();
                ushort rgchProjectNameBuffer = br.ReadUInt16();
                ushort rgchTimestampBuffer = br.ReadUInt16();

                verifier = new SignedCms();
                verifier.Decode(signature);

                return certificate;
            }
            else
            {
                certificate = null;
                verifier = null;

                return certificate;
            }
        }
    }
}
