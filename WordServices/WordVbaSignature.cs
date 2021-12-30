using System.Linq;
using System.IO.Packaging;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Pkcs;
using Logging;
using Common;
using System.Collections.Generic;
using System.IO.Compression;

namespace WordServices
{
    public class WordVbaSignatureService : IWordVbaSignatureService
    {
        private readonly ILogger logger;
        private readonly IFileService fileService;
        private const string schemaRelVbaSignature = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";
        private ZipPackagePart vbaProjectPart = null;
        private X509Certificate2 certificate = null;
        private SignedCms verifier = null;
        private List<string> fileList;

        public WordVbaSignatureService(ILoggerFactory loggerFactory, IFileService fileService)
        {
            this.logger = loggerFactory.Create<WordVbaSignatureService>();
            this.fileService = fileService;
            this.GetSignature();
        }

        public void SetPathToFiles(string targetDirectory)
        {
            this.fileList = this.ListAllWordFilesFromDirectory(targetDirectory);
        }

        public void GetSignatureFromZipPackage(string targetDirectory)
        {
            this.logger.Debug(targetDirectory);
            this.GetVbaPackagePart(targetDirectory);
            this.GetSignature();
        }

        private void GetVbaPackagePart(string targetDirectory)
        {
            using (ZipPackage appx = Package.Open(targetDirectory, FileMode.Open, FileAccess.Read) as ZipPackage)
            {
                var name = "/word/vbaProject.bin";
                PackagePartCollection packagePartCollection = appx.GetParts();
                var vbaProjectPart = packagePartCollection.FirstOrDefault(u => u.Uri.Equals(name));
                this.vbaProjectPart = (ZipPackagePart)vbaProjectPart; //Problem: Wenn der using Bereich verlassen wird, wird das Objekt abgeräumt. Problem auflösen!
            }
        }

        private void GetSignature()
        {
            if (this.vbaProjectPart == null)
            {
                return;
            }

            var rel = this.vbaProjectPart.GetRelationshipsByType(schemaRelVbaSignature).FirstOrDefault();

            if (rel != null)
            {
                var uri = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                var part = this.vbaProjectPart.Package.GetPart(uri);

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
                                this.certificate = new X509Certificate2(value);
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

                this.verifier = new SignedCms();
                this.verifier.Decode(signature);
            }
            else
            {
                this.certificate = null;
                this.verifier = null;
            }
        }

        private List<string> ListAllWordFilesFromDirectory(string targetDirectory)
        {
            var filesFromDirectory = this.fileService.
                ListAllFilesFromDirectoryByFileExtension(
                targetDirectory,
                OfficeFileExtensions.DOCM
                );
            return filesFromDirectory;
        }

        private void ClearFileList()
        {
            this.fileList.Clear();
        }
    }
}
