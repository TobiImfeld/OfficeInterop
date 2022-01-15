using System.Linq;
using System.IO.Packaging;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Pkcs;
using Logging;
using WordServices.Dto;
using System.Text;

namespace WordServices
{
    public class WordVbaSignatureService : IWordVbaSignatureService
    {
        public X509Certificate2 Certificate { get; internal set; }
        public int CodePage { get; internal set; }

        private readonly ILogger logger;
        private const string schemaRelVbaSignature = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";
        private const string schemaRelVba = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
        private const string schemaRelVbaData = "http://schemas.microsoft.com/office/2006/relationships/wordVbaData";

        public WordVbaSignatureService(ILoggerFactory loggerFactory)
        {
            this.logger = loggerFactory.Create<WordVbaSignatureService>();
        }

        public void AddDigitalSignature(string file, X509Certificate2 cert)
        {
            this.logger.Debug($"sign {file} with {cert.Issuer}");
            this.AddDigitalSignatureToVbaProjectPart(file, cert);
        }

        private void AddDigitalSignatureToVbaProjectPart(string file, X509Certificate2 cert)
        {
            using (ZipPackage appx = Package.Open(file, FileMode.Open, FileAccess.Read) as ZipPackage)
            {
                var name = "/word/vbaProject.bin";
                var doc = "/word/document.xml";
                PackagePartCollection packagePartCollection = appx.GetParts();
                var vbaProjectPart = (ZipPackagePart)packagePartCollection.FirstOrDefault(u => u.Uri.Equals(name));
                var documentPart = (ZipPackagePart)packagePartCollection.FirstOrDefault(u => u.Uri.Equals(doc));

                var proj = this.GetProject(documentPart);

                var currentCert = this.ReadSignature(vbaProjectPart);
                

                if(currentCert != null)
                {
                    if (cert.Thumbprint.Equals(currentCert.Thumbprint))
                    {
                        this.logger.Debug($"file {file} allready sign {cert.Thumbprint} equals with {currentCert.Thumbprint}");
                        this.Certificate = currentCert;
                    }
                    else
                    {
                        //this.SignProject(vbaProjectPart);
                    }
                }
                else
                {
                    this.logger.Debug($"{file} not signed");
                }
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

        private WordVbaProject GetProject(ZipPackagePart documentPart)
        {
            var vbaProj = new WordVbaProject();
            var codePage = 0;

            if (documentPart == null)
            {
                return null;
            }

            var rel = documentPart.GetRelationshipsByType(schemaRelVba).FirstOrDefault();

            if (rel != null)
            {
                var uri = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                var part = documentPart.Package.GetPart(uri);

                var stream = part.GetStream();
                BinaryReader br = new BinaryReader(stream);

                bool terminate = false;

                while (br.BaseStream.Position < br.BaseStream.Length && terminate == false)
                {
                    uint id = br.ReadUInt16();
                    uint size = br.ReadUInt32();
                    switch (id)
                    {
                        case 0x01:
                            vbaProj.SystemKind = (eSyskind)br.ReadUInt32();
                                uint test = br.ReadUInt32();
                            break;
                        case 0x02:
                            vbaProj.Lcid = (int)br.ReadUInt32();
                            break;
                        case 0x03:
                            this.CodePage = (int)br.ReadUInt16(); //2055 de-CH
                            this.CodePage = 2055;
                            vbaProj.CodePage = this.CodePage;
                            break;
                        case 0x04:
                            vbaProj.Name = GetString(br, size);
                            break;
                        case 0x05:
                            vbaProj.Description = GetUnicodeString(br, size);
                            break;
                        case 0x06:
                            vbaProj.HelpFile1 = GetString(br, size);
                            break;
                        case 0x3D:
                            vbaProj.HelpFile2 = GetString(br, size);
                            break;
                        case 0x07:
                            vbaProj.HelpContextID = (int)br.ReadUInt32();
                            break;
                        case 0x08:
                            vbaProj.LibFlags = (int)br.ReadUInt32();
                            break;
                        case 0x09:
                            vbaProj.MajorVersion = (int)br.ReadUInt32();
                            vbaProj.MinorVersion = (int)br.ReadUInt16();
                            break;
                        case 0x0C:
                            vbaProj.Constants = GetUnicodeString(br, size);
                            break;
                        case 0x0D:
                            uint sizeLibID = br.ReadUInt32();
                            //var regRef = new ExcelVbaReference();
                            //regRef.Name = referenceName;
                            //regRef.ReferenceRecordID = id;
                            //regRef.Libid = GetString(br, sizeLibID);
                            //uint reserved1 = br.ReadUInt32();
                            //ushort reserved2 = br.ReadUInt16();
                            //References.Add(regRef);
                            break;
                        case 0x0E:
                            //var projRef = new ExcelVbaReferenceProject();
                            //projRef.ReferenceRecordID = id;
                            //projRef.Name = referenceName;
                            //sizeLibID = br.ReadUInt32();
                            //projRef.Libid = GetString(br, sizeLibID);
                            //sizeLibID = br.ReadUInt32();
                            //projRef.LibIdRelative = GetString(br, sizeLibID);
                            //projRef.MajorVersion = br.ReadUInt32();
                            //projRef.MinorVersion = br.ReadUInt16();
                            //References.Add(projRef);
                            break;
                        case 0x0F:
                            ushort modualCount = br.ReadUInt16();
                            break;
                        case 0x13:
                            ushort cookie = br.ReadUInt16();
                            break;
                        case 0x14:
                            vbaProj.LcidInvoke = (int)br.ReadUInt32();
                            break;
                        case 0x16:
                            //referenceName = GetUnicodeString(br, size);
                            break;
                        case 0x19:
                            //currentModule = new ExcelVBAModule();
                            //currentModule.Name = GetUnicodeString(br, size);
                            //Modules.Add(currentModule);
                            break;
                        case 0x1A:
                            //currentModule.streamName = GetUnicodeString(br, size);
                            break;
                        case 0x1C:
                            //currentModule.Description = GetUnicodeString(br, size);
                            break;
                        case 0x1E:
                            //currentModule.HelpContext = (int)br.ReadUInt32();
                            break;
                        case 0x21:
                        case 0x22:
                            break;
                        case 0x2B:      //Modul Terminator
                            break;
                        case 0x2C:
                            //currentModule.Cookie = br.ReadUInt16();
                            break;
                        case 0x31:
                            //currentModule.ModuleOffset = br.ReadUInt32();
                            break;
                        case 0x10:
                            terminate = true;
                            break;
                        case 0x30:
                            //var extRef = (ExcelVbaReferenceControl)currentRef;
                            //var sizeExt = br.ReadUInt32();
                            //extRef.LibIdExternal = GetString(br, sizeExt);

                            //uint reserved4 = br.ReadUInt32();
                            //ushort reserved5 = br.ReadUInt16();
                            //extRef.OriginalTypeLib = new Guid(br.ReadBytes(16));
                            //extRef.Cookie = br.ReadUInt32();
                            break;
                        case 0x33:
                            //currentRef = new ExcelVbaReferenceControl();
                            //currentRef.ReferenceRecordID = id;
                            //currentRef.Name = referenceName;
                            //currentRef.Libid = GetString(br, size);
                            //References.Add(currentRef);
                            break;
                        case 0x2F:
                            //var contrRef = (ExcelVbaReferenceControl)currentRef;
                            //contrRef.ReferenceRecordID = id;

                            //var sizeTwiddled = br.ReadUInt32();
                            //contrRef.LibIdTwiddled = GetString(br, sizeTwiddled);
                            //var r1 = br.ReadUInt32();
                            //var r2 = br.ReadUInt16();

                            break;
                        case 0x25:
                            //currentModule.ReadOnly = true;
                            break;
                        case 0x28:
                            //currentModule.Private = true;
                            break;
                        default:
                            break;
                    }
                }

                this.logger.Debug($"CodePage: {this.CodePage}, SystemKind: {vbaProj.SystemKind} CodePage: {codePage}");

                return vbaProj;
            }
            else
            {
                return null;
            }
        }

        private string GetString(BinaryReader br, uint size)
        {
            var str = GetString(br, size, Encoding.GetEncoding(this.CodePage));
            return str;
        }

        private string GetString(BinaryReader br, uint size, Encoding enc)
        {
            if (size > 0)
            {
                byte[] byteTemp = new byte[size];
                byteTemp = br.ReadBytes((int)size);
                var str = enc.GetString(byteTemp);
                return str;
            }
            else
            {
                return "";
            }
        }

        private string GetUnicodeString(BinaryReader br, uint size)
        {
            string s = GetString(br, size);
            int reserved = br.ReadUInt16();
            uint sizeUC = br.ReadUInt32();
            string sUC = GetString(br, sizeUC, Encoding.Unicode);
            return sUC.Length == 0 ? s : sUC;
        }
    }
}
