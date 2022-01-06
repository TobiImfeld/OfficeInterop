using System.Linq;
using System.IO.Packaging;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Pkcs;
using Logging;

namespace WordServices
{
    public class WordVbaSignatureService : IWordVbaSignatureService
    {
        public X509Certificate2 Certificate { get; internal set; }

        private readonly ILogger logger;
        private const string schemaRelVbaSignature = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";
        private const string schemaRelVba = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

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
                var document = (ZipPackagePart)packagePartCollection.FirstOrDefault(u => u.Uri.Equals(doc));

                this.GetProject(document);

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

        private void GetProject(ZipPackagePart document)
        {
            int codePage = 0;

            var rel = document.GetRelationshipsByType(schemaRelVba).FirstOrDefault();

            if (rel != null)
            {
                var uri = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                var part = document.Package.GetPart(uri);

                var stream = part.GetStream();
                BinaryReader br = new BinaryReader(stream);

                bool terminate = false;

                var len = br.BaseStream.Length;

                while (br.BaseStream.Position < br.BaseStream.Length && terminate == false)
                {
                    uint id = br.ReadUInt32();
                    uint size = br.ReadUInt32();
                    switch (id)
                    {
                        case 0x03:
                            codePage = (int)br.ReadUInt16();
                            break;
                        case 0x10:
                            terminate = true;
                            break;
                        default:
                            break;
                    }

                }
            }
        }
                //private byte[] SignProject(ZipPackagePart vbaProjectPart)
                //{
                //    if (!Certificate.HasPrivateKey)
                //    {
                //        //throw (new InvalidOperationException("The certificate doesn't have a private key"));
                //        Certificate = null;
                //        return null;
                //    }
                //    var hash = GetContentHash(vbaProjectPart);

                //    //BinaryWriter bw = new BinaryWriter(new MemoryStream());
                //    //bw.Write((byte)0x30); //Constructed Type 
                //    //bw.Write((byte)0x32); //Total length
                //    //bw.Write((byte)0x30); //Constructed Type 
                //    //bw.Write((byte)0x0E); //Length SpcIndirectDataContent
                //    //bw.Write((byte)0x06); //Oid Tag Indentifier 
                //    //bw.Write((byte)0x0A); //Lenght OId
                //    //bw.Write(new byte[] { 0x2B, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x1D }); //Encoded Oid 1.3.6.1.4.1.311.2.1.29
                //    //bw.Write((byte)0x04);   //Octet String Tag Identifier
                //    //bw.Write((byte)0x00);   //Zero length

                //    //bw.Write((byte)0x30); //Constructed Type (DigestInfo)
                //    //bw.Write((byte)0x20); //Length DigestInfo
                //    //bw.Write((byte)0x30); //Constructed Type (Algorithm)
                //    //bw.Write((byte)0x0C); //length AlgorithmIdentifier
                //    //bw.Write((byte)0x06); //Oid Tag Indentifier 
                //    //bw.Write((byte)0x08); //Lenght OId
                //    //bw.Write(new byte[] { 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x02, 0x05 }); //Encoded Oid for 1.2.840.113549.2.5 (AlgorithmIdentifier MD5)
                //    //bw.Write((byte)0x05);   //Null type identifier
                //    //bw.Write((byte)0x00);   //Null length
                //    //bw.Write((byte)0x04);   //Octet String Identifier
                //    //bw.Write((byte)hash.Length);   //Hash length
                //    //bw.Write(hash);                //Content hash

                //    //ContentInfo contentInfo = new ContentInfo(((MemoryStream)bw.BaseStream).ToArray());
                //    //contentInfo.ContentType.Value = "1.3.6.1.4.1.311.2.1.4";

                //    //Verifier = new SignedCms(contentInfo);
                //    //var signer = new CmsSigner(Certificate);
                //    //Verifier.ComputeSignature(signer, false);
                //    //return Verifier.Encode();
                //    return null;
                //}

                //private byte[] GetContentHash(ZipPackagePart vbaProjectPart)
                //{
                //    //MS-OVBA 2.4.2
                //    var enc = System.Text.Encoding.GetEncoding(vbaProjectPart.CodePage);
                //    BinaryWriter bw = new BinaryWriter(new MemoryStream());
                //    bw.Write(enc.GetBytes(vbaProjectPart.Name));
                //    bw.Write(enc.GetBytes(vbaProjectPart.Constants));
                //    foreach (var reference in vbaProjectPart.References)
                //    {
                //        if (reference.ReferenceRecordID == 0x0D)
                //        {
                //            bw.Write((byte)0x7B);
                //        }
                //        if (reference.ReferenceRecordID == 0x0E)
                //        {
                //            //var r = (ExcelVbaReferenceProject)reference;
                //            //BinaryWriter bwTemp = new BinaryWriter(new MemoryStream());
                //            //bwTemp.Write((uint)r.Libid.Length);
                //            //bwTemp.Write(enc.GetBytes(r.Libid));              
                //            //bwTemp.Write((uint)r.LibIdRelative.Length);
                //            //bwTemp.Write(enc.GetBytes(r.LibIdRelative));
                //            //bwTemp.Write(r.MajorVersion);
                //            //bwTemp.Write(r.MinorVersion);
                //            foreach (byte b in BitConverter.GetBytes((uint)reference.Libid.Length))  //Length will never be an UInt with 4 bytes that aren't 0 (> 0x00FFFFFF), so no need for the rest of the properties.
                //            {
                //                if (b != 0)
                //                {
                //                    bw.Write(b);
                //                }
                //                else
                //                {
                //                    break;
                //                }
                //            }
                //        }
                //    }
                //    foreach (var module in proj.Modules)
                //    {
                //        var lines = module.Code.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                //        foreach (var line in lines)
                //        {
                //            if (!line.StartsWith("attribute", StringComparison.OrdinalIgnoreCase))
                //            {
                //                bw.Write(enc.GetBytes(line));
                //            }
                //        }
                //    }
                //    var buffer = (bw.BaseStream as MemoryStream).ToArray();
                //    var hp = System.Security.Cryptography.MD5.Create();
                //    return hp.ComputeHash(buffer);
                //}
    }
}
