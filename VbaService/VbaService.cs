using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using VbaServices.Utils;

namespace VbaServices
{
    public class VbaService : IVbaService
    {
        private eSyskind SystemKind { get; set; }
        private int CodePage { get; set; }
        private int Lcid { get; set; }
        private string Name { get; set; }

        public void GetVbaProject(string file)
        {
            var name = "/xl/vbaProject.bin";
            byte[] vba;
            var zipPackage = this.GetZipPackage(file);

            PackagePartCollection packagePartCollection = zipPackage.GetParts();
            var vbaProjectPart = (ZipPackagePart)packagePartCollection.FirstOrDefault(u => u.Uri.Equals(name));

            var stream = vbaProjectPart.GetStream();
            vba = new byte[stream.Length];
            stream.Read(vba, 0, (int)stream.Length);

            var document = new Document(vba);
            this.ReadDirStream(document);

            var projectStreamText = Encoding.GetEncoding(1250).GetString(document.Storage.DataStreams["PROJECT"]);
        }

        private ZipPackage GetZipPackage(string file)
        {
            return Package.Open(file, FileMode.Open, FileAccess.Read) as ZipPackage;
        }

        private void ReadDirStream(Document document)
        {
            byte[] dir = VBACompression.DecompressPart(document.Storage.SubStorage["VBA"].DataStreams["dir"]);
            MemoryStream ms = new MemoryStream(dir);
            BinaryReader br = new BinaryReader(ms);

            bool terminate = false;
            br.BaseStream.Position = 0;
            while(br.BaseStream.Position < br.BaseStream.Length && terminate == false)
            {
                ushort id = br.ReadUInt16();
                uint size = br.ReadUInt32();
                switch (id)
                {
                    case 0x01:
                        SystemKind = (eSyskind)br.ReadUInt32();
                        break;
                    case 0x02:
                        Lcid = (int)br.ReadUInt32();
                        break;
                    case 0x03:
                        CodePage = br.ReadUInt16(); //CodePage not found!
                        break;
                    case 0x04:
                        Name = this.GetString(br, size);
                        break;
                    case 0x10:
                        terminate = true;
                        break;
                }
            }

        }

        private string GetString(BinaryReader br, uint size)
        {
            int codePage = 1250;

            if(this.CodePage == 0)
            {
                this.CodePage = codePage;
            }

            return GetString(br, size, Encoding.GetEncoding(this.CodePage));
        }
        private string GetString(BinaryReader br, uint size, Encoding enc)
        {
            if (size > 0)
            {
                byte[] byteTemp = new byte[size];
                byteTemp = br.ReadBytes((int)size);
                return enc.GetString(byteTemp);
            }
            else
            {
                return "";
            }
        }

        //    private void ReadDirStream()
        //    {
        //        byte[] dir = VBACompression.DecompressPart(Document.Storage.SubStorage["VBA"].DataStreams["dir"]);
        //        MemoryStream ms = new MemoryStream(dir);
        //        BinaryReader br = new BinaryReader(ms);
        //        ExcelVbaReference currentRef = null;
        //        string referenceName = "";
        //        ExcelVBAModule currentModule = null;
        //        bool terminate = false;
        //        while (br.BaseStream.Position < br.BaseStream.Length && terminate == false)
        //        {
        //            ushort id = br.ReadUInt16();
        //            uint size = br.ReadUInt32();
        //            switch (id)
        //            {
        //                case 0x01:
        //                    SystemKind = (eSyskind)br.ReadUInt32();
        //                    break;
        //                case 0x02:
        //                    Lcid = (int)br.ReadUInt32();
        //                    break;
        //                case 0x03:
        //                    CodePage = (int)br.ReadUInt16();
        //                    break;
        //                case 0x04:
        //                    Name = GetString(br, size);
        //                    break;
        //                case 0x05:
        //                    Description = GetUnicodeString(br, size);
        //                    break;
        //                case 0x06:
        //                    HelpFile1 = GetString(br, size);
        //                    break;
        //                case 0x3D:
        //                    HelpFile2 = GetString(br, size);
        //                    break;
        //                case 0x07:
        //                    HelpContextID = (int)br.ReadUInt32();
        //                    break;
        //                case 0x08:
        //                    LibFlags = (int)br.ReadUInt32();
        //                    break;
        //                case 0x09:
        //                    MajorVersion = (int)br.ReadUInt32();
        //                    MinorVersion = (int)br.ReadUInt16();
        //                    break;
        //                case 0x0C:
        //                    Constants = GetUnicodeString(br, size);
        //                    break;
        //                case 0x0D:
        //                    uint sizeLibID = br.ReadUInt32();
        //                    var regRef = new ExcelVbaReference();
        //                    regRef.Name = referenceName;
        //                    regRef.ReferenceRecordID = id;
        //                    regRef.Libid = GetString(br, sizeLibID);
        //                    uint reserved1 = br.ReadUInt32();
        //                    ushort reserved2 = br.ReadUInt16();
        //                    References.Add(regRef);
        //                    break;
        //                case 0x0E:
        //                    var projRef = new ExcelVbaReferenceProject();
        //                    projRef.ReferenceRecordID = id;
        //                    projRef.Name = referenceName;
        //                    sizeLibID = br.ReadUInt32();
        //                    projRef.Libid = GetString(br, sizeLibID);
        //                    sizeLibID = br.ReadUInt32();
        //                    projRef.LibIdRelative = GetString(br, sizeLibID);
        //                    projRef.MajorVersion = br.ReadUInt32();
        //                    projRef.MinorVersion = br.ReadUInt16();
        //                    References.Add(projRef);
        //                    break;
        //                case 0x0F:
        //                    ushort modualCount = br.ReadUInt16();
        //                    break;
        //                case 0x13:
        //                    ushort cookie = br.ReadUInt16();
        //                    break;
        //                case 0x14:
        //                    LcidInvoke = (int)br.ReadUInt32();
        //                    break;
        //                case 0x16:
        //                    referenceName = GetUnicodeString(br, size);
        //                    break;
        //                case 0x19:
        //                    currentModule = new ExcelVBAModule();
        //                    currentModule.Name = GetUnicodeString(br, size);
        //                    Modules.Add(currentModule);
        //                    break;
        //                case 0x1A:
        //                    currentModule.streamName = GetUnicodeString(br, size);
        //                    break;
        //                case 0x1C:
        //                    currentModule.Description = GetUnicodeString(br, size);
        //                    break;
        //                case 0x1E:
        //                    currentModule.HelpContext = (int)br.ReadUInt32();
        //                    break;
        //                case 0x21:
        //                case 0x22:
        //                    break;
        //                case 0x2B:      //Modul Terminator
        //                    break;
        //                case 0x2C:
        //                    currentModule.Cookie = br.ReadUInt16();
        //                    break;
        //                case 0x31:
        //                    currentModule.ModuleOffset = br.ReadUInt32();
        //                    break;
        //                case 0x10:
        //                    terminate = true;
        //                    break;
        //                case 0x30:
        //                    var extRef = (ExcelVbaReferenceControl)currentRef;
        //                    var sizeExt = br.ReadUInt32();
        //                    extRef.LibIdExternal = GetString(br, sizeExt);

        //                    uint reserved4 = br.ReadUInt32();
        //                    ushort reserved5 = br.ReadUInt16();
        //                    extRef.OriginalTypeLib = new Guid(br.ReadBytes(16));
        //                    extRef.Cookie = br.ReadUInt32();
        //                    break;
        //                case 0x33:
        //                    currentRef = new ExcelVbaReferenceControl();
        //                    currentRef.ReferenceRecordID = id;
        //                    currentRef.Name = referenceName;
        //                    currentRef.Libid = GetString(br, size);
        //                    References.Add(currentRef);
        //                    break;
        //                case 0x2F:
        //                    var contrRef = (ExcelVbaReferenceControl)currentRef;
        //                    contrRef.ReferenceRecordID = id;

        //                    var sizeTwiddled = br.ReadUInt32();
        //                    contrRef.LibIdTwiddled = GetString(br, sizeTwiddled);
        //                    var r1 = br.ReadUInt32();
        //                    var r2 = br.ReadUInt16();

        //                    break;
        //                case 0x25:
        //                    currentModule.ReadOnly = true;
        //                    break;
        //                case 0x28:
        //                    currentModule.Private = true;
        //                    break;
        //                default:
        //                    break;
        //            }
        //        }
        //    }
        //}
    }
}
