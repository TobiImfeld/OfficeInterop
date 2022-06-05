using System;
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
        /// <summary>
        /// Codepage for encoding. Default is current regional setting.
        /// </summary>
        private int Lcid { get; set; }
        private int CodePage { get; set; }
        private string Name { get; set; }
        private string Description { get; set; }
        /// <summary>
        /// A helpfile
        /// </summary>
        private string HelpFile1 { get; set; }
        /// <summary>
        /// Secondary helpfile
        /// </summary>
        private string HelpFile2 { get; set; }
        /// <summary>
        /// Context if refering the helpfile
        /// </summary>
        private int HelpContextID { get; set; }
        /// <summary>
        /// Conditional compilation constants
        /// </summary>
        private string Constants { get; set; }
        internal int LibFlags { get; set; }
        internal int MajorVersion { get; set; }
        internal int MinorVersion { get; set; }

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
            var currentPosition = 0;

            for (int i = 0; i < dir.Length; i++)
            {
                if (0x003 == dir[i])
                {
                    var codePageByteOne = dir[i + 5];
                    var codePageByteTwo = dir[i + 6];

                    UInt16 combined = (ushort)(codePageByteOne << 8 | codePageByteTwo);
                    break;
                }
            }

            /// New try: Search in the dir[] for the id's and if there are found, then use the definitions from documentation
            /// to get the corresponding values, by bytes.
            var syskind = this.search(dir, dir.Length, 0x001);
            var lcid = this.search(dir, dir.Length, 0x002);
            var codePageId = this.search(dir, dir.Length, 0x003);


            while (br.BaseStream.Position < br.BaseStream.Length && terminate == false)
            {
                ///
                /// Problem: the br.BaseStream.Position, don't get the byte at pos = 40 in the dir-array
                /// to read the codepage correct. The basic problem is to read the id's from the correct positon in the dir-array.
                /// I can read the correct codepage if I set the br.BaseStream.Position = 40;
                /// br.BaseStream.Position = 40;
                /// siehe: https://jonskeet.uk/csharp/readbinary.html
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
                    case 0x05:
                        Description = GetUnicodeString(br, size);
                        break;
                    case 0x06:
                        HelpFile1 = GetString(br, size);
                        break;
                    case 0x3D:
                        HelpFile2 = GetString(br, size);
                        break;
                    case 0x07:
                        HelpContextID = (int)br.ReadUInt32();
                        break;
                    case 0x08:
                        LibFlags = (int)br.ReadUInt32();
                        break;
                    case 0x09:
                        MajorVersion = (int)br.ReadUInt32();
                        MinorVersion = (int)br.ReadUInt16();
                        break;
                    case 0x0C:
                        Constants = GetUnicodeString(br, size);
                        break;
                    case 0x2B:      //Modul Terminator
                        break;
                    case 0x10:
                        terminate = true;
                        break;
                    default:
                        break;
                }

            }

        }

        public static void ReadWholeArray(Stream stream, byte[] data)
        {
            int offset = 0;
            int remaining = data.Length;
            while (remaining > 0)
            {
                int read = stream.Read(data, offset, remaining);
                if (read <= 0)
                    throw new EndOfStreamException
                        (String.Format("End of stream reached with {0} bytes left to read", remaining));
                remaining -= read;
                offset += read;
            }
        }

        public ResultDto search(byte[] arr, int n, byte x)
        {
            // 1st comparison
            if (arr[n - 1] == x)
                return new ResultDto()
                {
                    Message = "Found",
                    Value = arr[n - 1],
                };

            byte backup = arr[n - 1];
            arr[n - 1] = x;

            // no termination condition and thus
            // no comparison
            for (int i = 0; ; i++)
            {

                // this would be executed at-most n times
                // and therefore at-most n comparisons
                if (arr[i] == x)
                {

                    // replace arr[n-1] with its actual element
                    // as in original 'arr[]'
                    arr[n - 1] = backup;

                    // if 'x' is found before the '(n-1)th'
                    // index, then it is present in the array
                    // final comparison
                    if (i < n - 1)
                        return new ResultDto()
                        {
                            Message = "Found",
                            Value = arr[i],
                        };

                    // else not present in the array
                    return new ResultDto()
                    {
                        Message = "Not Found"
                    };
                }
            }
        }



        private string GetString(BinaryReader br, uint size)
        {
            //int codePage = 1250;

            //if(this.CodePage == 0)
            //{
            //    this.CodePage = codePage;
            //}

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

        private string GetUnicodeString(BinaryReader br, uint size)
        {
            string s = GetString(br, size);
            uint sizeUC = br.ReadUInt32();
            string sUC = GetString(br, sizeUC, Encoding.Unicode);
            return sUC.Length == 0 ? s : sUC;
        }
    }

    public class ResultDto
    {
        public string Message { get; set; }
        public byte Value { get; set; }
    }
}
