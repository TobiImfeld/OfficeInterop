using System.IO.Packaging;
using System.IO;
using System.Linq;
using System.Text;

namespace VbaProjectUtils
{
    public class WordVBAProject
    {
        public enum eSyskind
        {
            Win16 = 0,
            Win32 = 1,
            Macintosh = 2,
            Win64 = 3
        }

        public eSyskind SystemKind { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string HelpFile1 { get; set; }
        public string HelpFile2 { get; set; }
        public int HelpContextID { get; set; }
        public string Constants { get; set; }
        public int CodePage { get; internal set; }
        internal int LibFlags { get; set; }
        internal int MajorVersion { get; set; }
        internal int MinorVersion { get; set; }
        internal int Lcid { get; set; }
        internal int LcidInvoke { get; set; }
        internal string ProjectID { get; set; }
        internal string ProjectStreamText { get; set; }  
        public ExcelVbaReferenceCollection References { get; set; }

        private ZipPackage zipPackage;

        public WordVBAProject(string file)
        {
            this.zipPackage = this.GetZipPackage(file);
            References = new ExcelVbaReferenceCollection();
            Modules = new ExcelVbaModuleCollection(this);
            var rel = _wb.Part.GetRelationshipsByType(schemaRelVba).FirstOrDefault();
            if (rel != null)
            {
                Uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                Part = _pck.GetPart(Uri);
                GetProject();
            }
            else
            {
                Lcid = 0;
                Part = null;
            }
        }

        private ZipPackage GetZipPackage(string file)
        {
            return (Package.Open(file, FileMode.Open, FileAccess.Read) as ZipPackage);
        }

        private void GetVbaProjectPart()
        {
            var name = "/word/vbaProject.bin";
            PackagePartCollection packagePartCollection = this.zipPackage.GetParts();
            var vbaProjectPart = (ZipPackagePart)packagePartCollection.FirstOrDefault(u => u.Uri.Equals(name));
        }

        /// <summary>
        /// MS-OVBA 2.3.4.1
        /// </summary>
        /// <returns></returns>
        private byte[] CreateDirStream()
        {
            BinaryWriter bw = new BinaryWriter(new MemoryStream());

            /****** PROJECTINFORMATION Record ******/
            bw.Write((ushort)1);        //ID
            bw.Write((uint)4);          //Size
            bw.Write((uint)SystemKind); //SysKind

            bw.Write((ushort)2);        //ID
            bw.Write((uint)4);          //Size
            bw.Write((uint)Lcid);       //Lcid

            bw.Write((ushort)0x14);     //ID
            bw.Write((uint)4);          //Size
            bw.Write((uint)LcidInvoke); //Lcid Invoke

            bw.Write((ushort)3);        //ID
            bw.Write((uint)2);          //Size
            bw.Write((ushort)CodePage); //Codepage

            //ProjectName
            bw.Write((ushort)4);                                     //ID
            bw.Write((uint)Name.Length);                             //Size
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(Name)); //Project Name

            //Description
            bw.Write((ushort)5);                                            //ID
            bw.Write((uint)Description.Length);                             //Size
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(Description)); //Project Name
            bw.Write((ushort)0x40);                                         //ID
            bw.Write((uint)Description.Length * 2);                         //Size
            bw.Write(Encoding.Unicode.GetBytes(Description));               //Project Description

            //Helpfiles
            bw.Write((ushort)6);                                           //ID
            bw.Write((uint)HelpFile1.Length);                              //Size
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(HelpFile1));  //HelpFile1            
            bw.Write((ushort)0x3D);                                        //ID
            bw.Write((uint)HelpFile2.Length);                              //Size
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(HelpFile2));  //HelpFile2

            //Help context id
            bw.Write((ushort)7);            //ID
            bw.Write((uint)4);              //Size
            bw.Write((uint)HelpContextID);  //Help context id

            //Libflags
            bw.Write((ushort)8);            //ID
            bw.Write((uint)4);              //Size
            bw.Write((uint)0);  //Help context id

            //Vba Version
            bw.Write((ushort)9);            //ID
            bw.Write((uint)4);              //Reserved
            bw.Write((uint)MajorVersion);   //Reserved
            bw.Write((ushort)MinorVersion); //Help context id

            //Constants
            bw.Write((ushort)0x0C);           //ID
            bw.Write((uint)Constants.Length);              //Size
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(Constants));     //Help context id
            bw.Write((ushort)0x3C);                                           //ID
            bw.Write((uint)Constants.Length / 2);                             //Size
            bw.Write(Encoding.Unicode.GetBytes(Constants));  //HelpFile2

            /****** PROJECTREFERENCES Record ******/
            foreach (var reference in References)
            {
                WriteNameReference(bw, reference);

                if (reference.ReferenceRecordID == 0x2F)
                {
                    WriteControlReference(bw, reference);
                }
                else if (reference.ReferenceRecordID == 0x33)
                {
                    WriteOrginalReference(bw, reference);
                }
                else if (reference.ReferenceRecordID == 0x0D)
                {
                    WriteRegisteredReference(bw, reference);
                }
                else if (reference.ReferenceRecordID == 0x0E)
                {
                    WriteProjectReference(bw, reference);
                }
            }

            bw.Write((ushort)0x0F);
            bw.Write((uint)0x02);
            bw.Write((ushort)Modules.Count);
            bw.Write((ushort)0x13);
            bw.Write((uint)0x02);
            bw.Write((ushort)0xFFFF);

            foreach (var module in Modules)
            {
                WriteModuleRecord(bw, module);
            }
            bw.Write((ushort)0x10);             //Terminator
            bw.Write((uint)0);

            return VBACompression.CompressPart(((MemoryStream)bw.BaseStream).ToArray());
        }

        private void WriteModuleRecord(BinaryWriter bw, ExcelVBAModule module)
        {
            bw.Write((ushort)0x19);
            bw.Write((uint)module.Name.Length);
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Name));     //Name

            bw.Write((ushort)0x47);
            bw.Write((uint)module.Name.Length * 2);
            bw.Write(Encoding.Unicode.GetBytes(module.Name));                   //Name

            bw.Write((ushort)0x1A);
            bw.Write((uint)module.Name.Length);
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Name));     //Stream Name  

            bw.Write((ushort)0x32);
            bw.Write((uint)module.Name.Length * 2);
            bw.Write(Encoding.Unicode.GetBytes(module.Name));                   //Stream Name

            module.Description = module.Description ?? "";
            bw.Write((ushort)0x1C);
            bw.Write((uint)module.Description.Length);
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Description));     //Description

            bw.Write((ushort)0x48);
            bw.Write((uint)module.Description.Length * 2);
            bw.Write(Encoding.Unicode.GetBytes(module.Description));                   //Description

            bw.Write((ushort)0x31);
            bw.Write((uint)4);
            bw.Write((uint)0);                              //Module Stream Offset (No PerformanceCache)

            bw.Write((ushort)0x1E);
            bw.Write((uint)4);
            bw.Write((uint)module.HelpContext);            //Help context ID

            bw.Write((ushort)0x2C);
            bw.Write((uint)2);
            bw.Write((ushort)0xFFFF);            //Help context ID

            bw.Write((ushort)(module.Type == eModuleType.Module ? 0x21 : 0x22));
            bw.Write((uint)0);

            if (module.ReadOnly)
            {
                bw.Write((ushort)0x25);
                bw.Write((uint)0);              //Readonly
            }

            if (module.Private)
            {
                bw.Write((ushort)0x28);
                bw.Write((uint)0);              //Private
            }

            bw.Write((ushort)0x2B);             //Terminator
            bw.Write((uint)0);
        }

        private void WriteNameReference(BinaryWriter bw, ExcelVbaReference reference)
        {
            //Name record
            bw.Write((ushort)0x16);                                             //ID
            bw.Write((uint)reference.Name.Length);                              //Size
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(reference.Name));  //HelpFile1
            bw.Write((ushort)0x3E);                                             //ID
            bw.Write((uint)reference.Name.Length * 2);                            //Size
            bw.Write(Encoding.Unicode.GetBytes(reference.Name));                //HelpFile2
        }
        private void WriteControlReference(BinaryWriter bw, ExcelVbaReference reference)
        {
            WriteOrginalReference(bw, reference);

            bw.Write((ushort)0x2F);
            var controlRef = (ExcelVbaReferenceControl)reference;
            bw.Write((uint)(4 + controlRef.LibIdTwiddled.Length + 4 + 2));    // Size of SizeOfLibidTwiddled, LibidTwiddled, Reserved1, and Reserved2.
            bw.Write((uint)controlRef.LibIdTwiddled.Length);                              //Size            
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(controlRef.LibIdTwiddled));  //LibID
            bw.Write((uint)0);      //Reserved1
            bw.Write((ushort)0);    //Reserved2
            WriteNameReference(bw, reference);  //Name record again
            bw.Write((ushort)0x30); //Reserved3
            bw.Write((uint)(4 + controlRef.LibIdExternal.Length + 4 + 2 + 16 + 4));    //Size of SizeOfLibidExtended, LibidExtended, Reserved4, Reserved5, OriginalTypeLib, and Cookie
            bw.Write((uint)controlRef.LibIdExternal.Length);                              //Size            
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(controlRef.LibIdExternal));  //LibID
            bw.Write((uint)0);      //Reserved4
            bw.Write((ushort)0);    //Reserved5
            bw.Write(controlRef.OriginalTypeLib.ToByteArray());
            bw.Write((uint)controlRef.Cookie);      //Cookie
        }

        private void WriteOrginalReference(BinaryWriter bw, ExcelVbaReference reference)
        {
            bw.Write((ushort)0x33);
            bw.Write((uint)reference.Libid.Length);
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(reference.Libid));  //LibID
        }
        private void WriteProjectReference(BinaryWriter bw, ExcelVbaReference reference)
        {
            bw.Write((ushort)0x0E);
            var projRef = (ExcelVbaReferenceProject)reference;
            bw.Write((uint)(4 + projRef.Libid.Length + 4 + projRef.LibIdRelative.Length + 4 + 2));
            bw.Write((uint)projRef.Libid.Length);
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(projRef.Libid));  //LibAbsolute
            bw.Write((uint)projRef.LibIdRelative.Length);
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(projRef.LibIdRelative));  //LibIdRelative
            bw.Write(projRef.MajorVersion);
            bw.Write(projRef.MinorVersion);
        }

        private void WriteRegisteredReference(BinaryWriter bw, ExcelVbaReference reference)
        {
            bw.Write((ushort)0x0D);
            bw.Write((uint)(4 + reference.Libid.Length + 4 + 2));
            bw.Write((uint)reference.Libid.Length);
            bw.Write(Encoding.GetEncoding(CodePage).GetBytes(reference.Libid));  //LibID            
            bw.Write((uint)0);      //Reserved1
            bw.Write((ushort)0);    //Reserved2
        }

    }
}
