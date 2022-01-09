namespace WordVba
{
    public class WordVbaReferenceProject : WordVbaReference
    {
        public WordVbaReferenceProject()
        {
            ReferenceRecordID = 0x0E;
        }
        /// <summary>
        /// LibIdRelative
        /// For more info check MS-OVBA 2.1.1.8 LibidReference and 2.3.4.2.2 PROJECTREFERENCES
        /// </summary>
        public string LibIdRelative { get; set; }

        public uint MajorVersion { get; set; }

        public ushort MinorVersion { get; set; }
    }
}
