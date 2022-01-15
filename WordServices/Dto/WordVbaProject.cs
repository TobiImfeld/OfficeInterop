namespace WordServices.Dto
{
    public class WordVbaProject
    {
        public eSyskind SystemKind { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string HelpFile1 { get; set; }
        public string HelpFile2 { get; set; }
        public int HelpContextID { get; set; }
        public string Constants { get; set; }
        public int CodePage { get; set; }
        public int LibFlags { get; set; }
        public int MajorVersion { get; set; }
        public int MinorVersion { get; set; }
        public int Lcid { get; set; }
        public int LcidInvoke { get; set; }
        public string ProjectID { get; set; }
        public string ProjectStreamText { get; set; }      
    }
}
