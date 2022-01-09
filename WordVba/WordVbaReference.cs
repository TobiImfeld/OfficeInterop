namespace WordVba
{
    public class WordVbaReference
    {
        public WordVbaReference()
        {
            ReferenceRecordID = 0xD;
        }

        public int ReferenceRecordID { get; internal set; }

        public string Name { get; set; }

        public string Libid { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}
