namespace WordVba
{
    public enum AttributeDataType
    {
        String = 0,
        NonString = 1
    }

    public class WordVbaModuleAttribute
    {
        internal WordVbaModuleAttribute() { }

        public string Name { get; internal set; }

        public AttributeDataType DataType { get; internal set; }

        public string Value { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}
