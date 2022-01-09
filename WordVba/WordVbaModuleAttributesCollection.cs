using System.Text;

namespace WordVba
{
    public class WordVbaModuleAttributesCollection : WordVbaCollectionBase<WordVbaModuleAttribute>
    {
        internal string GetAttributeText()
        {
            StringBuilder sb = new StringBuilder();

            foreach (var attr in this)
            {
                sb.AppendFormat("Attribute {0} = {1}\r\n", attr.Name, attr.DataType == AttributeDataType.String ? "\"" + attr.Value + "\"" : attr.Value);
            }
            return sb.ToString();
        }
    }
}
