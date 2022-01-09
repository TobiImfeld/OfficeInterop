using System;
using System.Linq;

namespace WordVba
{
    public enum eModuleType
    {
        Document = 0,
        Module = 1,
        Class = 2,
        Designer = 3
    }

    internal delegate void ModuleNameChange(string value);

    public class WordVbaModule
    {
        private string name = "";
        ModuleNameChange NameChangeCallback = null;

        internal WordVbaModule()
        {
            Attributes = new WordVbaModuleAttributesCollection();
        }

        internal WordVbaModule(ModuleNameChange nameChangeCallback) : this()
        {
            this.NameChangeCallback = nameChangeCallback;
        }

        public string Name
        {
            get
            {
                return name;
            }
            set
            {
                if (value.Any(c => c > 255))
                {
                    throw (new InvalidOperationException("Vba module names can't contain unicode characters"));
                }
                if (value != name)
                {
                    name = value;
                    StreamName = value;
                    NameChangeCallback?.Invoke(value);
                }
            }
        }

        public string Description { get; set; }
        private string _code = "";

        public string Code
        {
            get
            {
                return _code;
            }
            set
            {
                if (value.StartsWith("Attribute", StringComparison.OrdinalIgnoreCase) || value.StartsWith("VERSION", StringComparison.OrdinalIgnoreCase))
                {
                    throw (new InvalidOperationException("Code can't start with an Attribute or VERSION keyword. Attributes can be accessed through the Attributes collection."));
                }
                _code = value;
            }
        }

        public int HelpContext { get; set; }

        public WordVbaModuleAttributesCollection Attributes { get; internal set; }

        public eModuleType Type { get; internal set; }

        public bool ReadOnly { get; set; }

        public bool Private { get; set; }

        internal string StreamName { get; set; }
        internal ushort Cookie { get; set; }
        internal uint ModuleOffset { get; set; }
        internal string ClassID { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}
