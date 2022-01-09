using System;

namespace WordVba
{
    public class ExcelVbaModuleCollection : WordVbaCollectionBase<WordVbaModule>
    {
        private WordVbaProject project;

        internal ExcelVbaModuleCollection(WordVbaProject project)
        {
            this.project = project;
        }
        internal void Add(WordVbaModule Item)
        {
            list.Add(Item);
        }

        public WordVbaModule AddModule(string Name)
        {
            if (this[Name] != null)
            {
                throw (new ArgumentException("Vba modulename already exist."));
            }

            var m = new WordVbaModule();
            m.Name = Name;
            m.Type = eModuleType.Module;
            m.Attributes.list.Add(new WordVbaModuleAttribute() { Name = "VB_Name", Value = Name, DataType = AttributeDataType.String });
            m.Type = eModuleType.Module;
            list.Add(m);
            return m;
        }
        
        public WordVbaModule AddClass(string Name, bool Exposed)
        {
            var m = new WordVbaModule();
            m.Name = Name;
            m.Type = eModuleType.Class;
            m.Attributes.list.Add(new WordVbaModuleAttribute() { Name = "VB_Name", Value = Name, DataType = AttributeDataType.String });
            m.Attributes.list.Add(new WordVbaModuleAttribute() { Name = "VB_Base", Value = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}", DataType = AttributeDataType.String });
            m.Attributes.list.Add(new WordVbaModuleAttribute() { Name = "VB_GlobalNameSpace", Value = "False", DataType = AttributeDataType.NonString });
            m.Attributes.list.Add(new WordVbaModuleAttribute() { Name = "VB_Creatable", Value = "False", DataType = AttributeDataType.NonString });
            m.Attributes.list.Add(new WordVbaModuleAttribute() { Name = "VB_PredeclaredId", Value = "False", DataType = AttributeDataType.NonString });
            m.Attributes.list.Add(new WordVbaModuleAttribute() { Name = "VB_Exposed", Value = Exposed ? "True" : "False", DataType = AttributeDataType.NonString });
            m.Attributes.list.Add(new WordVbaModuleAttribute() { Name = "VB_TemplateDerived", Value = "False", DataType = AttributeDataType.NonString });
            m.Attributes.list.Add(new WordVbaModuleAttribute() { Name = "VB_Customizable", Value = "False", DataType = AttributeDataType.NonString });

            m.Private = !Exposed;
            list.Add(m);
            return m;
        }
    }
}
