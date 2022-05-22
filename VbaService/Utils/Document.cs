using System.Collections.Generic;
using System.IO;

namespace VbaServices.Utils
{
    internal class Document
    {
        internal class StoragePart
        {
            public StoragePart() { }
            internal Dictionary<string, StoragePart> SubStorage = new Dictionary<string, StoragePart>();
            internal Dictionary<string, byte[]> DataStreams = new Dictionary<string, byte[]>();
        }

        internal StoragePart Storage = null;

        internal Document()
        {
            Storage = new StoragePart();
        }

        internal Document(byte[] bs)
        {
            Read(bs);
        }

        internal void Read(byte[] doc)
        {
            Read(new MemoryStream(doc));
        }

        internal void Read(MemoryStream ms)
        {
            using (var doc = new CompoundDocumentFile(ms))
            {
                Storage = new StoragePart();
                GetStorageAndStreams(Storage, doc.RootItem);
            }
        }

        private void GetStorageAndStreams(StoragePart storage, CompoundDocumentItem parent)
        {
            foreach (var item in parent.Children)
            {
                if (item.ObjectType == 1) //Substorage
                {
                    var part = new StoragePart();
                    storage.SubStorage.Add(item.Name, part);
                    GetStorageAndStreams(part, item);
                }
                else if (item.ObjectType == 2) //Stream
                {
                    storage.DataStreams.Add(item.Name, item.Stream);
                }
            }
        }
    }
}
