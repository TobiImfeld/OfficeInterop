using System;
using System.Collections.Generic;

namespace WordVba
{
    public class WordVbaCollectionBase<T> : IEnumerable<T>
    {
        internal protected List<T> list = new List<T>();

        public IEnumerator<T> GetEnumerator()
        {
            return list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return list.GetEnumerator();
        }

        public T this[string Name]
        {
            get
            {
                return list.Find((f) => TypeCompat.GetPropertyValue(f, "Name").ToString().Equals(Name, StringComparison.OrdinalIgnoreCase));
            }
        }
        
        public T this[int Index]
        {
            get
            {
                return list[Index];
            }
        }
        
        public int Count
        {
            get { return list.Count; }
        }

        public bool Exists(string Name)
        {
            return list.Exists((f) => TypeCompat.GetPropertyValue(f, "Name").ToString().Equals(Name, StringComparison.OrdinalIgnoreCase));
        }
        
        public void Remove(T Item)
        {
            list.Remove(Item);
        }
        
        public void RemoveAt(int index)
        {
            list.RemoveAt(index);
        }

        internal void Clear()
        {
            list.Clear();
        }
    }
}
