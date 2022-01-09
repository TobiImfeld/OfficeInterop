namespace WordVba
{
    public class WordVbaReferenceCollection : WordVbaCollectionBase<WordVbaReference>
    {
        internal WordVbaReferenceCollection() { }

        public void Add(WordVbaReference Item)
        {
            list.Add(Item);
        }
    }
}
