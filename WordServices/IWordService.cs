namespace WordServices
{
    public interface IWordService
    {
        void SignAllWordFiles(string targetDirectory, string certName);
    }
}
