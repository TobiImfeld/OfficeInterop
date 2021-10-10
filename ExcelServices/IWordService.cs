namespace ExcelServices
{
    public interface IWordService
    {
        void SignAllWordFiles(string targetDirectory, string certName);
    }
}
