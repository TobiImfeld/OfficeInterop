namespace ExcelServices
{
    public interface IExcelVbaService
    {
        void SetPathToVbaFiles(string targetDirectory);
        void AddDigitalSignatureToVbaMacro(string certName);
        void SignOneExcelFileWithDigitalSignature(string fileName, string certName);
        void DeleteOneDigitalSignatureFromExcelFile(string fileName);
    }
}
