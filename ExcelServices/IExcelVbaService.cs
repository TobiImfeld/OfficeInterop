namespace ExcelServices
{
    public interface IExcelVbaService
    {
        void SetPathToVbaFiles(string targetDirectory);
        void SignAllVbaExcelFiles(string filePath, string certName);
        void DeleteAllExcelVbaSignatures(string filePath);
        void SignOneVbaExcelFileWithDigitalSignature(string fileName, string certName);
        void DeleteDigitalSignatureFromOneVbaExcelFile(string fileName);
    }
}
