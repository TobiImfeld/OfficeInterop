namespace ExcelServices
{
    public interface IExcelService
    {
        void AddDigitalSignature(string certName);
        void SetPathToFiles(string filePath);
    }
}
