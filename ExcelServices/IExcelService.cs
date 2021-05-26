namespace ExcelServices
{
    public interface IExcelService
    {
        void AddDigitalSignature(string filePath, string certName);
    }
}
