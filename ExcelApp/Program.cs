using ExcelServices;

namespace ExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ICertificateStoreService certificateStoreService = new CertificateStoreService();
            IExcelService excelService = new ExcelService(certificateStoreService);
            excelService.AddDigitalSignature("TobiOfficeCert");
        }
    }
}
