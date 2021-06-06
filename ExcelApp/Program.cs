using System.Configuration;
using ExcelServices;
using CommandLine;
using Logging;

namespace ExcelApp
{
    public class Program
    {
        private static readonly ILoggerFactory loggerFactory = new LoggerFactory();

        static void Main(string[] args)
        {
            var logFile = ConfigurationManager.AppSettings["LogFilePath"];
            LoggerComponent.InitLogger(logFile);

            ILogger logger = loggerFactory.Create<Program>();

            logger.Info("Start ExcelApp");

            ICertificateStoreService certificateStoreService = new CertificateStoreService(loggerFactory);
            IExcelService excelService = new ExcelService(loggerFactory, certificateStoreService);
            excelService.SetPathToFiles(@"C:\Temp\");
            excelService.AddDigitalSignature("TobiOfficeCert");

            logger.Info("Stop ExcelApp");
            LoggerComponent.Close();
        }
    }
}
