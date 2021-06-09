using CommandLine;
using Logging;

namespace ExcelApp
{
    public class Program
    {
        static void Main(string[] args)
        {
            var main = MainInitializer
                .Init()
                .AddLogging();

            var logger = main.GetLoggerFactory().Create<Program>();

            logger.Info("Start ExcelApp");

            main.GetExcelService().SetPathToFiles(@"C:\Temp\");
            main.GetExcelService().AddDigitalSignature("TobiOfficeCert");

            logger.Info("Stop ExcelApp");
            main.CloseLogging();
        }
    }
}
