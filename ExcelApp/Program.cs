using System;

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

            //main.GetExcelService().SetPathToFiles(@"C:\Temp\");
            //main.GetExcelService().AddDigitalSignature("TobiOfficeCert");

            var parser = main.GetParserService();

            var args1 = Console.ReadLine();
            parser.ParseInput(args);


            logger.Info("Stop ExcelApp");
            main.CloseLogging();
        }
    }
}
