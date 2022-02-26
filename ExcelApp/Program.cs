using CommandLineParser;
using System;
using System.Configuration;

namespace ExcelApp
{
    public class Program
    {
        static void Main(string[] args)
        {
            var path = ConfigurationManager.AppSettings["LogFilePath"];
            var version = ConfigurationManager.AppSettings["Version"];
            Console.WriteLine($"Excel Signing App v{version}");

            var main = MainInitializer
                .Init()
                .AddLogging(path);

            var logger = main.GetLoggerFactory().Create<Program>();

            logger.Info("Start ExcelApp");

            Console.WriteLine("Enter command");

            var parser = main.GetParserService();
            var run = true;

            while (run)
            {
                var input = Console.ReadLine();
                var exitCode = parser.ParseInput(input);
                if(exitCode == ExitCode.Stop)
                {
                    run = false;
                }
            }

            logger.Info("Stop ExcelApp");
            main.CloseLogging();
        }
    }
}
