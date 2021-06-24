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

            Console.WriteLine("Excel signing app run..");

            var parser = main.GetParserService();
            var run = true;

            while (run)
            {
                var input = Console.ReadLine();
                var exitCode = parser.ParseInput(input.Split());
                if(exitCode == 1)
                {
                    run = false;
                }
            }

            logger.Info("Stop ExcelApp");
            main.CloseLogging();
        }
    }
}
