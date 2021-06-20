using CommandLineParser;
using ExcelServices;
using Logging;
using Microsoft.Extensions.DependencyInjection;
using System.Configuration;

namespace ExcelApp
{
    public class MainInitializer
    {
        private ILoggerFactory loggerFactory;
        private ICertificateStoreService certificateStoreService;
        private IExcelService excelService;
        private IParserService parserService;

        private MainInitializer() { }

        public static MainInitializer Init()
        {
            var mainInitalizier = new MainInitializer();
            mainInitalizier.InitServices();

            return mainInitalizier;
        }

        public MainInitializer AddLogging()
        {
            var logFile = ConfigurationManager.AppSettings["LogFilePath"];
            LoggerComponent.InitLogger(logFile);
            return this;
        }

        public ILoggerFactory GetLoggerFactory()
        {
            return this.loggerFactory;
        }

        public IExcelService GetExcelService()
        {
            return this.excelService;
        }

        public IParserService GetParserService()
        {
            return this.parserService;
        }

        public void CloseLogging()
        {
            LoggerComponent.Close();
        }

        private void InitServices()
        {

            var serviceProvider = new ServiceCollection()
                .AddSingleton<ILoggerFactory, LoggerFactory>()
                .AddTransient<ICertificateStoreService, CertificateStoreService>()
                .AddTransient<IExcelService, ExcelService>()
                .AddTransient<IParserService, ParserService>()
                .BuildServiceProvider();

            this.loggerFactory = serviceProvider.GetService<ILoggerFactory>();
            this.certificateStoreService = serviceProvider.GetService<ICertificateStoreService>();
            this.excelService = serviceProvider.GetService<IExcelService>();
            this.parserService = serviceProvider.GetService<IParserService>();
        }
    }
}
