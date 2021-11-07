using CommandLineParser;
using Common;
using ExcelServices;
using Logging;
using Microsoft.Extensions.DependencyInjection;
using WordServices;

namespace ExcelApp
{
    public class MainInitializer
    {
        private ILoggerFactory loggerFactory;
        private IParserService parserService;

        private MainInitializer() { }

        public static MainInitializer Init()
        {
            var mainInitalizier = new MainInitializer();
            mainInitalizier.InitServices();

            return mainInitalizier;
        }

        public MainInitializer AddLogging(string logFilePath)
        {
            LoggerComponent.InitLogger(logFilePath);
            return this;
        }

        public ILoggerFactory GetLoggerFactory()
        {
            return this.loggerFactory;
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
                .AddTransient<IFileService, FileService>()
                .AddTransient<IExcelVbaService, ExcelVbaService>()
                .AddTransient<IWordService, WordService>()
                .BuildServiceProvider();

            this.loggerFactory = serviceProvider.GetService<ILoggerFactory>();
            this.parserService = serviceProvider.GetService<IParserService>();
        }
    }
}
