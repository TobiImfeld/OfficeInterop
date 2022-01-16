using CommandLineParser;
using ExcelServices;
using Logging;
using Microsoft.Extensions.DependencyInjection;

namespace TestCommandLineParser
{
    public class TestingBase
    {
        private ServiceProvider serviceProvider;

        public TestingBase InitServices()
        {
            this.serviceProvider = new ServiceCollection()
                .AddSingleton<ILoggerFactory, LoggerFactory>()
                .AddTransient<ICertificateStoreService, CertificateStoreService>() //service mocken! kein echter zugriff auf certificate store!
                .AddTransient<IExcelService, ExcelService>()
                .AddTransient<IParserService, ParserService>()
                .AddTransient<IFileService, FileService>()
                .AddTransient<IExcelVbaService, ExcelVbaService>()
                .BuildServiceProvider();

            return this;       
        }

        public IParserService GetParserService()
        {
            return this.serviceProvider.GetService<IParserService>();
        }
    }
}
