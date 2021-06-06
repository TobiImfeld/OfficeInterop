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

        private MainInitializer() { }

        public static MainInitializer Init()
        {
            var mainInitalizier = new MainInitializer();
            mainInitalizier.InitServices();

            return mainInitalizier;
        }

        public ILoggerFactory GetLoggerFactory()
        {
            return this.loggerFactory;
        }

        public ICertificateStoreService GetCertificateService()
        {
            return this.certificateStoreService;
        }

        public IExcelService GetExcelService()
        {
            return this.excelService;
        }

        private void InitServices()
        {
            var logFile = ConfigurationManager.AppSettings["LogFilePath"];
            LoggerComponent.InitLogger(logFile);

            var serviceProvider = new ServiceCollection()
                .AddSingleton<ILoggerFactory, LoggerFactory>()
                .AddTransient<ICertificateStoreService, CertificateStoreService>()
                .AddTransient<IExcelService, ExcelService>()
                .BuildServiceProvider();

            this.loggerFactory = serviceProvider.GetService<ILoggerFactory>();
            this.certificateStoreService = serviceProvider.GetService<ICertificateStoreService>();
            this.excelService = serviceProvider.GetService<IExcelService>();
        }
    }
}
