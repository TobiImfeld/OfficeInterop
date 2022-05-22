using CommandLineParser;
using ExcelServices;
using Logging;
using Microsoft.Extensions.DependencyInjection;
using NSubstitute;
using System;

namespace TestCommandLineParser
{
    public class TestingBase
    {
        private IServiceProvider serviceProvider;

        public TestingBase Init()
        {
            var certificateStoreServiceMock = Substitute.For<ICertificateStoreService>();
            var excelServiceMock = Substitute.For<IExcelService>();
            var fileServiceMock = Substitute.For<IFileService>();
            var vbaExcelServiceMock = Substitute.For<IExcelVbaService>();

            this.serviceProvider = new ServiceCollection()
                .AddSingleton<ILoggerFactory, LoggerFactory>()
                .AddSingleton<IParserService, ParserService>()
                .AddSingleton(certificateStoreServiceMock)
                .AddSingleton(excelServiceMock)
                .AddSingleton(fileServiceMock)
                .AddSingleton(vbaExcelServiceMock)
                .BuildServiceProvider();

            return this;
        }

        public IParserService GetParserService()
        {
            return this.serviceProvider.GetService<IParserService>();
        }
    }
}
