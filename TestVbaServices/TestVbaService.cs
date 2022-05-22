using Microsoft.Extensions.DependencyInjection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VbaServices;

namespace TestVbaServices
{
    [TestClass]
    public class TestVbaService
    {
        private IServiceProvider serviceProvider;

        [TestMethod]
        public void GetVbaProject_GetCurrentVbaProjectFromExcelFile()
        {
            this.serviceProvider = new ServiceCollection()
                .AddSingleton<IVbaService, VbaService>()
                .BuildServiceProvider();

            var vba = serviceProvider.GetService<IVbaService>();
            var file = "C:\\Temp\\Files\\ATW Ragnatsch - Palfries, Ragnatsch.xlsm";
            vba.GetVbaProject(file);
        }
    }
}
