using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestCommandLineParser
{
    [TestClass]
    public class TestParserService
    {
        [TestMethod]
        public void IsInputCorrect()
        {
            var test = new TestingBase();
            var parserService = test
                .InitServices()
                .GetParserService();

            var input = "F:\\Firma\\Firma\\SERVICE\\Anlagen\\05_GBK\\GBK Vercorin - Sigeroulaz - Crêt du Midi, Vercorin\\GBK Vercorin - Sigeroulaz - Crêt du Midi, Vercorin.xlsm";

            Assert.AreEqual(0, parserService.ParseInput(input));
        }
    }
}
