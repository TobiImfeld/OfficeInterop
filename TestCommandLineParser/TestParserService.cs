using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestCommandLineParser
{
    [TestClass]
    public class TestParserService
    {
        [TestMethod]
        public void IllegalCharacterInPath_OneCommandAndPathOptionWithPathValue()
        {
            var test = new TestingBase();
            var parserService = test
                .InitServices()
                .GetParserService();

            var input = "delsigfromvba -f \"C:\\Users\\sasc\\Desktop\\ICAD_Cbl_material_DRIVE_GAA0014631_FAG.xlsm\"";

            Assert.AreEqual(0, parserService.ParseInput(input));
        }
    }
}
