using CommandLineParser;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestCommandLineParser
{
    [TestClass]
    public class TestParserService
    {
        private IParserService parserService;

        [TestMethod]
        public void IllegalCharacterInPath_OneCommandTwoOptionsWithFileAndCertificateName()
        {
            this.BeforeTest();

            var input = "signvbafile -f C:\\Users\\sasc\\Desktop\\ICAD_Cbl_material_DRIVE_GAA0014631_FAG.xlsm -c Frey AG Stans";

            Assert.AreEqual(ExitCode.OK, parserService.ParseInput(input));
        }

        [TestMethod]
        public void IllegalCharacterInPath_OneCommandAndPathOptionWithPathValue()
        {
            this.BeforeTest();

            var input = "delsigfromvba -f \"C:\\Users\\sasc\\Desktop\\ICAD_Cbl_material_DRIVE_GAA0014631_FAG.xlsm\"";

            Assert.AreEqual(ExitCode.Error, parserService.ParseInput(input));
        }

        [TestMethod]
        public void IllegalCharacterInPath_OnlyFilePathWithoutCommandAndOptions()
        {
            this.BeforeTest();

            var input = "\"C:\\Users\\sasc\\Desktop\\ICAD_Cbl_material_DRIVE_GAA0014631_FAG.xlsm\"";

            Assert.AreEqual(ExitCode.Error, parserService.ParseInput(input));
        }

        private void BeforeTest()
        {
            var test = new TestingBase();
            this.parserService = test
                .Init()
                .GetParserService();
        }
    }
}
