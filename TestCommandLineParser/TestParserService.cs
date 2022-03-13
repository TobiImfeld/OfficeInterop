using CommandLineParser;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestCommandLineParser
{
    [TestClass]
    public class TestParserService
    {
        [TestMethod]
        public void IllegalCharacterInPath_OneCommandTwoOptionsWithFileAndCertificateName()
        {
            var test = new TestingBase();
            var parserService = test
                .InitServices()
                .GetParserService();

            var input = "signvbafile -f C:\\Users\\sasc\\Desktop\\ICAD_Cbl_material_DRIVE_GAA0014631_FAG.xlsm -c Frey AG Stans";

            Assert.AreEqual(ExitCode.OK, parserService.ParseInput(input));
        }


        [TestMethod]
        public void IllegalCharacterInPath_OneCommandAndPathOptionWithPathValue()
        {
            var test = new TestingBase();
            var parserService = test
                .InitServices()
                .GetParserService();

            var input = "delsigfromvba -f \"C:\\Users\\sasc\\Desktop\\ICAD_Cbl_material_DRIVE_GAA0014631_FAG.xlsm\"";

            Assert.AreEqual(ExitCode.Error, parserService.ParseInput(input));
        }

        [TestMethod]
        public void IllegalCharacterInPath_OnlyFilePathWithoutCommandAndOptions()
        {
            var test = new TestingBase();
            var parserService = test
                .InitServices()
                .GetParserService();

            var input = "\"C:\\Users\\sasc\\Desktop\\ICAD_Cbl_material_DRIVE_GAA0014631_FAG.xlsm\"";

            Assert.AreEqual(ExitCode.Error, parserService.ParseInput(input));
        }
    }
}
