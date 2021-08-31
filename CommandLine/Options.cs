using CommandLine;

namespace CommandLineParser
{
    //example command: "path -p C:\Temp"
    [Verb("path", HelpText = "Set path to files")]
    public class PathOptions
    {
        [Option('p', "PathToFiles", Required = false)]
        public string PathToFiles { get; set; }
    }

    //example command: "vbapath -p C:\Temp"
    [Verb("vbapath", HelpText = "Set path to files")]
    public class VbaPathOptions
    {
        [Option('p', "PathToVbaFiles", Required = false)]
        public string PathToVbaFiles { get; set; }
    }

    //example command: "cert -c CertificateName"
    [Verb("cert", HelpText = "Set certificate name")]
    public class CertificateNameOptions
    {
        [Option('c', "CertName", Required = false)]
        public string CertName { get; set; }
    }

    //example command: "signvba -c CertificateName"
    [Verb("signvba", HelpText = "Set certificate name")]
    public class SignVbaOptions
    {
        [Option('c', "CertName", Required = false)]
        public string CertName { get; set; }
    }

    //example command: "signvbafile -p FilePath -c CertificateName"
    [Verb("signvbafile", HelpText = "Sign one excel file with vba project with certificate")]
    public class SignOneExcelFileOptions
    {
        [Option('p', "filePath", Required = false)]
        public string FilePath { get; set; }
        [Option('c', "CertName", Required = false)]
        public string CertName { get; set; }
    }

    //example command: "stop -s 1"
    [Verb("stop", HelpText = "Stops the app")]
    public class StopOptions
    {
        [Option('s', "Stop", Required = false)]
        public int Stop { get; set; }
    }

    //example command: "delSig -p C:\Temp "
    [Verb("delSig", HelpText = "Delete all certificates from path")]
    public class DeleteSignatureOptions
    {
        [Option('p', "PathToFiles", Required = false)]
        public string PathToFiles { get; set; }
    }

    //example command: "delSig -p C:\Temp "
    [Verb("delsigfromvba", HelpText = "Delete certificate from specific excel vba project file")]
    public class DeleteSignatureFromFileOptions
    {
        [Option('f', "fileName", Required = false)]
        public string FileName { get; set; }
    }
}
