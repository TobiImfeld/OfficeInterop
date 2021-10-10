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

    //example command: "cert -c CertificateName"
    [Verb("cert", HelpText = "Set certificate name")]
    public class CertificateNameOptions
    {
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
    
    //example command: "signallvba -p FilePath -c CertificateName"
    [Verb("signallvba", HelpText = "Set certificate name")]
    public class SignAllVbaOptions
    {
        [Option('p', "filePath", Required = false)]
        public string FilePath { get; set; }
        [Option('c', "CertName", Required = false)]
        public string CertName { get; set; }
    }

    //example command: "signvbafile -f FileName -c CertificateName"
    [Verb("signvbafile", HelpText = "Sign one excel file with vba project with certificate")]
    public class SignOneVbaExcelFileOptions
    {
        [Option('f', "fileName", Required = false)]
        public string FileName { get; set; }
        [Option('c', "CertName", Required = false)]
        public string CertName { get; set; }
    }

    //example command: "delsigfromvba -f FileName"
    [Verb("delsigfromvba", HelpText = "Delete certificate from specific excel vba project file")]
    public class DeleteSignatureFromOneVbaExcelFileOptions
    {
        [Option('f', "fileName", Required = false)]
        public string FileName { get; set; }
    }

    //example command: "delallvbasig -p FilePath"
    [Verb("delallvbasig", HelpText = "Delete signature from all excel vba project files")]
    public class DeleteAllExcelVbaSignaturesOptions
    {
        [Option('p', "filePath", Required = false)]
        public string FilePath { get; set; }
    }

    //example command: "signalldocxword -p FilePath -c CertificateName"
    [Verb("signalldocxword", HelpText = "Set certificate name")]
    public class SignAllDocxWordOptions
    {
        [Option('p', "filePath", Required = false)]
        public string FilePath { get; set; }
        [Option('c', "CertName", Required = false)]
        public string CertName { get; set; }
    }
}


