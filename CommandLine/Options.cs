using CommandLine;

namespace CommandLineParser
{
    //example command: "path -p C:\\TEMP"
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
}
