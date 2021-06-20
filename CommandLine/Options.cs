using CommandLine;
using System;

namespace CommandLineParser
{
    [Verb("path", HelpText = "Set path to files")]
    public class PathOptions
    {
        [Option('p', "pathToFiles", Required = true)]
        public string PathToFiles { get; set; }
    }

    [Verb("certName", HelpText = "Set certificate name")]
    public class CertificateNameOptions
    {
        [Option('c', "certName", Required = true)]
        public string CertName { get; set; }
    }
}
