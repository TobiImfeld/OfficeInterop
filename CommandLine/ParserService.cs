using CommandLine;
using ExcelServices;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CommandLineParser
{
    public class ParserService : IParserService
    {
        private readonly IExcelService excelService;
        private readonly Parser parser;

        public ParserService(IExcelService excelService)
        {
            this.excelService = excelService;
            parser = new Parser();
        }

        public void ParseInput(string[] args)
        {
            this.parser.ParseArguments<PathOptions, CertificateNameOptions>(args)
                .MapResult(
                (PathOptions opts) => this.SetPathToFiles(opts),
                (CertificateNameOptions opts) => this.SetCertificateName(opts),
                errs => 1
                );
        }

        private int SetPathToFiles(PathOptions options)
        {
            var exitCode = 0;

            var path = options.PathToFiles;
            if (path.Equals(null))
            {
                exitCode = -1;
                return exitCode;
            }
            else
            {
                this.excelService.SetPathToFiles(options.PathToFiles);
                Console.WriteLine($"path= {path}");
            }
            
            return exitCode;
        }

        private int SetCertificateName(CertificateNameOptions options)
        {
            var exitCode = 0;

            var certName = options.CertName;
            if (certName.Equals(null))
            {
                exitCode = -1;
                return exitCode;
            }
            else
            {
                this.excelService.AddDigitalSignature(certName);
                Console.WriteLine($"certificate name= {certName}");
            }

            return exitCode;
        }

        private int HandleParseError(IEnumerable<Error> errs)
        {
            var result = -2;
            Console.WriteLine("errors {0}", errs.Count());
            if (errs.Any(x => x is HelpRequestedError || x is VersionRequestedError))
                result = -1;
            Console.WriteLine("Exit code {0}", result);
            return result;
        }
    }
}
