using CommandLine;
using ExcelServices;
using Logging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CommandLineParser
{
    public class ParserService : IParserService
    {
        private readonly ILogger logger;
        private readonly IExcelService excelService;
        private readonly Parser parser;

        public ParserService(ILoggerFactory loggerFactory, IExcelService excelService)
        {
            this.logger = loggerFactory.Create<ParserService>(); ;
            this.excelService = excelService;
            parser = new Parser();
        }

        public int ParseInput(string[] args)
        {
           return this.parser.ParseArguments<PathOptions, CertificateNameOptions, DeleteSignatureOptions, StopOptions>(args)
                .MapResult(
                (PathOptions opts) => this.SetPathToFiles(opts),
                (CertificateNameOptions opts) => this.SetCertificateName(opts),
                (DeleteSignatureOptions opts) => this.DeleteAllDigitalSignatures(opts),
                (StopOptions opts) => this.StopApp(opts),
                errs => this.HandleParseError(errs)
                );
        }

        private int SetPathToFiles(PathOptions options)
        {
            var exitCode = 0;

            var path = options.PathToFiles;
            if (path != null)
            {
                this.logger.Debug($"path= {path}");
                this.excelService.SetPathToFiles(options.PathToFiles);
            }
            else
            {
                exitCode = -1;
                return exitCode;
            }
            
            return exitCode;
        }

        private int SetCertificateName(CertificateNameOptions options)
        {
            var exitCode = 0;

            var certName = options.CertName;
            if (certName != null)
            {
                this.logger.Debug($"certificate name= {certName}");
                this.excelService.AddDigitalSignature(certName);
            }
            else
            {
                exitCode = -1;
                return exitCode;
            }

            return exitCode;
        }

        private int StopApp(StopOptions options)
        {
            var value = options.Stop;
            this.logger.Debug($"stop= {value}");
            return value;
        }

        private int DeleteAllDigitalSignatures(DeleteSignatureOptions options)
        {
            var exitCode = 0;

            var path = options.PathToFiles;
            if (path != null)
            {
                this.logger.Debug($"Delete file signature, path= {path}");
                this.excelService.DeleteAllDigitalSignatures(options.PathToFiles);
            }
            else
            {
                exitCode = -1;
                return exitCode;
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
