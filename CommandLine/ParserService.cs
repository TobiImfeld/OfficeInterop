using CommandLine;
using ExcelServices;
using Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace CommandLineParser
{
    public class ParserService : IParserService
    {
        private readonly ILogger logger;
        private readonly IExcelService excelService;
        private readonly IExcelVbaService excelVbaService;
        private readonly Parser parser;

        public ParserService(
            ILoggerFactory loggerFactory,
            IExcelService excelService,
            IExcelVbaService excelVbaService)
        {
            this.logger = loggerFactory.Create<ParserService>(); ;
            this.excelService = excelService;
            this.excelVbaService = excelVbaService;
            parser = new Parser();
        }

        public int ParseInput(string input)
        {
            var args = this.SplitInputStringIntoArgumentsArray(input);
            return this.ParseInputArguments(args);
        }

        private int ParseInputArguments(string[] args)
        {
            return this.parser.ParseArguments<
                PathOptions,
                CertificateNameOptions,
                DeleteSignatureOptions,
                StopOptions,
                SignAllVbaOptions,
                SignOneVbaExcelFileOptions,
                DeleteSignatureFromOneVbaExcelFileOptions,
                DeleteAllExcelVbaSignaturesOptions>(args)
                 .MapResult(
                 (PathOptions opts) => this.SetPathToFiles(opts),
                 (CertificateNameOptions opts) => this.SetCertificateName(opts),
                 (DeleteSignatureOptions opts) => this.DeleteAllDigitalSignatures(opts),
                 (StopOptions opts) => this.StopApp(opts),
                 (SignAllVbaOptions opts) => this.SignAllVbaExcelFiles(opts),
                 (SignOneVbaExcelFileOptions opts) => this.SignOneVbaExcelFile(opts),
                 (DeleteSignatureFromOneVbaExcelFileOptions opts) => this.DeleteDigitalSignatureFromOneVbaExcelFile(opts),
                 (DeleteAllExcelVbaSignaturesOptions opts) => this.DeleteAllExcelVbaSignatures(opts),
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

        private int SignAllVbaExcelFiles(SignAllVbaOptions options)
        {
            var exitCode = 0;

            var certName = options.CertName;
            var filePath = options.FilePath;

            if (certName != null && filePath != null)
            {
                this.logger.Debug($"file path= {filePath} certificate name= {certName}");

                try
                {
                    this.excelVbaService.SignAllVbaExcelFiles(filePath, certName);
                }
                catch (Exception ex)
                {
                    this.logger.Debug($"Execution stopped with Exception: {ex}");
                    Console.WriteLine($"Execution stopped with Exception!");
                    exitCode = -1;
                    return exitCode;
                }
            }
            else
            {
                exitCode = -1;
                return exitCode;
            }

            return exitCode;
        }

        private int SignOneVbaExcelFile(SignOneVbaExcelFileOptions options)
        {
            var exitCode = 0;

            var fileName = options.FileName;
            var certName = options.CertName;

            if (certName != null && fileName != null)
            {
                this.logger.Debug($"file name= {fileName} certificate name= {certName}");

                try
                {
                    this.excelVbaService.SignOneVbaExcelFileWithDigitalSignature(fileName, certName);
                }
                catch (Exception ex)
                {
                    this.logger.Debug($"Execution stopped with Exception: {ex}");
                    Console.WriteLine($"Execution stopped with Exception!");
                    exitCode = -1;
                    return exitCode;
                }
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

        private int DeleteDigitalSignatureFromOneVbaExcelFile(DeleteSignatureFromOneVbaExcelFileOptions options)
        {
            var exitCode = 0;

            var fileName = options.FileName;
            if (fileName != null)
            {
                this.logger.Debug($"Delete file signature from= {fileName}");
                this.excelVbaService.DeleteDigitalSignatureFromOneVbaExcelFile(fileName);
            }
            else
            {
                exitCode = -1;
                return exitCode;
            }

            return exitCode;
        }

        private int DeleteAllExcelVbaSignatures(DeleteAllExcelVbaSignaturesOptions options)
        {
            var exitCode = 0;

            var filePath = options.FilePath;
            if (filePath != null)
            {
                this.logger.Debug($"file path name= {filePath}");

                try
                {
                    this.excelVbaService.DeleteAllExcelVbaSignatures(filePath);
                }
                catch (Exception ex)
                {
                    this.logger.Debug($"Execution stopped with Exception: {ex}");
                    Console.WriteLine($"Execution stopped with Exception!");
                    exitCode = -1;
                    return exitCode;
                }
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

            this.logger.Debug("Number of errors: {0}", errs.Count());
            Console.WriteLine("Number of errors: {0}", errs.Count());

            foreach(var error in errs)
            {
                this.logger.Debug("Parser error: {0}", error.Tag.ToString());
                Console.WriteLine("Parser error: {0}", error.Tag.ToString());
            }

            return result;
        }

        private string[] SplitInputStringIntoArgumentsArray(string input)
        {
            string optionsAsDelimiterPattern = @"(\-[a-z])";
            string removeWhiteSpaceAtStartOrEndPattern = @"^\s+|\s+$";
            string empty = string.Empty;

            string[] substrings = Regex.Split(input, optionsAsDelimiterPattern);
            string[] arguments = new string[substrings.Length];

            for(int i = 0; i < substrings.Length; i++)
            {
                arguments[i] = Regex.Replace(substrings[i], removeWhiteSpaceAtStartOrEndPattern, empty);
            }

            return arguments;
        }
    }
}
