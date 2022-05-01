using CommandLine;
using ExcelServices;
using Logging;
using System;
using System.Collections.Generic;
using System.IO;
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

        public ExitCode ParseInput(string input)
        {
            var actualCommand = this.SplitInputStringIntoArgumentsArray(input);

            if(actualCommand.NoFileName)
            {
                return this.ParseInputArguments(actualCommand.Arguments);
            }

            if (actualCommand.FileName.Valid)
            {
                return this.ParseInputArguments(actualCommand.Arguments);
            }
            else
            {
                this.PrintInvalidFileName(actualCommand);
                return ExitCode.Error;
            }
        }

        private ExitCode ParseInputArguments(string[] args)
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

        private ExitCode SetPathToFiles(PathOptions options)
        {
            var path = options.PathToFiles;
            if (path != null)
            {
                this.logger.Debug($"path= {path}");
                this.excelService.SetPathToFiles(options.PathToFiles);
            }
            else
            {
                return ExitCode.Error;
            }

            return ExitCode.OK;
        }

        private ExitCode SetCertificateName(CertificateNameOptions options)
        {
            var certName = options.CertName;
            if (certName != null)
            {
                this.logger.Debug($"certificate name= {certName}");
                this.excelService.AddDigitalSignature(certName);
            }
            else
            {
                return ExitCode.Error;
            }

            return ExitCode.OK;
        }

        private ExitCode SignAllVbaExcelFiles(SignAllVbaOptions options)
        {
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
                    return ExitCode.Error;
                }
            }
            else
            {
                return ExitCode.Error;
            }

            return ExitCode.OK;
        }

        private ExitCode SignOneVbaExcelFile(SignOneVbaExcelFileOptions options)
        {
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
                    return ExitCode.Error;
                }
            }
            else
            {
                return ExitCode.Error;
            }

            return ExitCode.OK;
        }

        private ExitCode StopApp(StopOptions options)
        {
            if(options.Stop == ExitCode.Stop)
            {
                this.logger.Debug($"stop= {options.Stop}");
                return ExitCode.Stop;
            }

            return ExitCode.OK;
        }

        private ExitCode DeleteAllDigitalSignatures(DeleteSignatureOptions options)
        {
            var path = options.PathToFiles;
            if (path != null)
            {
                this.logger.Debug($"Delete file signature, path= {path}");
                this.excelService.DeleteAllDigitalSignatures(options.PathToFiles);
            }
            else
            {
                return ExitCode.Error;
            }

            return ExitCode.OK;
        }

        private ExitCode DeleteDigitalSignatureFromOneVbaExcelFile(DeleteSignatureFromOneVbaExcelFileOptions options)
        {
            var fileName = options.FileName;
            if (fileName != null)
            {
                this.logger.Debug($"Delete file signature from= {fileName}");
                this.excelVbaService.DeleteDigitalSignatureFromOneVbaExcelFile(fileName);
            }
            else
            {
                return ExitCode.Error;
            }

            return ExitCode.OK;
        }

        private ExitCode DeleteAllExcelVbaSignatures(DeleteAllExcelVbaSignaturesOptions options)
        {
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
                    return ExitCode.Error;
                }
            }
            else
            {
                return ExitCode.Error;
            }

            return ExitCode.OK;
        }

        private ExitCode HandleParseError(IEnumerable<Error> errs)
        {
            this.logger.Debug("Number of errors: {0}", errs.Count());
            Console.WriteLine("Number of errors: {0}", errs.Count());

            foreach(var error in errs)
            {
                this.logger.Debug("Parser error: {0}", error.Tag.ToString());
                Console.WriteLine("Parser error: {0}", error.Tag.ToString());
            }

            return ExitCode.Error;
        }

        private ActualCommandDto SplitInputStringIntoArgumentsArray(string input)
        {
            string optionsAsDelimiterPattern = @"(\-[a-z])";
            string removeWhiteSpaceAtStartOrEndPattern = @"^\s+|\s+$";
            string searchForDiskDriveLetter = @"[A-Z]:\\";
            string empty = string.Empty;

            string[] substrings = Regex.Split(input, optionsAsDelimiterPattern);
            string[] arguments = new string[substrings.Length];

            var validFileName = new ValidFileNameDto();

            for(int i = 0; i < substrings.Length; i++)
            {
                arguments[i] = Regex.Replace(substrings[i], removeWhiteSpaceAtStartOrEndPattern, empty);
            }

            for(int i = 0; i< arguments.Length; i++)
            {
                if(Regex.IsMatch(arguments[i], searchForDiskDriveLetter))
                {
                    validFileName = this.IsFileNameValid(arguments[i]);
                }
            }

            return new ActualCommandDto
            {
                Arguments = arguments,
                FileName = validFileName,
                NoFileName = validFileName.FileName == null ? true : false,
            };
        }

        private ValidFileNameDto IsFileNameValid(string fileName)
        {
            Regex containsABadCharacter = new Regex("["
                  + Regex.Escape(new string(Path.GetInvalidPathChars())) + "]");

            if (containsABadCharacter.IsMatch(fileName))
            {
                return new ValidFileNameDto()
                {
                    FileName = fileName,
                    Valid = false,
                    IllegalString = containsABadCharacter.Match(fileName).ToString()
                };
            };

            return new ValidFileNameDto()
            {
                FileName = fileName,
                Valid = true,
                IllegalString = ""
            };
        }

        private void PrintInvalidFileName(ActualCommandDto actualCommand)
        {
            this.logger.Debug($"Foud illegal character: {actualCommand.FileName.IllegalString} in file name: {actualCommand.FileName.FileName}");
            Console.WriteLine($"Foud illegal character: {actualCommand.FileName.IllegalString} in file name: {actualCommand.FileName.FileName}");
        }
    }
}
