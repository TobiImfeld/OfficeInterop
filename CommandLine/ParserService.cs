﻿using CommandLine;
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

        public int ParseInput(string[] args)
        {
           return this.parser.ParseArguments<
               PathOptions,
               CertificateNameOptions,
               DeleteSignatureOptions,
               StopOptions,
               VbaPathOptions,
               SignVbaOptions,
               SignOneExcelFileOptions,
               DeleteSignatureFromFileOptions>(args)
                .MapResult(
                (PathOptions opts) => this.SetPathToFiles(opts),
                (CertificateNameOptions opts) => this.SetCertificateName(opts),
                (DeleteSignatureOptions opts) => this.DeleteAllDigitalSignatures(opts),
                (StopOptions opts) => this.StopApp(opts),
                (VbaPathOptions opts) => this.SetPathToVbaFiles(opts),
                (SignVbaOptions opts) => this.SignVbaExcelFiles(opts),
                (SignOneExcelFileOptions opts) => this.SignOneExcelFile(opts),
                (DeleteSignatureFromFileOptions opts) => this.DeleteOneDigitalSignature(opts),
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

        private int SetPathToVbaFiles(VbaPathOptions options)
        {
            var exitCode = 0;

            var path = options.PathToVbaFiles;
            if (path != null)
            {
                this.logger.Debug($"vbapath= {path}");
                this.excelVbaService.SetPathToVbaFiles(options.PathToVbaFiles);
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

        private int SignVbaExcelFiles(SignVbaOptions options)
        {
            var exitCode = 0;

            var certName = options.CertName;
            if (certName != null)
            {
                this.logger.Debug($"certificate name= {certName}");

                try
                {
                    this.excelVbaService.AddDigitalSignatureToVbaMacro(certName);
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

        private int SignOneExcelFile(SignOneExcelFileOptions options)
        {
            var exitCode = 0;

            var filePath = options.FilePath;
            var certName = options.CertName;
            
            if (certName != null)
            {
                this.logger.Debug($"file path= {filePath}");
                this.logger.Debug($"certificate name= {certName}");

                try
                {
                    this.excelVbaService.SignOneExcelFileWithDigitalSignature(filePath, certName);
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

        private int DeleteOneDigitalSignature(DeleteSignatureFromFileOptions options)
        {
            var exitCode = 0;

            var fileName = options.FileName;
            if (fileName != null)
            {
                this.logger.Debug($"Delete file signature from= {fileName}");
                this.excelVbaService.DeleteOneDigitalSignatureFromExcelFile(options.FileName);
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
