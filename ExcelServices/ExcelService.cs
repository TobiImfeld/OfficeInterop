﻿using Logging;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelServices
{
    public class ExcelService : IExcelService
    {
        private readonly ILogger logger;
        private readonly ICertificateStoreService certificateStoreService;
        private readonly IFileService fileService;
        private Application excelApp = null;
        private Workbooks books = null;
        private Workbook book = null;
        private List<string> fileList;

        public ExcelService(ILoggerFactory loggerFactory, ICertificateStoreService certificateStoreService, IFileService fileService)
        {
            this.logger = loggerFactory.Create<ExcelService>();
            this.certificateStoreService = certificateStoreService;
            this.fileService = fileService;
        }

        public void SetPathToFiles(string targetDirectory)
        {
            this.fileList = this.ListAllExcelFilesFromDirectory(targetDirectory);
        }

        public void AddDigitalSignature(string certName)
        {
            var cert = this.certificateStoreService.GetCertificateFromStore(certName);

            if (cert == null)
            {
                this.logger.Error($"{certName} not found!");
                this.ClearFileList();
            }
            else
            {
                try
                {
                    foreach(var file in this.fileList)
                    {
                        try
                        {
                            this.excelApp = new Application();
                            this.books = this.excelApp.Workbooks;
                            this.book = this.books.Open(file);

                            var vbaSigned = this.book.VBASigned;
                            var vbaProjName = this.book.VBProject.Collection.VBE;

                            this.excelApp.DisplayAlerts = false;
                            this.excelApp.Visible = false;

                            var signatureSet = this.book.Signatures;

                            Signature signature = signatureSet.AddNonVisibleSignature(cert);
                            if (signature != null)
                            {
                                signatureSet.ShowSignaturesPane = false;
                                var signed = signature.IsSigned;

                                this.logger.Debug($"Is vba macro {vbaProjName} signed: {vbaSigned}");
                                Console.WriteLine($"Is vba macro {vbaProjName} signed: {vbaSigned}");

                                this.logger.Debug($"Is file {Path.GetFileName(file)} signed: {signed}");
                                this.logger.Debug($"Signature issuer: {signature.Issuer}");
                                Console.WriteLine($"Is file {Path.GetFileName(file)} signed: {signed}");
                            }
                            else
                            {
                                Console.WriteLine($"Error: Could not set signature on file: {Path.GetFileName(file)}");
                                this.logger.Error($"Could not set signature on file: {Path.GetFileName(file)}");
                            }

                            this.book.Close();
                            this.books.Close();
                            this.excelApp.Quit();
                            this.DisposeComObjects();
                        }
                        catch (Exception ex)
                        {
                            this.logger.Error(ex);
                            this.DisposeComObjects();
                        }
                        finally
                        {
                            this.DisposeComObjects();
                        }
                    }

                    Console.WriteLine($"Work done, all found files signed!");
                }
                catch (Exception ex)
                {
                    this.logger.Error(ex);
                    this.ClearFileList();
                    this.DisposeComObjects();
                }
                finally
                {
                    this.ClearFileList();
                    this.DisposeComObjects();
                }
            }
        }

        public void DeleteAllDigitalSignatures(string targetDirectory)
        {
            var fileList = this.ListAllExcelFilesFromDirectory(targetDirectory);

            try
            {
                foreach (var file in fileList)
                {
                    try
                    {
                        this.excelApp = new Application();
                        this.books = this.excelApp.Workbooks;
                        this.book = this.books.Open(file);

                        this.excelApp.DisplayAlerts = false;
                        this.excelApp.Visible = false;

                        var signatureSet = this.book.Signatures;
                        var enumerator = signatureSet.GetEnumerator();

                        while (enumerator.MoveNext())
                        {
                            Signature signature = enumerator.Current as Signature;
                            Console.WriteLine($"Delete Signature: {signature.Details.SignatureText} from file: {file}");
                            this.logger.Debug($"Delete Signature: {signature.Details.SignatureText} from file: {file}");
                            signature.Delete();
                        }

                        this.book.Close();
                        this.books.Close();
                        this.excelApp.Quit();
                        this.DisposeComObjects();
                    }
                    catch (Exception ex)
                    {
                        this.logger.Error(ex);
                        this.DisposeComObjects();
                    }
                    finally
                    {
                        this.DisposeComObjects();
                    }
                }

                Console.WriteLine($"Work done, all signatures deleted!");
            }
            catch (Exception ex)
            {
                this.logger.Error(ex);
                this.DisposeComObjects();
            }
            finally
            {
                this.DisposeComObjects();
            }
        }

        private List<string> ListAllExcelFilesFromDirectory(string targetDirectory)
        {
            var fileList = new List<string>();

            var filesFromDirectory = this.fileService.ListAllExcelFilesFromDirectory(targetDirectory);

            foreach (var files in filesFromDirectory)
            {
                foreach(var file in files.FileList)
                {
                    fileList.Add(file);
                }
            }

            return fileList;
        }

        private void ClearFileList()
        {
            this.fileList.Clear();
        }

        private void DisposeComObjects()
        {
            if (book != null) Marshal.ReleaseComObject(book);
            if (books != null) Marshal.ReleaseComObject(books);
            if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            this.logger.Debug($"Dispose all com-objects");
        }
    }
}