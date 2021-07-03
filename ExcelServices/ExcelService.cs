using Logging;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace ExcelServices
{
    public class ExcelService : IExcelService
    {
        private readonly ILogger logger;
        private readonly ICertificateStoreService certificateStoreService;
        private Application excelApp = null;
        private Workbooks books = null;
        private Workbook book = null;
        private List<string> fileList;
        private HashSet<string> fileExtensions = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase){ ".xls", ".xlsx", ".xlsm" };

        public ExcelService(ILoggerFactory loggerFactory, ICertificateStoreService certificateStoreService)
        {
            this.logger = loggerFactory.Create<ExcelService>();
            this.certificateStoreService = certificateStoreService;
        }

        public void SetPathToFiles(string filePath)
        {
            this.CountFiles(filePath);
            this.fileList = this.ListAllExcelFilesFrom(filePath);
        }

        public void AddDigitalSignature(string certName)
        {
            var cert = this.certificateStoreService.GetCertificateFromStore(certName);

            if (cert == null)
            {
                this.logger.Error($"{certName} not found!");
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

                            this.excelApp.DisplayAlerts = false;
                            this.excelApp.Visible = false;

                            var signatureSet = this.book.Signatures;

                            Signature signature = signatureSet.AddNonVisibleSignature(cert);
                            if (signature != null)
                            {
                                signatureSet.ShowSignaturesPane = false;
                                var signed = signature.IsSigned;

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
                    this.DisposeComObjects();
                }
                finally
                {
                    this.DisposeComObjects();
                }
            }
        }

        public void DeleteAllDigitalSignatures(string filePath)
        {
            var fileList = this.ListAllExcelFilesFrom(filePath);

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

        private int CountFiles(string filePath)
        {
            var count = Directory
                .EnumerateFiles(filePath)
                .Count(filename =>
                    fileExtensions.Contains(Path.GetExtension(filename)));

            this.logger.Debug($"Found {count} files in {filePath}");
            Console.WriteLine($"Found {count} files in {filePath}");

            return count;
        }

        private List<string> ListAllExcelFilesFrom(string filePath)
        {
            var filePaths = Directory.GetFiles(filePath);

            var fileList = Directory
                .EnumerateFiles(filePath)
                .Where(filename =>
                    fileExtensions.Contains(Path.GetExtension(filename))).ToList();

            Console.WriteLine($"File list: ");
            foreach (var file in fileList)
            {
                Console.WriteLine(Path.GetFileName(file));
            }

            return fileList;
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