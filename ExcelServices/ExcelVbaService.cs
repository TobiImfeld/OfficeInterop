using Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelServices
{
    public class ExcelVbaService : IExcelVbaService
    {
        private readonly ILogger logger;
        private readonly ICertificateStoreService certificateStoreService;
        private readonly IFileService fileService;
        private List<string> fileList;

        public ExcelVbaService(ILoggerFactory loggerFactory, ICertificateStoreService certificateStoreService, IFileService fileService)
        {
            this.logger = loggerFactory.Create<ExcelVbaService>();
            this.certificateStoreService = certificateStoreService;
            this.fileService = fileService;
        }

        public void SetPathToVbaFiles(string targetDirectory)
        {
            this.fileList = this.ListAllXlsmExcelFilesFromDirectory(targetDirectory);
        }

        public void AddDigitalSignatureToVbaMacro(string certName)
        {
            var cert = this.certificateStoreService.GetCertificateFromStore(certName);

            if (cert == null)
            {
                this.logger.Error($"{certName} not found!");
                this.ClearFileList();
            }
            else
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                foreach (var file in this.fileList)
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(file)))
                    {
                        using (ExcelWorkbook workbook = excelPackage.Workbook)
                        {
                            OfficeOpenXml.VBA.ExcelVbaProject vbaProject = null;

                            var check = this.IsVbaProjectExisting(workbook);

                            try
                            {
                                vbaProject = workbook.VbaProject;
                            }
                            catch (Exception ex)
                            {
                                workbook.Dispose();
                                excelPackage.Dispose();

                                throw new Exception($"Error at file: {file} ! Execution stopped with Exception: {ex}");
                            }

                            if (vbaProject != null)
                            {
                                this.IsVbaProjectSigned(workbook);

                                var vbaProjName = workbook.VbaProject.Name;
                                this.logger.Debug($"vba project name: {vbaProjName} in excel file: {file}");
                                Console.WriteLine($"vba project name: {vbaProjName} in excel file: {file}");

                                workbook.VbaProject.Signature.Certificate = cert;
                                excelPackage.SaveAs(new FileInfo(file));

                                workbook.Dispose();
                                excelPackage.Dispose();
                            }
                            else
                            {
                                var vbaProjName = workbook.VbaProject.Name;
                                workbook.CreateVBAProject();
                                workbook.VbaProject.Signature.Certificate = cert;

                                this.logger.Debug($"vba project name: {vbaProjName} in excel file: {file}");
                                Console.WriteLine($"vba project name: {vbaProjName} in excel file: {file}");

                                excelPackage.SaveAs(new FileInfo(file));

                                workbook.Dispose();
                                excelPackage.Dispose();
                            }
                        }
                    }
                }
            }
        }

        public void SignOneExcelFileWithDigitalSignature(string fileName, string certName)
        {
            var cert = this.certificateStoreService.GetCertificateFromStore(certName);

            if (cert == null)
            {
                this.logger.Error($"{certName} not found!");
                this.ClearFileList();
            }
            else
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(fileName)))
                {
                    using (ExcelWorkbook workbook = excelPackage.Workbook)
                    {
                        var vbaProjectExisting = this.IsVbaProjectExisting(workbook);

                        if (vbaProjectExisting)
                        {
                            this.IsVbaProjectSigned(workbook);

                            var vbaProjName = workbook.VbaProject.Name;
                            this.logger.Debug($"vba project name: {vbaProjName} in excel file: {fileName}");
                            Console.WriteLine($"vba project name: {vbaProjName} in excel file: {fileName}");

                            workbook.VbaProject.Signature.Certificate = cert;
                            excelPackage.SaveAs(new FileInfo(fileName));

                            workbook.Dispose();
                            excelPackage.Dispose();
                        }
                        else
                        {
                            workbook.CreateVBAProject();
                            workbook.VbaProject.Signature.Certificate = cert;
                            var vbaProjName = workbook.VbaProject.Name;

                            this.logger.Debug($"new vba project created with name: {vbaProjName} in excel file: {fileName}");
                            Console.WriteLine($"new vba project created with name: {vbaProjName} in excel file: {fileName}");

                            excelPackage.SaveAs(new FileInfo(fileName));

                            workbook.Dispose();
                            excelPackage.Dispose();
                        }
                    }
                }
            }
        }

        private List<string> ListAllXlsmExcelFilesFromDirectory(string targetDirectory)
        {
            var fileList = new List<string>();

            var filesFromDirectory = this.fileService.ListAllXlsmExcelFilesFromDirectory(targetDirectory);

            foreach (var files in filesFromDirectory)
            {
                foreach (var file in files.FileList)
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

        private bool IsVbaProjectSigned(ExcelWorkbook workbook)
        {
            var isSigned = false;

            try
            {
                var cert = workbook.VbaProject.Signature.Certificate;
                var issuerName = cert.IssuerName;

                if (cert != null)
                {
                    isSigned = true;
                }

                this.logger.Debug($"vba project in excel file is signed: {isSigned} with issuer name: {issuerName}");
                Console.WriteLine($"vba project in excel file is signed: {isSigned} with issuer name: {issuerName}");

                return isSigned;
            }
            catch (Exception ex)
            {
                workbook.Dispose();

                throw new Exception($"Error! Execution stopped with Exception: {ex}");
            }
        }

        private bool IsVbaProjectExisting(ExcelWorkbook workbook)
        {
            var vbaProjectExisting = false;

            try
            {
                var codeModule = workbook.CodeModule;

                if (codeModule != null)
                {
                    vbaProjectExisting = true;
                }

                this.logger.Debug($"vba project in excel file existing: {vbaProjectExisting}");
                Console.WriteLine($"vba project in excel file existing: {vbaProjectExisting}");

                return vbaProjectExisting;
            }
            catch (Exception ex)
            {
                workbook.Dispose();
                throw new Exception($"Error in excel file! Execution stopped with Exception: {ex}");
            }
        }
    }
}
