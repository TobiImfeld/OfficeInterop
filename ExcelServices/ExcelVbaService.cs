using Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelServices
{
    public class ExcelVbaService: IExcelVbaService
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

                            try
                            {
                                vbaProject = workbook.VbaProject;
                            }
                            catch (Exception ex)
                            {
                                this.logger.Debug($"Error at file: {file} ! Execution stopped with Exception: {ex}");
                                workbook.Dispose();
                                excelPackage.Dispose();
                                return;
                            }

                            if(vbaProject != null)
                            {
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
    }
}
