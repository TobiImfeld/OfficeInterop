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
            this.fileList = this.ListAllExcelFilesFromDirectory(targetDirectory);
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
                            OfficeOpenXml.VBA.ExcelVbaProject vbaProject  = null;

                            try
                            {
                                vbaProject = workbook.VbaProject;
                            }
                            catch (Exception ex)
                            {
                                this.logger.Debug($"no vba project added at file: {file}");
                                this.logger.Debug($"Exception: {ex}");
                            }
                           
                            if(vbaProject != null)
                            {
                                var vbaProjName = workbook.VbaProject.Name;
                                this.logger.Debug($"vba project name: {vbaProjName}");
                                Console.WriteLine($"vba project name: {vbaProjName}");

                                workbook.VbaProject.Signature.Certificate = cert;
                                excelPackage.SaveAs(new FileInfo(file));
                            }
                            else
                            {
                                workbook.CreateVBAProject();
                                workbook.VbaProject.Signature.Certificate = cert;
                                excelPackage.SaveAs(new FileInfo(file));
                            }

                            
                        }
                    }
                }
            }
        }

        private List<string> ListAllExcelFilesFromDirectory(string targetDirectory)
        {
            var fileList = new List<string>();

            var filesFromDirectory = this.fileService.ListAllExcelFilesFromDirectory(targetDirectory);

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
