using Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.X509Certificates;

namespace ExcelServices
{
    public class ExcelVbaService : IExcelVbaService
    {
        private readonly ILogger logger;
        private readonly ICertificateStoreService certificateStoreService;
        private readonly IFileService fileService;
        private List<string> fileList;

        public ExcelVbaService(
            ILoggerFactory loggerFactory,
            ICertificateStoreService certificateStoreService,
            IFileService fileService)
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
                foreach (var file in this.fileList)
                {
                    this.SignExcelFileWithDigitalSignature(file, cert);
                }
            }
        }

        public void SignOneExcelFileWithDigitalSignature(string fileName, string certName)
        {
            var cert = this.certificateStoreService.GetCertificateFromStore(certName);

            if (cert == null)
            {
                this.logger.Error($"{certName} not found!");
            }
            else
            {
                this.SignExcelFileWithDigitalSignature(fileName, cert);
            }
        }

        public void DeleteOneDigitalSignatureFromExcelFile(string fileName)
        {
            var certWithoutPrivateKey = this.certificateStoreService.GetCertificateWithoutPrivateKeyFromStore();

            if (certWithoutPrivateKey == null)
            {
                this.logger.Error($"None certificate without private key found in certificate store Root");
            }
            else
            {
                this.DeleteDigitalSignatureFromExcelFile(fileName, certWithoutPrivateKey);
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

        private void SignExcelFileWithDigitalSignature(string fileName, X509Certificate2 cert)
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
                        
                        workbook.VbaProject.Signature.Certificate = cert;
                        excelPackage.SaveAs(new FileInfo(fileName));

                        this.logger.Debug($"vba project name: {vbaProjName} in excel file: {fileName} signed");
                        Console.WriteLine($"vba project name: {vbaProjName} in excel file: {fileName} signed");

                        workbook.Dispose();
                        excelPackage.Dispose();
                    }
                    else
                    {
                        workbook.CreateVBAProject();
                        workbook.VbaProject.Signature.Certificate = cert;
                        var vbaProjName = workbook.VbaProject.Name;

                        excelPackage.SaveAs(new FileInfo(fileName));

                        this.logger.Debug($"new vba project created with name: {vbaProjName} in excel file: {fileName} and signed");
                        Console.WriteLine($"new vba project created with name: {vbaProjName} in excel file: {fileName} and signed");

                        workbook.Dispose();
                        excelPackage.Dispose();
                    }
                }
            }
        }

        private void DeleteDigitalSignatureFromExcelFile(string fileName, X509Certificate2 certWithoutPrivateKey)
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
                        workbook.VbaProject.Signature.Certificate = certWithoutPrivateKey;
                        excelPackage.SaveAs(new FileInfo(fileName));

                        this.logger.Debug($"Signature from vba project name: {vbaProjName} in excel file: {fileName} deleted");
                        Console.WriteLine($"Signature from vba project name: {vbaProjName} in excel file: {fileName} deleted");

                        workbook.Dispose();
                        excelPackage.Dispose();
                    }
                    else
                    {
                        this.logger.Debug($"No vba project existing in excel file: {fileName}");
                        Console.WriteLine($"No vba project existing in excel file: {fileName}");

                        workbook.Dispose();
                        excelPackage.Dispose();
                    }
                }
            }
        }

        private bool IsVbaProjectSigned(ExcelWorkbook workbook)
        {
            var isSigned = false;

            try
            {
                var cert = workbook.VbaProject.Signature.Certificate;
                
                if (cert != null)
                {
                    isSigned = true;
                    var issuerName = cert.IssuerName.Name;

                    this.logger.Debug($"vba project in excel file is signed: {isSigned} with issuer name: {issuerName}");
                    Console.WriteLine($"vba project in excel file is signed: {isSigned} with issuer name: {issuerName}");
                }
                else
                {
                    this.logger.Debug($"vba project in excel file is not signed, certficate= null");
                }

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
                    if (workbook.VbaProject != null)
                    {
                        vbaProjectExisting = true;

                        this.logger.Debug($"vba project in excel file existing: {vbaProjectExisting}");
                        Console.WriteLine($"vba project in excel file existing: {vbaProjectExisting}");
                    }
                }

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
