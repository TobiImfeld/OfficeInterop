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

        public ExcelVbaService(
            ILoggerFactory loggerFactory,
            ICertificateStoreService certificateStoreService,
            IFileService fileService)
        {
            this.logger = loggerFactory.Create<ExcelVbaService>();
            this.certificateStoreService = certificateStoreService;
            this.fileService = fileService;
        }

        public void SignAllVbaExcelFiles(string targetDirectory, string certName)
        {
            var fileList = this.ListAllXlsmExcelFilesFromDirectory(targetDirectory);
            var cert = this.certificateStoreService.GetCertificateFromStore(certName);

            if (cert == null)
            {
                fileList.Clear();
            }
            else
            {
                foreach (var file in fileList)
                {
                    this.SignVbaExcelFileWithDigitalSignature(file, cert);
                }

                fileList.Clear();
            }
        }

        public void DeleteAllExcelVbaSignatures(string targetDirectory)
        {
            var fileList = this.ListAllXlsmExcelFilesFromDirectory(targetDirectory);
            var certWithoutPrivateKey = this.certificateStoreService.GetCertificateWithoutPrivateKeyFromStore();

            if (certWithoutPrivateKey == null)
            {
                this.logger.Error($"None certificate without private key found in certificate store Root");
            }
            else
            {
                foreach (var file in fileList)
                {
                    this.DeleteDigitalSignatureFromVbaExcelFile(file, certWithoutPrivateKey);
                }

                fileList.Clear();
            }
        }

        public void SignOneVbaExcelFileWithDigitalSignature(string fileName, string certName)
        {
            var cert = this.certificateStoreService.GetCertificateFromStore(certName);

            if (cert != null)
            {
                this.SignVbaExcelFileWithDigitalSignature(fileName, cert);
            }
        }

        public void DeleteDigitalSignatureFromOneVbaExcelFile(string fileName)
        {
            var certWithoutPrivateKey = this.certificateStoreService.GetCertificateWithoutPrivateKeyFromStore();

            if (certWithoutPrivateKey == null)
            {
                this.logger.Error($"None certificate without private key found in certificate store Root");
            }
            else
            {
                this.DeleteDigitalSignatureFromVbaExcelFile(fileName, certWithoutPrivateKey);
            }
        }

        private List<string> ListAllXlsmExcelFilesFromDirectory(string targetDirectory)
        {
            return this.fileService.ListAllXlsmExcelFilesFromDirectory(targetDirectory);
        }

        private void SignVbaExcelFileWithDigitalSignature(string fileName, X509Certificate2 cert)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(fileName)))
            {
                using (ExcelWorkbook workbook = excelPackage.Workbook)
                {
                    string vbaProjName = string.Empty;

                    var vbaProjectExisting = this.IsVbaProjectExisting(workbook, fileName);

                    switch (vbaProjectExisting)
                    {
                        case VbaProjectStates.Existing:

                            var vbaSigned = this.IsVbaProjectSigned(workbook);

                            if (!vbaSigned)
                            {
                                vbaProjName = workbook.VbaProject.Name;
                                workbook.VbaProject.Signature.Certificate = cert;
                                excelPackage.SaveAs(new FileInfo(fileName));

                                this.logger.Debug($"vba project name: {vbaProjName} in excel file: {fileName} signed");
                                Console.WriteLine($"vba project name: {vbaProjName} in excel file: {fileName} signed");
                            }
                            else
                            {
                                Console.WriteLine($"vba project in file: {fileName}  already signed");
                            }

                            workbook.Dispose();
                            excelPackage.Dispose();

                            break;

                        case VbaProjectStates.Inexisting:

                            workbook.CreateVBAProject();
                            workbook.VbaProject.Signature.Certificate = cert;
                            vbaProjName = workbook.VbaProject.Name;

                            excelPackage.SaveAs(new FileInfo(fileName));

                            this.logger.Debug($"new vba project created with name: {vbaProjName} in excel file: {fileName} and signed");
                            Console.WriteLine($"new vba project created with name: {vbaProjName} in excel file: {fileName} and signed");

                            workbook.Dispose();
                            excelPackage.Dispose();

                            break;

                        case VbaProjectStates.Error:

                            Console.WriteLine($"Error in excel file {fileName}, see log-file for more error details!");

                            workbook.Dispose();
                            excelPackage.Dispose();

                            break;
                    }
                }
            }
        }

        private void DeleteDigitalSignatureFromVbaExcelFile(string fileName, X509Certificate2 certWithoutPrivateKey)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(fileName)))
            {
                using (ExcelWorkbook workbook = excelPackage.Workbook)
                {
                    var vbaProjectExisting = this.IsVbaProjectExisting(workbook, fileName);

                    switch (vbaProjectExisting)
                    {
                        case VbaProjectStates.Existing:

                            var vbaSigned = this.IsVbaProjectSigned(workbook);

                            if (vbaSigned)
                            {
                                var vbaProjName = workbook.VbaProject.Name;
                                workbook.VbaProject.Signature.Certificate = certWithoutPrivateKey;
                                excelPackage.SaveAs(new FileInfo(fileName));

                                this.logger.Debug($"Signature from vba project name: {vbaProjName} in excel file: {fileName} deleted");
                                Console.WriteLine($"Signature from vba project name: {vbaProjName} in excel file: {fileName} deleted");
                            }
                            else
                            {
                                Console.WriteLine($"vba project signature in file: {fileName} already deleted");
                            }

                            workbook.Dispose();
                            excelPackage.Dispose();

                            break;

                        case VbaProjectStates.Inexisting:

                            this.logger.Debug($"No vba project existing in excel file: {fileName}");
                            Console.WriteLine($"No vba project existing in excel file: {fileName}");

                            workbook.Dispose();
                            excelPackage.Dispose();

                            break;

                        case VbaProjectStates.Error:

                            Console.WriteLine($"Error in excel file {fileName}, see log-file for more error details!");

                            workbook.Dispose();
                            excelPackage.Dispose();

                            break;
                    }
                }
            }
        }

        private bool IsVbaProjectSigned(ExcelWorkbook workbook)
        {
            var isSigned = false;
            var cert = workbook.VbaProject.Signature.Certificate;

            if (cert != null)
            {
                isSigned = true;
                var issuerName = cert.IssuerName.Name;

                this.logger.Debug($"vba project in excel file is signed: {isSigned} with issuer name: {issuerName}");
            }
            else
            {
                this.logger.Debug($"vba project in excel file is not signed, certficate= null");
            }

            return isSigned;
        }

        private VbaProjectStates IsVbaProjectExisting(ExcelWorkbook workbook, string fileName)
        {
            try
            {
                var codeModule = workbook.CodeModule;

                if (codeModule != null)
                {
                    if (workbook.VbaProject != null)
                    {
                        this.logger.Debug($"vba project in excel file {fileName}: {VbaProjectStates.Existing.ToString()}");
                        return VbaProjectStates.Existing;
                    }
                }

                this.logger.Debug($"vba project in excel file {fileName}: {VbaProjectStates.Inexisting.ToString()}");
                return VbaProjectStates.Inexisting;
            }
            catch (Exception ex)
            {
                this.logger.Error($"Error in excel file: {fileName} Exception: {ex}");
                return VbaProjectStates.Error;
            }
        }
    }
}
