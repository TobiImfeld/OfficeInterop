using Logging;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using System.IO;
using Common;

namespace WordServices
{
    public class WordService : IWordService
    {
        private readonly ILogger logger;
        private readonly ICertificateStoreService certificateStoreService;
        private readonly IFileService fileService;
        private readonly IWordVbaSignatureService wordVbaSignatureService;

        public WordService(
            ILoggerFactory loggerFactory,
            ICertificateStoreService certificateStoreService,
            IFileService fileService,
            IWordVbaSignatureService wordVbaSignatureService)
        {
            this.logger = loggerFactory.Create<WordService>();
            this.certificateStoreService = certificateStoreService;
            this.fileService = fileService;
            this.wordVbaSignatureService = wordVbaSignatureService;
        }

        public void SignAllWordFiles(string targetDirectory, string certName)
        {
            X509Certificate2 certificate = null;
            var fileList = this.ListAllWordFilesFromDirectory(targetDirectory);
            var cert = this.certificateStoreService.GetCertificateFromStore(certName);


            foreach (var file in fileList)
            {
                certificate = this.wordVbaSignatureService.GetSignatureFromZipPackage(file);
            }

            fileList.Clear();
        }

        private List<string> ListAllWordFilesFromDirectory(string targetDirectory)
        {
            return this.fileService.
                ListAllFilesFromDirectoryByFileExtension(
                targetDirectory,
                OfficeFileExtensions.DOCM
                );
        }

        private void AddDigitalSignature(string file, X509Certificate2 cert)
        {
            Application wordApp = null;
            Document document = null;

            try
            {
                wordApp = new Application();
                document = wordApp.Documents.Open(file);

                wordApp.Activate();

                //wordApp.DisplayAlerts = WdAlertLevel.wdAlertsMessageBox;
                //wordApp.Visible = false;

                var signatureSet = document.Signatures;
                Signature signature = signatureSet.AddNonVisibleSignature(cert);

                document.Save();

                if (signature != null)
                {
                    signatureSet.ShowSignaturesPane = false;
                    var signed = signature.IsSigned;

                    this.logger.Debug($"Is word file {document.Name} signed: {signed}");
                    Console.WriteLine($"Is word file {document.Name} signed: {signed}");
                }
                else
                {
                    Console.WriteLine($"Error: Could not set signature on file: {Path.GetFileName(file)}");
                    this.logger.Error($"Could not set signature on file: {Path.GetFileName(file)}");
                }

                document.Close();
                wordApp.Quit();
                this.DisposeComObjects(wordApp, document);
            }
            catch (Exception ex)
            {
                this.logger.Error(ex);
                this.DisposeComObjects(wordApp, document);
            }
            finally
            {
                this.DisposeComObjects(wordApp, document);
            }
        }

        private void DisposeComObjects(Application wordApp, Document document)
        {
            if (document != null) Marshal.ReleaseComObject(document);
            if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            this.logger.Debug($"Dispose all com-objects");
        }

    }
}
