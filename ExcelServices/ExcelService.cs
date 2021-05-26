﻿using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace ExcelServices
{
    public class ExcelService : IExcelService
    {
        private ICertificateStoreService certificateStoreService;
        private Application excelApp = null;
        private Workbooks books = null;
        private Workbook book = null;

        public ExcelService(ICertificateStoreService certificateStoreService)
        {
            this.certificateStoreService = certificateStoreService;
        }

        public void AddDigitalSignature(string filePath, string certName)
        {
            var cert = this.certificateStoreService.GetCertificateFromStore(certName);

            if (cert == null)
            {
                Console.WriteLine($"AddDigitalSignature(): Certificate CN={certName}' not found.");
            }
            else
            {
                try
                {
                    this.excelApp = new Application();
                    this.books = this.excelApp.Workbooks;
                    book = this.books.Open(filePath);

                    this.excelApp.DisplayAlerts = false;
                    this.excelApp.Visible = false;

                    var signatureSet = book.Signatures;
                    Signature signature = signatureSet.AddNonVisibleSignature(cert);
                    signatureSet.ShowSignaturesPane = false;

                    var signed = signature.IsSigned;
                    Console.WriteLine($"Is {filePath} signed: {signed}");

                    this.book.Close();
                    this.books.Close();
                    this.excelApp.Quit();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    this.disposeComObjects();
                    Console.ReadLine();
                }
                finally
                {
                    this.disposeComObjects();
                }
            }
        }

        private void disposeComObjects()
        {
            if (book != null) Marshal.ReleaseComObject(book);
            if (books != null) Marshal.ReleaseComObject(books);
            if (excelApp != null) Marshal.ReleaseComObject(excelApp);
        }
    }
}
