using System;
using ExcelServices;
using CommandLine;

namespace ExcelApp
{
    class Program
    {
        

        static void Main(string[] args)
        {
            var logFile = ConfigurationManager.AppSettings["LogFilePath"];

            ICertificateStoreService certificateStoreService = new CertificateStoreService();
            IExcelService excelService = new ExcelService(certificateStoreService);
            excelService.AddDigitalSignature(@"C:\Temp\Test1.xlsx","TobiOfficeCert");



            //Parser.Default.ParseArguments<PushCommand, CommitCommand>(args)
            //    .WithParsed<ICommand>(t => t.Execute());
        }
    }
}
