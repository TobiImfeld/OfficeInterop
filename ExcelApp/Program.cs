using ExcelServices;
using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelApp
{
    class Program
    {
        

        static void Main(string[] args)
        {
            ICertificateStoreService certificateStoreService = new CertificateStoreService();
            IExcelService excelService = new ExcelService(certificateStoreService);
            excelService.AddDigitalSignature(@"C:\Temp\Test1.xlsx","TobiOfficeCert");
            //Parser.Default.ParseArguments<PushCommand, CommitCommand>(args)
            //    .WithParsed<ICommand>(t => t.Execute());
        }
    }
}
