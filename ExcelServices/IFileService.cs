using System.Collections.Generic;

namespace ExcelServices
{
    public interface IFileService
    {
        List<string> ListAllFilesFromDirectoryByFileExtension(string filePath, string fileExtension);
    }

    public static class OfficeFileExtensions
    {
        public const string XLS = ".xls";
        public const string XLSX = ".xlsx";
        public const string XLSM = ".xlsm";
        public const string DOC = ".doc";
        public const string DOCX = ".docx";
    }
}
