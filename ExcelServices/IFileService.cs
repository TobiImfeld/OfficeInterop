using System.Collections.Generic;

namespace ExcelServices
{
    public interface IFileService
    {
        List<FileListDto> ListAllExcelFilesFromDirectory(string filePath);
        List<FileListDto> ListAllXlsmExcelFilesFromDirectory(string filePath);

    }

    public class FileListDto
    {
        public int NumberOfFiles { get; }
        public List<string> FileList { get; }

        public FileListDto(int numberOfFiles, List<string> fileList)
        {
            this.NumberOfFiles = numberOfFiles;
            this.FileList = fileList;
        }
    }

    public static class ExcelFileExtensions
    {
        public const string XLS = ".xls";
        public const string XLSX = ".xlsx";
        public const string XLSM = ".xlsm";
    }
}
