using System.Collections.Generic;

namespace ExcelServices
{
    public interface IFileService
    {
        List<FileListDto> ListAllExcelFilesFromDirectory(string filePath);
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
}
