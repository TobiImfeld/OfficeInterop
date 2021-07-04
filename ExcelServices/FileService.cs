using Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelServices
{
    public class FileService : IFileService
    {
        private readonly ILogger logger;
        private HashSet<string> fileExtensions = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase) { ".xls", ".xlsx", ".xlsm" };
        private List<FileListDto> directoryFileList = new List<FileListDto>();

        public FileService(ILoggerFactory loggerFactory)
        {
            this.logger = loggerFactory.Create<FileService>();
        }

        public List<FileListDto> ListAllExcelFilesFromDirectory(string filePath)
        {
            this.ProcessDirectory(filePath);
            return this.RemoveEntriesWithZeroFiles(this.directoryFileList);
        }

        private void ProcessDirectory(string targetDirectory)
        {
            var count = this.CountFilesInDirectory(targetDirectory);
            this.PrintNumberOfFilesFromDirectory(count, targetDirectory);

            var fileEntries = Directory
                .EnumerateFiles(targetDirectory)
                .Where(filename =>
                    fileExtensions.Contains(Path.GetExtension(filename))).ToList();

            foreach (string filePath in fileEntries)
            {
                PrintFileNamesFromDirectory(filePath, targetDirectory);
            }

            var FileListDto = new FileListDto(count, fileEntries);
            this.directoryFileList.Add(FileListDto);
                
            var subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
            {
                ProcessDirectory(subdirectory);
            }
        }

        private int CountFilesInDirectory(string targetDirectory)
        {
            return Directory
                .EnumerateFiles(targetDirectory)
                .Count(filename =>
                    fileExtensions.Contains(Path.GetExtension(filename)));
        }

        private List<FileListDto> RemoveEntriesWithZeroFiles(List<FileListDto> fileList)
        {
            return fileList
                .Where(item => item.NumberOfFiles != 0)
                .Select(s => s).ToList();
        }

        private void PrintFileNamesFromDirectory(string filePath, string targetDirectory)
        {
            this.logger.Debug($"Found {Path.GetFileName(filePath)} files in {targetDirectory}");
            Console.WriteLine($"Found {Path.GetFileName(filePath)} files in {targetDirectory}");
        }

        private void PrintNumberOfFilesFromDirectory(int count, string targetDirectory)
        {
            this.logger.Debug($"Found {count} files in {targetDirectory}");
            Console.WriteLine($"Found {count} files in {targetDirectory}");
        }
    }
}
