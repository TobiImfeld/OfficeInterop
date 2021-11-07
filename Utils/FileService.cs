using Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Common
{
    public class FileService : IFileService
    {
        private readonly ILogger logger;
        private List<string> directoryFileList = new List<string>();

        public FileService(ILoggerFactory loggerFactory)
        {
            this.logger = loggerFactory.Create<FileService>();
        }

        public List<string> ListAllFilesFromDirectoryByFileExtension(string filePath, string fileExtension)
        {
            var fileList = new List<string>();

            this.ProcessDirectory(filePath, fileExtension);

            foreach (var file in this.directoryFileList)
            {
                fileList.Add(file);
            }

            this.directoryFileList.Clear();

            return fileList;
        }

        private void ProcessDirectory(string targetDirectory, string fileExtension)
        {
            var count = this.CountFilesInDirectoryByFileExtension(targetDirectory, fileExtension);
            this.PrintNumberOfFilesFromDirectory(count, targetDirectory);

            var fileEntries = Directory
                .EnumerateFiles(targetDirectory)
                .Where(filename =>
                    fileExtension.Equals(Path.GetExtension(filename))).ToList();

            foreach (string filePath in fileEntries)
            {
                PrintFileNamesFromDirectory(filePath, targetDirectory);
                this.directoryFileList.Add(filePath);
            }

            var subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
            {
                ProcessDirectory(subdirectory, fileExtension);
            }
        }

        private int CountFilesInDirectoryByFileExtension(string targetDirectory, string fileExtension)
        {
            return Directory
                .EnumerateFiles(targetDirectory)
                .Count(filename =>
                    fileExtension.Equals(Path.GetExtension(filename)));
        }

        private void PrintFileNamesFromDirectory(string filePath, string targetDirectory)
        {
            this.logger.Debug($"Found {Path.GetFileName(filePath)} file in {targetDirectory}");
            Console.WriteLine($"Found {Path.GetFileName(filePath)} file in {targetDirectory}");
        }

        private void PrintNumberOfFilesFromDirectory(int count, string targetDirectory)
        {
            this.logger.Debug($"Found {count} files in {targetDirectory}");
            Console.WriteLine($"Found {count} files in {targetDirectory}");
        }
    }
}
