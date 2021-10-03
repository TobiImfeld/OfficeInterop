using Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelServices
{
    public class WordService : IWordService
    {
        private readonly ILogger logger;
        private readonly ICertificateStoreService certificateStoreService;
        private readonly IFileService fileService;

        public WordService(
            ILoggerFactory loggerFactory,
            ICertificateStoreService certificateStoreService,
            IFileService fileService)
        {
            this.logger = loggerFactory.Create<ExcelVbaService>();
            this.certificateStoreService = certificateStoreService;
            this.fileService = fileService;
        }


    }
}
