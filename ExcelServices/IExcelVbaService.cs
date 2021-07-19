using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelServices
{
    public interface IExcelVbaService
    {
        void SetPathToVbaFiles(string targetDirectory);
        void AddDigitalSignatureToVbaMacro(string certName);
    }
}
