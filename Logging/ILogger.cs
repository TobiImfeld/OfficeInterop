using System;

namespace Logging
{
    public interface ILogger
    {
        void Debug(string message);
        void Debug<T>(string message, T propertyValue);
        void Info(string message);
        void Info<T>(string message, T propertyValue);
        void Error(Exception ex);
        void Error(string message);
    }
}
