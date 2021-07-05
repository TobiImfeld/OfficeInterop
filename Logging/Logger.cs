using Serilog;
using System;

namespace Logging
{
    public class Logger : ILogger
    {
        private readonly string callerName;

        public Logger(Type callerType)
        {
            this.callerName = callerType.Name;
        }

        public void Debug(string message)
        {
            Log.Debug("[{0}] message: {1}", this.callerName, message);
        }

        public void Debug<T>(string message, T propertyValue)
        {
            var msg = string.Format(message, propertyValue);
            Log.Debug("[{0}] message: {1}", this.callerName, msg);
        }

        public void Info(string message)
        {
            Log.Information("[{0}] message: {1}", this.callerName, message);
        }

        public void Info<T>(string message, T propertyValue)
        {
            var msg = string.Format(message, propertyValue);
            Log.Information("[{0}] message: {1}", this.callerName, msg);
        }

        public void Error(Exception ex)
        {
            Log.Error("[{0}] message: {1}", this.callerName, ex);
        }

        public void Error(string message)
        {
            Log.Error("[{0}] message: {1}", this.callerName, message);
        }
    }
}
