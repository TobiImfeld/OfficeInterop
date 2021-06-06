namespace Logging
{
    public class LoggerFactory : ILoggerFactory
    {
        public ILogger Create<T>()
        {
            return new Logger(typeof(T));
        }
    }
}
