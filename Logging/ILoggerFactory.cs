namespace Logging
{
    public interface ILoggerFactory
    {
        ILogger Create<T>();
    }
}
