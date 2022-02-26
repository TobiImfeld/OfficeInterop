namespace CommandLineParser
{
    public enum ExitCode
    {
        OK,
        Stop,
        Error = -1
    }

    public interface IParserService
    {
        ExitCode ParseInput(string input);
    }
}
