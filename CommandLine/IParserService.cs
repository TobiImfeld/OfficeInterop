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

    public class ValidFileNameDto
    {
        public string FileName { get; set; }
        public string IllegalString { get; set; }
        public bool Valid { get; set; }
    }

    public class ActualCommandDto
    {
        public string[] Arguments { get; set; }
        public ValidFileNameDto FileName { get; set; }
        public bool NoFileName { get; set; }
    }
}
