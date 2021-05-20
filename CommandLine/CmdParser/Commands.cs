using System;

namespace CommandLine
{
    [Verb("push", HelpText = "Save all your commits to the cloud")]
    public class PushCommand : ICommand
    {
        public void Execute()
        {
            Console.WriteLine("Executing Push");
        }
    }

    [Verb("commit", HelpText = "Save a code change")]
    public class CommitCommand : ICommand
    {
        [Option('m', "message", Required = true, HelpText = "Explain what code change you did")]
        public string Message { get; set; }
        public void Execute()
        {
            Console.WriteLine($"Executing Commit with message: {Message}");
        }
    }
}
