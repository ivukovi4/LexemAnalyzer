using CommandLine;

namespace LexemAnalyzer;

public class Options
{
    [Option('f', "file", Required = false, HelpText = "Путь до файла на разбор")]
    public string? FilePath { get; set; }
}