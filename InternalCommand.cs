using Spectre.Console.Rendering;

namespace LexemAnalyzer;

public record InternalCommand(string Name, Func<AnalyzeContext, Task> Action);