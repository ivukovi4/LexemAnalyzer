namespace LexemAnalyzer;

public record InternalCommand(string Name, Func<AnalyzeContext, Task> Action);