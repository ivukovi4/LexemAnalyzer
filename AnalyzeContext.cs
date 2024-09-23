namespace LexemAnalyzer;

public record AnalyzeContext(IList<Entry> Entries, IReadOnlyList<Poem> Poems, IReadOnlyList<Category> Categories);