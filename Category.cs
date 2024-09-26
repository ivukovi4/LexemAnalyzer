namespace LexemAnalyzer;

public record Category(string Name, int Count, int TotalCount, Entry[] Entries)
{
    public string EntriesString => string.Join(", ", Entries
        .OrderByDescending(x => x.Count)
        .Select(x => x.Count > 1 ? $"{x.Name}[gray]({x.Count})[/]" : x.Name));

    public override string ToString()
    {
        string[] fields = [Name, Count.ToString(), TotalCount.ToString(), EntriesString];

        return string.Join('\t', fields);
    }
}