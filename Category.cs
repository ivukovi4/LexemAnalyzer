namespace LexemAnalyzer;

public record Category(string Name, int Count, int TotalCount, Entry[] Entries)
{
    public string GetEntriesString(bool withStyles) => string.Join(", ", Entries
        .OrderByDescending(x => x.Count)
        .Select(x => x.Count > 1 ? withStyles  ?$"{x.Name}[gray]({x.Count})[/]" : $"{x.Name}({x.Count})" : x.Name));

    public string ToString(bool withStyles)
    {
        string[] fields = [Name, Count.ToString(), TotalCount.ToString(), GetEntriesString(withStyles)];

        return string.Join('\t', fields);
    }

    public override string ToString()
    {
        string[] fields = [Name, Count.ToString(), TotalCount.ToString(), GetEntriesString(true)];

        return string.Join('\t', fields);
    }
}