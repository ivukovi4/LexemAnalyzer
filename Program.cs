// See https://aka.ms/new-console-template for more information

using CommandLine;
using DevExpress.Spreadsheet;
using LexemAnalyzer;
using Spectre.Console;
using Table = Spectre.Console.Table;

await Parser.Default.ParseArguments<Options>(args).WithParsedAsync(async o =>
{
    using var workbook = new Workbook();

    if (string.IsNullOrEmpty(o.FilePath))
        o.FilePath = Directory.GetFiles("./", "*.xlsx", SearchOption.AllDirectories).FirstOrDefault();

    if (!File.Exists(o.FilePath))
        throw new FileNotFoundException(o.FilePath);

    if (!await workbook.LoadDocumentAsync(o.FilePath))
        throw new InvalidOperationException("Can't open file");

    Entry[] entries =
    [
        ..GetEntries(workbook.Worksheets[0])
    ];

    Category[] categories =
    [
        ..entries
            .SelectMany(x => x.Categories)
            .Distinct()
            .Select(x => new Category(
                x,
                entries.Count(y => y.Categories.Contains(x)),
                entries.Sum(y => y.Categories.Contains(x) ? y.Count : 0),
                [..entries.Where(y => y.Categories.Contains(x))]))
    ];

    var categoriesTable = new Table
    {
        ShowRowSeparators = true
    };
    categoriesTable.AddColumn("Категория");
    categoriesTable.AddColumn("Кол-во уникальных компонентов", column => { column.Alignment = Justify.Right; });
    categoriesTable.AddColumn("Общее кол-во компонентов", column => { column.Alignment = Justify.Right; });
    categoriesTable.AddColumn("Компоненты");

    foreach (var category in categories.OrderByDescending(x => x.TotalCount))
    {
        categoriesTable.AddRow(
            category.Name,
            category.Count.ToString(),
            category.TotalCount.ToString(),
            category.EntriesString);
    }

    AnsiConsole.Write(categoriesTable);

    await TextCopy.ClipboardService.SetTextAsync(string.Join(Environment.NewLine,
        categories.Select(x => x.ToString())));

    AnsiConsole.WriteLine();

    AnsiConsole.WriteLine("Таблица скопирована в буфер обмена ...");
});

return;

IEnumerable<Entry> GetEntries(Worksheet? worksheet)
{
    if (worksheet == null) yield break;

    for (var i = worksheet.Rows.LastUsedIndex; i >= 0; i--)
    {
        var name = worksheet.Rows[i][0].Value.ToString();

        var count = ParseCount(worksheet.Rows[i][1].Value);

        string?[] categories =
        [
            worksheet.Rows[i][2].Value.ToString(),
            worksheet.Rows[i][3].Value.ToString(),
            worksheet.Rows[i][3].Value.ToString()
        ];

        yield return new Entry(name, count, [
            ..categories
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Select(x => x!.ToLower().Trim())
        ]);
    }
}

int ParseCount(CellValue cell)
{
    if (cell.IsNumeric)
    {
        return (int)cell.NumericValue;
    }

    return -1;
}