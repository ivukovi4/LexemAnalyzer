// See https://aka.ms/new-console-template for more information

using System.Text;
using CommandLine;
using ExcelDataReader;
using LexemAnalyzer;
using Spectre.Console;
using Table = Spectre.Console.Table;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

await Parser.Default.ParseArguments<Options>(args).WithParsedAsync(async o =>
{
    if (string.IsNullOrEmpty(o.FilePath))
        o.FilePath = Directory.GetFiles("./", "*.xlsx", SearchOption.AllDirectories).FirstOrDefault();

    if (!File.Exists(o.FilePath))
        throw new FileNotFoundException(o.FilePath);

    await using var fileStream = File.Open(o.FilePath, FileMode.Open, FileAccess.Read);

    using var excelDataReader = ExcelReaderFactory.CreateReader(fileStream);

    Entry[] entries =
    [
        ..GetEntries(excelDataReader)
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

    AnsiConsole.WriteLine();

    AnsiConsole.WriteLine("Нажмите любую кнопку чтобы выйти (Нажмите 'C' для копирования таблицы) ...");

    var key = Console.ReadKey(true);
    if (key.KeyChar is 'c' or 'C' or 'с' or 'С')
    {
        await TextCopy.ClipboardService.SetTextAsync(string.Join(Environment.NewLine,
            categories.Select(x => x.ToString())));

        AnsiConsole.WriteLine("Таблица скопирована в буфер обмена ...");
    }

    AnsiConsole.WriteLine("Выходим ...");
});

return;

IEnumerable<Entry> GetEntries(IExcelDataReader reader)
{
    while (reader.Read())
    {
        var name = reader.GetString(0);
        var count = (int)reader.GetDouble(1);

        string?[] categories =
        [
            reader.GetString(2),
            reader.GetString(3),
            reader.GetString(4)
        ];

        yield return new Entry(name, count, [
            ..categories
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Select(x => x!.ToLower().Trim())
        ]);
    }
}