using System.ComponentModel;
using System.Diagnostics;
using System.Text.RegularExpressions;
using ExcelDataReader;
using Spectre.Console;
using Spectre.Console.Cli;
using TextCopy;

namespace LexemAnalyzer;

// ReSharper disable once ClassNeverInstantiated.Global
internal sealed partial class AnalyzeCommand : AsyncCommand<AnalyzeCommand.Options>
{
    public override async Task<int> ExecuteAsync(CommandContext commandContext, Options settings)
    {
        if (string.IsNullOrEmpty(settings.FilePath))
            settings.FilePath = Directory.GetFiles("./", "*.xlsx", SearchOption.AllDirectories).FirstOrDefault();

        if (!File.Exists(settings.FilePath))
            throw new FileNotFoundException(settings.FilePath);

        AnalyzeContext? analyzeContext = null;

        await AnsiConsole.Status()
            .StartAsync("Парсим файл...", async _ =>
            {
                await using var fileStream = File.Open(settings.FilePath, FileMode.Open, FileAccess.Read);

                using var excelDataReader = ExcelReaderFactory.CreateReader(fileStream);

                var keyRegex = KeyRegex();

                List<Entry> entries = [];
                List<Poem> poems = [];

                do
                {
                    if (keyRegex.IsMatch(excelDataReader.Name))
                        entries.AddRange([..GetEntries(excelDataReader)]);
                    else
                        poems.Add(GetPoem(excelDataReader));
                } while (excelDataReader.NextResult());

                analyzeContext = new AnalyzeContext(entries, poems,
                [
                    ..entries
                        .SelectMany(x => x.Categories)
                        .Distinct()
                        .Select(x => new Category(
                            x,
                            entries.Count(y => y.Categories.Contains(x)),
                            entries.Sum(y => y.Categories.Contains(x) ? y.Count : 0),
                            [..entries.Where(y => y.Categories.Contains(x))]))
                ]);
            });

        Debug.Assert(analyzeContext != null, nameof(analyzeContext) + " != null");

        await ShowEntriesAsync(analyzeContext);

        InternalCommand[] commands =
        [
            new("Скопировать результат", CopyEntriesAsync),
            // new("Показать результат", ShowEntriesAsync),
            new("Найти слово", SearchAsync),
            new("Выход", QuitAsync)
        ];

        while (true)
            try
            {
                AnsiConsole.WriteLine();

                var commandName = AnsiConsole.Prompt(
                    new SelectionPrompt<string>()
                        .Title("[bold blue]Что сделать дальше?[/]")
                        .AddChoices(commands.Select(x => x.Name)));

                await commands.First(x => x.Name == commandName).Action(analyzeContext);
            }
            catch (QuitException)
            {
                AnsiConsole.Write(new Markup("[grey]Выходим ...[/]"));
                AnsiConsole.WriteLine();

                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.WriteException(ex);
            }
    }

    private static Task QuitAsync(AnalyzeContext arg)
    {
        throw new QuitException();
    }

    private static Task SearchAsync(AnalyzeContext arg)
    {
        var searchQuery = AnsiConsole.Prompt(new TextPrompt<string>("Какое слово ищем?")).ToLower();

        var entry = arg.Entries.FirstOrDefault(x => x.Name == searchQuery);
        if (entry != null)
        {
            AnsiConsole.WriteLine();
            AnsiConsole.Write(new Markup("[bold yellow]Категории[/]"));
            AnsiConsole.WriteLine();

            foreach (var category in entry.Categories) AnsiConsole.WriteLine(category);
        }

        var found = false;

        foreach (var poem in arg.Poems)
        {
            if (!poem.Words.Any(x => x.Name.Contains(searchQuery))) continue;

            if (found == false)
            {
                AnsiConsole.WriteLine();
                AnsiConsole.Write(new Markup("[bold green]Результаты поиска[/]"));
                AnsiConsole.WriteLine();

                found = true;
            }

            AnsiConsole.WriteLine(poem.Name);
        }

        if (!found) AnsiConsole.WriteLine("Ничего не найдено");

        return Task.CompletedTask;
    }

    private static async Task CopyEntriesAsync(AnalyzeContext context)
    {
        await ClipboardService.SetTextAsync(string.Join(Environment.NewLine,
            context.Categories.Select(x => x.ToString())));

        AnsiConsole.WriteLine("Таблица скопирована в буфер обмена ...");
    }

    private static Task ShowEntriesAsync(AnalyzeContext context)
    {
        Category[] categories =
        [
            ..context.Entries
                .SelectMany(x => x.Categories)
                .Distinct()
                .Select(x => new Category(
                    x,
                    context.Entries.Count(y => y.Categories.Contains(x)),
                    context.Entries.Sum(y => y.Categories.Contains(x) ? y.Count : 0),
                    [..context.Entries.Where(y => y.Categories.Contains(x))]))
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
            categoriesTable.AddRow(
                category.Name,
                category.Count.ToString(),
                category.TotalCount.ToString(),
                category.EntriesString);

        AnsiConsole.Write(categoriesTable);

        return Task.CompletedTask;
    }

    [GeneratedRegex(".*частотный словарь.*")]
    private static partial Regex KeyRegex();

    private static Poem GetPoem(IExcelDataReader reader)
    {
        return new Poem(reader.Name, [..GetWords(reader)]);
    }

    private static IEnumerable<PoemWord> GetWords(IExcelDataReader reader)
    {
        while (reader.Read()) yield return new PoemWord(reader.GetString(0).ToLower(), (int)reader.GetDouble(1));
    }

    private static IEnumerable<Entry> GetEntries(IExcelDataReader reader)
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

    // ReSharper disable once ClassNeverInstantiated.Global
    public sealed class Options : CommandSettings
    {
        [Description("The file to analyze.")]
        [CommandArgument(0, "[filePath]")]
        public string? FilePath { get; set; }
    }
}