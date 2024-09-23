// See https://aka.ms/new-console-template for more information

using System.Text;
using LexemAnalyzer;
using Spectre.Console;
using Spectre.Console.Cli;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

var app = new CommandApp<AnalyzeCommand>();
app.Configure(config => { config.SetExceptionHandler((exception, _) => { AnsiConsole.WriteException(exception); }); });
return await app.RunAsync(args);