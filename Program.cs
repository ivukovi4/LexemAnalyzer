// See https://aka.ms/new-console-template for more information

using System.Text;
using LexemAnalyzer;
using Spectre.Console.Cli;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

var app = new CommandApp<AnalyzeCommand>();
return await app.RunAsync(args);