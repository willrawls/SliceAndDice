using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Sandy.Core;

return await RunAsync(args);

static async Task<int> RunAsync(string[] arguments)
{
    try
    {
        if (arguments.Length == 0 || arguments[0] is "help" or "--help")
        {
            Console.WriteLine("""
                Sandy C# 10 — library inspection and side-effect-free template preview
                list <Category.xml> [search]
                preview <Category.xml> <template name> [Name=Value ...]
                import-csv <folder> <output.xml> [utf-8|windows-1252]
                roundtrip <Category.xml> <output.xml>
                audit <Category.xml>

                Preview returns planned output as JSON. It never writes generated files,
                changes the clipboard, starts processes, or sends keystrokes.
                """);
            return 0;
        }
        if (arguments[0] == "import-csv")
        {
            if (arguments.Length is < 3 or > 4) throw new ArgumentException("import-csv requires folder, output path, and optional encoding.");
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var encoding = Encoding.GetEncoding(arguments.Length == 4 ? arguments[3] : "utf-8", EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
            using var categories = new StreamReader(Path.Combine(arguments[1], "Category.csv"), encoding);
            using var templates = new StreamReader(Path.Combine(arguments[1], "Template.csv"), encoding);
            var imported = LibraryCsv.Read(categories, templates, Path.GetFileName(Path.GetFullPath(arguments[1])));
            SaveNew(imported, arguments[2]);
            Console.WriteLine($"Imported {imported.Categories.Count} categories, {imported.Templates.Count()} templates.");
            return 0;
        }
        if (arguments.Length < 2) throw new ArgumentException("A Category.xml path is required. Use --help.");
        var library = LibraryXml.Load(arguments[1]);
        switch (arguments[0])
        {
            case "list":
                foreach (var category in library.Categories)
                    foreach (var template in category.Templates.Where(t => arguments.Length < 3 || t.Name.Contains(arguments[2], StringComparison.OrdinalIgnoreCase)))
                        Console.WriteLine($"{template.Id}\t{category.Name}\t{template.Name}");
                return 0;
            case "roundtrip":
                if (arguments.Length != 3) throw new ArgumentException("roundtrip requires input and output XML paths.");
                SaveNew(library, arguments[2]); return 0;
            case "preview":
                if (arguments.Length < 3) throw new ArgumentException("preview requires a template name.");
                var options = new ExecutionOptions();
                foreach (var argument in arguments.Skip(3))
                {
                    var equals = argument.IndexOf('=');
                    if (equals < 1) throw new ArgumentException($"Expected Name=Value: {argument}");
                    options.Variables[argument[..equals]] = argument[(equals + 1)..];
                }
                options.PromptAsync = (request, _) =>
                {
                    Console.Error.Write($"{request.Name} [{request.DefaultValue}] (Enter uses default; EOF cancels): ");
                    var answer = Console.ReadLine();
                    return Task.FromResult(answer is { Length: 0 } ? request.DefaultValue : answer);
                };
                var result = await new TemplateEngine().ExecuteAsync(library, library.Find(arguments[2]), options);
                Console.WriteLine(JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true, Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() } }));
                return 0;
            case "audit":
                var counts = library.Templates.SelectMany(t => Enum.GetValues<TemplateArea>().SelectMany(a =>
                    Regex.Matches(t.GetCode(a), @"(?m)^\s*~~(?:_[^_\r\n]+_)?([A-Za-z]+)").Cast<Match>().Select(m => m.Groups[1].Value.ToUpperInvariant())))
                    .GroupBy(x => x).OrderByDescending(g => g.Count()).Select(g => new { Command = g.Key, Occurrences = g.Count() });
                Console.WriteLine(JsonSerializer.Serialize(new { Categories = library.Categories.Count, Templates = library.Templates.Count(), Commands = counts }, new JsonSerializerOptions { WriteIndented = true }));
                return 0;
            default: throw new ArgumentException($"Unknown command: {arguments[0]}");
        }
    }
    catch (OperationCanceledException) { Console.Error.WriteLine("Cancelled. No generated output was applied."); return 2; }
    catch (Exception ex) { Console.Error.WriteLine(ex.Message); return 1; }
}

static void SaveNew(SandyLibrary library, string path)
{
    // Never replace a source library accidentally, including through an alternate spelling of its path.
    using (new FileStream(path, FileMode.CreateNew, FileAccess.Write)) { }
    try { LibraryXml.Save(library, path); }
    catch { File.Delete(path); throw; }
}
