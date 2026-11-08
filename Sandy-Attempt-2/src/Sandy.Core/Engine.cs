using System.Text;
using System.Text.RegularExpressions;

namespace Sandy.Core;

public sealed record Output(CodeArea Area, string Target, string Text);
public sealed record Prompt(string Name, string Default);
public sealed record Evaluation(IReadOnlyList<Output> Outputs, IReadOnlyDictionary<string, string> Variables);

/// <summary>Fresh implementation of Sandy's %%variables%% and ~~ line commands.
/// Evaluation has no filesystem, clipboard, registry or IDE effects.</summary>
public sealed class Engine
{
    public static readonly string[] InlineCommands = ["UCASE", "UPPERCASE", "LCASE", "LOWERCASE", "TRIM", "LTRIM", "RTRIM", "LEN", "STRINGLENGTH", "CHR", "ASC", "REPLACE", "TOKENCOUNT", "TOKEN", "LEFT", "RIGHT", "MID"];
    public async Task<Evaluation> EvaluateAsync(Library library, Template template,
        IReadOnlyDictionary<string, string>? input = null,
        Func<Prompt, CancellationToken, Task<string?>>? prompt = null, CancellationToken cancellation = default)
    {
        var variables = new Dictionary<string, string>(input ?? new Dictionary<string, string>(), StringComparer.OrdinalIgnoreCase);
        variables.TryAdd("EOL", "\r\n"); variables.TryAdd("EOL2", "\r\n\r\n"); variables.TryAdd("EOLTAB", "\r\n\t");
        var output = new List<Output>();
        cancellation.ThrowIfCancellationRequested();
        var budget = 10000;
        Dictionary<string, string>? saved = null;
        bool ignoreBlanks = false;

        string Value(string name) => variables.GetValueOrDefault(name, "");
        string Inline(string command, string argument)
        {
            var args = argument.Split(',');
            string Arg(int n) => n < args.Length ? args[n].Trim() : "";
            var value = Value(Arg(0));
            if (!variables.ContainsKey(Arg(0))) value = Arg(0);
            int Number(int n, int fallback = 0) => int.TryParse(Arg(n), out var i) ? i : fallback;
            return command.ToUpperInvariant() switch
            {
                "UCASE" or "UPPERCASE" => value.ToUpperInvariant(), "LCASE" or "LOWERCASE" => value.ToLowerInvariant(),
                "TRIM" => value.Trim(), "LTRIM" => value.TrimStart(), "RTRIM" => value.TrimEnd(),
                "LEN" or "STRINGLENGTH" => value.Length.ToString(),
                "CHR" => char.ConvertFromUtf32(int.Parse(argument)), "ASC" => value.Length == 0 ? "0" : ((int)value[0]).ToString(),
                "REPLACE" => value.Replace(Arg(1), Arg(2), StringComparison.Ordinal),
                "LEFT" => value[..Math.Clamp(Number(1), 0, value.Length)],
                "RIGHT" => value[(value.Length - Math.Clamp(Number(1), 0, value.Length))..],
                "MID" => value.Substring(Math.Clamp(Number(1, 1) - 1, 0, value.Length), Math.Clamp(Number(2, value.Length), 0, value.Length - Math.Clamp(Number(1, 1) - 1, 0, value.Length))),
                "TOKENCOUNT" => value.Split(Arg(1).Length == 0 ? " " : Arg(1), StringSplitOptions.None).Length.ToString(),
                "TOKEN" => value.Split(Arg(2).Length == 0 ? " " : Arg(2), StringSplitOptions.None).ElementAtOrDefault(Number(1, 1) - 1) ?? "",
                _ => throw new NotSupportedException("Unsupported inline command: " + command)
            };
        }
        async Task<string> Expand(string source)
        {
            while (source.Contains("%%", StringComparison.Ordinal))
            {
                cancellation.ThrowIfCancellationRequested();
                if (--budget < 0) throw new InvalidOperationException("Variable expansion limit exceeded.");
                var begin = source.IndexOf("%%", StringComparison.Ordinal);
                var end = source.IndexOf("%%", begin + 2, StringComparison.Ordinal);
                if (end < 0) throw new InvalidDataException("Unclosed SoftVariable.");
                var phrase = source[(begin + 2)..end];
                var parts = phrase.Split("::", 2, StringSplitOptions.None);
                var name = parts[0];
                string value;
                if (parts.Length == 2 && InlineCommands.Contains(name.ToUpperInvariant())) value = Inline(name, parts[1]);
                else
                {
                    if (!variables.TryGetValue(name, out value!) || value.Length == 0)
                    {
                        var fallback = parts.Length == 2 ? parts[1] : "";
                        if (fallback.StartsWith('@')) fallback = Value(fallback[1..]);
                        value = ignoreBlanks ? "" : prompt == null ? fallback : await prompt(new Prompt(name, fallback), cancellation)
                            ?? throw new OperationCanceledException("Prompt cancelled.", cancellation);
                        variables[name] = value;
                    }
                }
                source = source[..begin] + value + source[(end + 2)..];
                if (source.Length > 10000000) throw new InvalidOperationException("Output limit exceeded.");
            }
            return source;
        }
        async Task<string> Render(Template current, CodeArea area, HashSet<string> includes)
        {
            var key = current["TemplateID"] + ":" + area;
            if (includes.Count >= 32 || !includes.Add(key)) throw new InvalidOperationException("Recursive template include: " + current.Name);
            try
            {
                var lines = current.Code(area).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
                var buffer = new StringBuilder();
                for (var i = 0; i < lines.Length; i++)
                {
                    cancellation.ThrowIfCancellationRequested();
                    if (--budget < 0) throw new InvalidOperationException("Execution limit exceeded.");
                    var line = lines[i];
                    var trimmed = line.TrimStart();
                    if (trimmed.StartsWith("~~") || trimmed.StartsWith("~##~"))
                    {
                        var commandLine = trimmed[(trimmed.StartsWith("~##~") ? 4 : 2)..].Trim();
                        var split = commandLine.IndexOfAny([' ', '\t']);
                        var command = (split < 0 ? commandLine : commandLine[..split]).ToUpperInvariant();
                        var parameter = split < 0 ? "" : commandLine[(split + 1)..].Trim();
                        var equal = parameter.IndexOf('=');
                        var result = (equal < 0 ? parameter : parameter[..equal]).Trim();
                        var expression = equal < 0 ? "" : parameter[(equal + 1)..];
                        switch (command)
                        {
                            case "": case "'": case "*": case "REM": case "COMMENT": break;
                            case "SET": case "SETVALUE": variables[result] = await Expand(expression); break;
                            case "STRCOPY": variables[result] = variables.GetValueOrDefault(expression, await Expand(expression)); break;
                            case "APPEND": variables[result] = Value(result) + await Expand(expression.Replace("$EOL$", "\r\n").Replace("$13$", "\r").Replace("$9$", "\t").Replace("$SP$", " ")); break;
                            case "INC": case "INCREMENT": case "DEC": case "DECREMENT":
                                variables[result] = ((int.TryParse(Value(result), out var n) ? n : 0) + (command.StartsWith("I") ? 1 : -1)).ToString(); break;
                            case "TRIM": case "LTRIM": case "RTRIM": variables[result] = Inline(command, result); break;
                            case "SAVEVARS": saved = new(variables, StringComparer.OrdinalIgnoreCase); break;
                            case "RESTOREVARS": if (saved != null) foreach (var pair in saved) variables[pair.Key + " saved"] = pair.Value; break;
                            case "IGNOREBLANKS": ignoreBlanks = true; break;
                            case "INCLUDE":
                                var args = parameter.Split(',', 2);
                                var includeArea = CodeArea.Declarations;
                                if (args.Length == 2 && int.TryParse(args[0], out var a)) includeArea = a switch { 2 => CodeArea.Cursor, 3 => CodeArea.Bottom, 4 => CodeArea.File, _ => CodeArea.Declarations };
                                buffer.Append(await Render(library.Find(args[^1].Trim()), includeArea, includes)); break;
                            default: throw new NotSupportedException($"{current.Name}, {area}, line {i + 1}: unsupported SoftCommand '{command}'.");
                        }
                    }
                    else { buffer.Append(await Expand(line)); if (i < lines.Length - 1) buffer.Append("\r\n"); }
                    if (buffer.Length > 10000000) throw new InvalidOperationException("Output limit exceeded.");
                }
                return buffer.ToString();
            }
            finally { includes.Remove(key); }
        }
        foreach (var area in new[] { CodeArea.Cursor, CodeArea.Declarations, CodeArea.Bottom, CodeArea.File })
            if (template.Code(area).Length != 0)
            {
                var text = await Render(template, area, []);
                var target = area == CodeArea.File ? await Expand(template["Filename"]) : "";
                if (area == CodeArea.File && target.Length == 0) throw new InvalidOperationException("File output requires a filename.");
                output.Add(new(area, target, text));
            }
        return new(output, variables);
    }
}
