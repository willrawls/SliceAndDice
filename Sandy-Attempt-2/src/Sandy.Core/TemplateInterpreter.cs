using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using MetXStrings = Sandy.Core.LegacyStrings;
using MetXTokens = Sandy.Core.LegacyTokens;

namespace Sandy.Core;

/// <summary>Each execution owns its variables and buffers. Nothing is written to the IDE or filesystem.</summary>
internal sealed class TemplateInterpreter
{
    public async Task<ExecutionResult> ExecuteAsync(Library library, Template template,
        ExecutionOptions? options = null, CancellationToken cancellationToken = default)
    {
        var run = new Run(library, template, options ?? new ExecutionOptions(), cancellationToken);
        return await run.ExecuteAsync().ConfigureAwait(false);
    }

    private sealed class Run
    {
        private readonly Library library;
        private readonly Template template;
        private readonly ExecutionOptions options;
        private readonly CancellationToken cancellation;
        private readonly Dictionary<string, string> variables;
        private readonly ExecutionResult result = new();
        private readonly StringBuilder output = new();
        private Dictionary<string, string> saved = new(StringComparer.OrdinalIgnoreCase);
        private CodeArea area;
        private int line;
        private int steps;
        private int expansions;
        private int outputCharacters;
        private int inlineIncludeDepth;
        private LibraryMutation? libraryMutation;
        private bool ignoreBlanks;
        private bool endsWithNewLine;
        private OutputDestination destination;
        private string target = "";

        internal Run(Library library, Template template, ExecutionOptions options, CancellationToken cancellation)
        {
            this.library = library; this.template = template; this.options = options; this.cancellation = cancellation;
            variables = new Dictionary<string, string>(options.Variables, StringComparer.OrdinalIgnoreCase);
            foreach (var pair in new Dictionary<string, string>
            {
                ["EOL"] = options.NewLine, ["EOLTAB"] = options.NewLine + "\t",
                ["EOL2"] = options.NewLine + options.NewLine, ["EOL2Tab"] = options.NewLine + options.NewLine + "\t"
            }) if (!variables.ContainsKey(pair.Key)) variables.Add(pair.Key, pair.Value);
        }

        internal async Task<ExecutionResult> ExecuteAsync()
        {
            cancellation.ThrowIfCancellationRequested();
            foreach (var currentArea in new[] { CodeArea.Cursor, CodeArea.Declarations, CodeArea.Bottom, CodeArea.File })
            {
                area = currentArea;
                line = 1;
                try
                {
                    var source = template.Code(area);
                    if (source.Length == 0) continue;
                    options.Host?.BeginArea(area);
                    destination = area == CodeArea.File ? OutputDestination.FileAppend : OutputDestination.Editor;
                    target = area == CodeArea.File ? await ExpandAsync(template["Filename"]).ConfigureAwait(false) : "";
                    if (area == CodeArea.File && target.Length == 0) throw new InvalidOperationException("File area requires Filename.");
                    endsWithNewLine = source.EndsWith("\n", StringComparison.Ordinal);
                    var lines = Preprocess(source, library.CategoryOf(template), new HashSet<string>(StringComparer.OrdinalIgnoreCase));
                    ValidateStructure(lines);
                    await ExecuteLinesAsync(lines).ConfigureAwait(false);
                    Flush();
                }
                catch (OperationCanceledException) { throw; }
                catch (SandyExecutionException) { throw; }
                catch (Exception ex) { throw Error(ex.Message); }
            }
            result.Variables = new Dictionary<string, string>(variables, StringComparer.OrdinalIgnoreCase);
            result.LibraryMutation=libraryMutation;
            return result;
        }

        private List<string> Preprocess(string source, string category, HashSet<string> stack)
        {
            if (stack.Count >= options.MaxIncludeDepth) throw Error("Include depth limit exceeded.");
            var lines = new List<string>();
            var script = false;
            var comments = false;
            foreach (var raw in Normalize(source).Split('\n'))
            {
                cancellation.ThrowIfCancellationRequested();
                var trimmed = raw.TrimStart();
                if (trimmed.StartsWith("~##~", StringComparison.Ordinal))
                {
                    SplitCommand(trimmed.Substring(4), out var name, out var parameters);
                    switch (name)
                    {
                        case "INCLUDE":
                            var includedArea = CodeArea.Declarations; // Legacy default is memoCodeAtTop.
                            var comma = parameters.IndexOf(',');
                            if (comma > 0 && int.TryParse(parameters.Substring(0, comma), out var number))
                            {
                                includedArea = number == 2 ? CodeArea.Cursor : number == 3 ? CodeArea.Bottom : number == 1 || number == 0 ? CodeArea.Declarations : CodeArea.File;
                                parameters = parameters.Substring(comma + 1).Trim();
                            }
                            var included = library.Find(parameters, category);
                            var key = included.Id + ":" + includedArea;
                            if (!stack.Add(key)) throw Error($"Cyclic include: {included.Name}.");
                            lines.AddRange(Preprocess(included.Code(includedArea), library.CategoryOf(included), stack));
                            stack.Remove(key);
                            break;
                        case "SOFTCODE": case "BEGINSOFTCODE": case "SOFTCOMMANDS": case "BEGINSOFTCOMMANDS":
                        case "COMMANDS": case "BEGINCOMMANDS": case "SCRIPT": case "BEGINSCRIPT": case "CODE": case "BEGINCODE":
                            script = true; break;
                        case "COMMENTS": case "BEGINCOMMENTS": case "NOTES": case "BEGINNOTES": comments = true; break;
                        case "ENDCOMMENTS": case "ENDNOTES": comments = false; break;
                        case "MIXEDMODE": case "NORMAL": case "ENDSOFTCODE": case "ENDSOFTCOMMANDS": case "ENDCOMMANDS": case "ENDSCRIPT": case "ENDCODE":
                            script = false; comments = false; break;
                        default: throw Error($"Unsupported preprocessor directive: {name}.");
                    }
                }
                else if (!comments)
                    lines.Add(script && trimmed.Length > 0 ? trimmed.StartsWith("!!", StringComparison.Ordinal) ? trimmed.Substring(2) : "~~" + trimmed : raw);
            }
            // Split preserves the final empty segment. It denotes a terminator, not another output line.
            if (lines.Count > 0 && lines[lines.Count - 1] == "" && Normalize(source).EndsWith("\n", StringComparison.Ordinal)) lines.RemoveAt(lines.Count - 1);
            return lines;
        }

        private static string Normalize(string value) => value.Replace("\r\n", "\n").Replace('\r', '\n');
        private SandyExecutionException Error(string message) => new(template.Name, area, line, message);

        private static void SplitCommand(string text, out string command, out string parameters)
        {
            text = text.TrimStart().Split("~~", 2, StringSplitOptions.None)[0];
            var split = text.IndexOfAny(new[] { ' ', '\t' });
            command = (split < 0 ? text : text.Substring(0, split)).ToUpperInvariant();
            // Old command references (_i1_, _c1_, or user labels) identify nested blocks.
            if (command.StartsWith("_", StringComparison.Ordinal))
            {
                var end = command.IndexOf('_', 1);
                if (end > 0) command = command.Substring(end + 1);
            }
            parameters = split < 0 ? "" : text.Substring(split + 1).TrimStart();
        }

        private static string CommandAt(string text)
        {
            text = text.TrimStart();
            if (!text.StartsWith("~~", StringComparison.Ordinal)) return "";
            SplitCommand(text.Substring(2), out var name, out _);
            return name;
        }

        private void ValidateStructure(List<string> lines)
        {
            var blocks = new Stack<string>();
            var hasElse = new Stack<bool>();
            for (var i = 0; i < lines.Count; i++)
            {
                line = i + 1;
                var command = CommandAt(lines[i]);
                if (command == "IF") { blocks.Push("IF"); hasElse.Push(false); }
                else if (command == "ELSE")
                {
                    if (blocks.Count == 0 || blocks.Peek() != "IF" || hasElse.Pop()) throw Error("Unexpected Else.");
                    hasElse.Push(true);
                }
                else if (command == "ENDIF")
                {
                    if (blocks.Count == 0 || blocks.Pop() != "IF") throw Error("Unexpected EndIf.");
                    hasElse.Pop();
                }
                else if (command == "STARTCODA" || command == "STARTLOOPWHILE" || command == "STARTLOOPUNTIL") blocks.Push("LOOP");
                else if (command == "CODA" || command == "LOOPWHILE" || command == "LOOPUNTIL")
                {
                    if (blocks.Count == 0 || blocks.Pop() != "LOOP") throw Error("Loop has no matching start.");
                }
            }
            if (blocks.Count > 0) throw Error("Unclosed conditional or loop.");
        }

        private async Task ExecuteLinesAsync(List<string> lines)
        {
            var loops = new Stack<int>();
            for (var pc = 0; pc < lines.Count; pc++)
            {
                line = pc + 1;
                cancellation.ThrowIfCancellationRequested();
                if (++steps > options.MaxSteps) throw Error("Execution step limit exceeded (possible endless loop).");
                var raw = lines[pc];
                if (!raw.TrimStart().StartsWith("~~", StringComparison.Ordinal))
                {
                    var expanded = await ExpandAsync(raw).ConfigureAwait(false);
                    if(raw.Contains("%%INCLUDE::",StringComparison.OrdinalIgnoreCase))
                    {
                        if(++inlineIncludeDepth>options.MaxIncludeDepth)throw Error("Include depth limit exceeded.");
                        try{var includedLines=Preprocess(expanded,library.CategoryOf(template),new HashSet<string>(StringComparer.OrdinalIgnoreCase));ValidateStructure(includedLines);await ExecuteLinesAsync(includedLines).ConfigureAwait(false);}
                        finally{inlineIncludeDepth--;}
                    }
                    else Append(expanded + (pc < lines.Count - 1 || endsWithNewLine ? options.NewLine : ""));
                    continue;
                }
                if (raw.TrimStart().StartsWith("~~'", StringComparison.Ordinal)) continue;
                SplitCommand(raw.TrimStart().Substring(2), out var command, out var rawParameters);
                if(command.Contains("%%",StringComparison.Ordinal))SplitCommand(await ExpandAsync(raw.TrimStart().Substring(2)).ConfigureAwait(false),out command,out rawParameters);
                if (command == "'" || command == "REM" || command == "COMMENT" || command.Length == 0) continue;
                var parameters = await ExpandAsync(rawParameters).ConfigureAwait(false);
                if (options.Host != null && LegacyEditor.Commands.Contains(command)) Flush();
                switch (command)
                {
                    case "GOTOSANDYTEMPLATE":
                        (libraryMutation??=new(library,template)).Select(parameters);break;
                    case "SETCODEAREA":case "SETCODEAREATEXT":case "SETSOFTCODEAREA":
                        var codeChange=Assignment(parameters);(libraryMutation??=new(library,template)).SetCode(codeChange.Key,Value(codeChange.Value));break;
                    case "CHOOSE":case "DECIDE":case "YESNO":case "YESNOCANCEL":case "ASKUSER":
                        var choice=Assignment(parameters);variables[choice.Key]=await ChooseAsync(command,Value(choice.Value)).ConfigureAwait(false);break;
                    case "FOREACHCONTROL": case "FOREACHCONTROLBYFRAME":
                        Flush();
                        if(options.Host is not LegacyEditor controlEditor)throw Error("Control iteration needs a form editor host.");
                        var controlLines=new List<string>();
                        while(pc+1<lines.Count&&!lines[pc+1].TrimStart().StartsWith("~~",StringComparison.Ordinal))controlLines.Add(lines[++pc]);
                        foreach(var block in controlEditor.ExpandControlBlocks(controlLines,command=="FOREACHCONTROLBYFRAME"))Append(await ExpandAsync(block).ConfigureAwait(false));
                        break;
                    case "IF":
                        if (!SandyExpression.Truth(SandyExpression.Evaluate(parameters, variables))) pc = FindBranch(lines, pc, true);
                        break;
                    case "ELSE": pc = FindBranch(lines, pc, false); break;
                    case "ENDIF": break;
                    case "STARTCODA": case "STARTLOOPWHILE": case "STARTLOOPUNTIL": loops.Push(pc); break;
                    case "CODA": case "LOOPWHILE": case "LOOPUNTIL":
                        var repeat = SandyExpression.Truth(SandyExpression.Evaluate(parameters, variables));
                        if (command == "LOOPUNTIL") repeat = !repeat;
                        if (repeat) pc = loops.Peek(); else loops.Pop();
                        break;
                    case "SET": case "SETVALUE": case "LET": case "!": Assign(parameters, false); break;
                    case "EVAL": case "EVALUATE": case "KTEVAL": Assign(parameters, true); break;
                    case "EXECUTESOFTCODE": case "RUNSOFTCODE":
                        if(++inlineIncludeDepth>options.MaxIncludeDepth)throw Error("SoftCode execution depth limit exceeded.");
                        try{var snippet=Preprocess(Value(parameters),library.CategoryOf(template),new HashSet<string>(StringComparer.OrdinalIgnoreCase));ValidateStructure(snippet);await ExecuteLinesAsync(snippet).ConfigureAwait(false);}
                        finally{inlineIncludeDepth--;}
                        break;
                    case "APPEND":
                        var assignment = Assignment(parameters);
                        variables[assignment.Key] = Stored(assignment.Key) + DecodeMarkers(Value(assignment.Value)); break;
                    case "INCREMENT": case "INC": case "DECREMENT": case "DEC":
                        var parts = parameters.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length == 0 || parts.Length > 2) throw new FormatException("Increment/Decrement requires a variable and optional amount.");
                        var amount = parts.Length == 2 ? SandyExpression.Number(Value(parts[1])) : 1;
                        if (amount == 0) amount = 1;
                        double.TryParse(Stored(parts[0]), NumberStyles.Float, CultureInfo.InvariantCulture, out var oldNumber);
                        variables[parts[0]] = SandyExpression.Format(oldNumber + amount * (command == "INC" || command == "INCREMENT" ? 1 : -1)); break;
                    case "STRCOPY": case "COPYSTRING": case "STRINGCOPY": case "COPYSOFTVAR": case "SOFTVARCOPY":
                        var copy = Assignment(parameters); variables[copy.Key] = Value(copy.Value); break;
                    case "TRIM": case "LTRIM": case "RTRIM": case "LEFTTRIM": case "RIGHTTRIM":
                        variables[parameters] = command == "TRIM" ? Stored(parameters).Trim() : command == "LTRIM" || command == "LEFTTRIM" ? Stored(parameters).TrimStart() : Stored(parameters).TrimEnd(); break;
                    case "REPLACE":
                        var replacement = parameters.Contains('=') ? Assignment(parameters) : new KeyValuePair<string, string>(parameters, parameters);
                        var search = Stored("ToFind");
                        if (search.Length > 0) variables[replacement.Key] = Value(replacement.Value).Replace(search, Stored("ReplaceWith"));
                        break;
                    case "STRING": case "STR": case "REPSTR": case "REPLICATESTRING":
                        var repetition = Regex.Match(parameters, @"^(\S+)\s+(\d+)\s+(?:Chr\((\d+)\)|""([^""]*)"")$", RegexOptions.IgnoreCase);
                        if (!repetition.Success) throw new FormatException("String requires Variable Count Chr(code) or Variable Count \"character\".");
                        var count = Math.Max(1, int.Parse(repetition.Groups[2].Value, CultureInfo.InvariantCulture));
                        if (count > options.MaxOutputCharacters) throw Error("String size limit exceeded.");
                        var character = repetition.Groups[3].Success ? (char)int.Parse(repetition.Groups[3].Value, CultureInfo.InvariantCulture) : (repetition.Groups[4].Value + " ")[0];
                        variables[repetition.Groups[1].Value] = new string(character, count);
                        break;
                    case "SAVEVARS": saved = new Dictionary<string, string>(variables, StringComparer.OrdinalIgnoreCase); break;
                    case "RESTOREVARS": foreach (var pair in saved) variables[pair.Key + " saved"] = pair.Value; break;
                    case "IGNOREBLANKS": case "BLANKSOKAY": case "NOBLANKS": ignoreBlanks = true; break;
                    case "WATCHBLANKS": case "BLANKSNOTOKAY": case "YESBLANKS": ignoreBlanks = false; break;
                    case "CANCEL": case "CANCELINSERTION": case "ABORT": case "ABORTINSERTION":
                        if (parameters.Length == 0 || SandyExpression.Truth(SandyExpression.Evaluate(parameters, variables))) throw new OperationCanceledException("Template cancelled.", cancellation);
                        break;
                    case "NOINSERT": case "STOPCODEINSERTION": case "STOPINSERTION": return;
                    case "RESUME": case "RESUMEINSERTION": ChangeDestination(OutputDestination.Editor, ""); break;
                    case "FILENAME": case "GOTOFILE": case "OUTPUTTOFILE": case "SENDTOFILE": case "FILEOUT": ChangeDestination(OutputDestination.FileAppend, parameters); break;
                    case "GOTONEWFILE": case "OVERWRITE": case "OVERWRITEFILE": case "NEWFILE": ChangeDestination(OutputDestination.FileOverwrite, parameters); break;
                    case "GOTOCLIPBOARD": case "GOTOCLIP": case "OUTPUTTOCLIPBOARD": case "OUTTOCLIP": case "OUTTOCLIPBOARD": case "CLIPOUT": ChangeDestination(OutputDestination.Clipboard, ""); break;
                    case "GOTOMESSAGE": case "GOTOSHOWMESSAGE": case "OUTPUTTOMESSAGE": case "MESSAGEOUT": case "MSGOUT": ChangeDestination(OutputDestination.Message, ""); break;
                    case "GOTOSOFTVAR": case "GOTOSOFTVARIABLE": case "GOTOVAR": ChangeDestination(OutputDestination.Variable, parameters); break;
                    case "GOTOKB": case "GOTOKEYBOARD": case "OUTPUTTOKEYBOARD": case "OUTPUTTOKB": case "OUTTOKB": case "OUTTOKEYBOARD": case "KEYOUT": ChangeDestination(OutputDestination.Keyboard,parameters); break;
                    case "DEBUG": variables["DEBUG"] = string.Join(options.NewLine,variables.Select(p=>p.Key+"="+p.Value)); break;
                    case "MSGBOX": case "SHOWMESSAGE":
                        Flush();
                        var message = Value(parameters);
                        result.Outputs.Add(new OutputBlock(area, OutputDestination.Message, "", message));
                        options.Host?.Write(area, OutputDestination.Message, "", message); break;
                    case "WORDWRAP":
                        var wrap = Regex.Match(parameters, "^(\\d+)\\s+\"([^\"]*)\"\\s*(.*)$");
                        if (!wrap.Success || !int.TryParse(wrap.Groups[1].Value, out var width) || width < 1)
                            throw new FormatException("WordWrap requires a positive width, quoted prefix and text.");
                        var remaining = wrap.Groups[3].Value.Trim();
                        while (remaining.Length > 0)
                        {
                            var cut = remaining.Length <= width ? remaining.Length : remaining.LastIndexOf(' ', Math.Min(width, remaining.Length - 1));
                            if (cut <= 0) cut = Math.Min(width, remaining.Length);
                            Append(wrap.Groups[2].Value + remaining[..cut] + options.NewLine);
                            remaining = remaining[cut..].TrimStart();
                        }
                        break;
                    case "LEFT": case "RIGHT": case "MID":
                        var slice = Assignment(parameters); var sliceArgs = slice.Key.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                        if (!int.TryParse(sliceArgs[0], out var length)) { variables[slice.Key] = Value(slice.Value); break; }
                        var sliceValue = Value(slice.Value);
                        if (sliceArgs.Length < 2) throw new FormatException(command + " needs a length and variable name before '='.");
                        var sliceName = string.Join(" ", sliceArgs.Skip(command == "MID" && sliceArgs.Length > 2 && int.TryParse(sliceArgs[1], out _) ? 2 : 1));
                        var offset = Math.Clamp(length - 1, 0, sliceValue.Length);
                        variables[sliceName] = command == "LEFT" ? LegacyStrings.Left(sliceValue, length) : command == "RIGHT" ? LegacyStrings.Right(sliceValue, length)
                            : sliceValue.Substring(offset, sliceArgs.Length > 2 && int.TryParse(sliceArgs[1], out var take) ? Math.Clamp(take, 0, sliceValue.Length - offset) : sliceValue.Length - offset);
                        break;
                    case "BLOCK": Append(Stored(parameters)); break;
                    default:
                        if (options.Host != null && await options.Host.ExecuteAsync(command, parameters, variables, cancellation).ConfigureAwait(false)) break;
                        if (!options.Providers.Any(p => p.TryExecute(command, parameters, variables))) throw Error($"Unsupported SoftCommand: {command}.");
                        break;
                }
            }
        }

        private static int FindBranch(List<string> lines, int start, bool allowElse)
        {
            var depth = 0;
            for (var i = start + 1; i < lines.Count; i++)
            {
                var command = CommandAt(lines[i]);
                if (command == "IF") depth++;
                if (command == "ENDIF") { if (depth == 0) return i; depth--; }
                if (command == "ELSE" && depth == 0 && allowElse) return i;
            }
            throw new FormatException("Missing EndIf.");
        }

        private static KeyValuePair<string, string> Assignment(string text)
        {
            var equals = text.IndexOf('=');
            if (equals < 1) throw new FormatException("Expected Variable=Value.");
            return new KeyValuePair<string, string>(text.Substring(0, equals).Trim(), text.Substring(equals + 1));
        }

        private void Assign(string text, bool expression)
        {
            var pair = Assignment(text);
            variables[pair.Key] = expression ? SandyExpression.Evaluate(pair.Value, variables) : pair.Value;
        }

        private string Value(string text) => variables.TryGetValue(text, out var value) ? value : text;
        private string Stored(string name) => variables.TryGetValue(name, out var value) ? value : "";
        private string DecodeMarkers(string text) => text.Replace("$EOL$", options.NewLine).Replace("$13$", "\r").Replace("$9$", "\t").Replace("$SP$", " ");

        private void Append(string text)
        {
            outputCharacters += text.Length;
            if (outputCharacters > options.MaxOutputCharacters) throw Error("Output size limit exceeded.");
            output.Append(text);
        }

        private void ChangeDestination(OutputDestination next, string name)
        {
            if ((next == OutputDestination.FileAppend || next == OutputDestination.FileOverwrite || next == OutputDestination.Variable) && name.Length == 0)
                throw Error("Output destination requires a name.");
            Flush(); destination = next; target = name;
        }

        private void Flush()
        {
            if (output.Length == 0) return;
            var text = output.ToString(); output.Clear();
            if (destination == OutputDestination.Variable) variables[target] = text;
            else { result.Outputs.Add(new OutputBlock(area, destination, target, text)); options.Host?.Write(area, destination, target, text); }
        }

        private async Task<string> ExpandAsync(string text)
        {
            while (true)
            {
                cancellation.ThrowIfCancellationRequested();
                var start = text.IndexOf("%%", StringComparison.Ordinal);
                if (start < 0) return text;
                if (++expansions > options.MaxExpansionSteps) throw Error("SoftVariable expansion limit exceeded.");
                var end = text.IndexOf("%%", start + 2, StringComparison.Ordinal);
                if (end < 0) throw Error("Unclosed SoftVariable delimiter.");
                var phrase = text.Substring(start + 2, end - start - 2);
                var separator = phrase.IndexOf("::", StringComparison.Ordinal);
                var name = separator < 0 ? phrase : phrase.Substring(0, separator);
                var argument = separator < 0 ? "" : phrase.Substring(separator + 2);
                string? replacement;
                if(name.ToUpperInvariant() is "CHOOSE" or "DECIDE" or "YESNO" or "YESNOCANCEL" or "ASKUSER")replacement=await ChooseAsync(name,Value(argument)).ConfigureAwait(false);
                else if (name.ToUpperInvariant() is "CHOOSEFILE" or "PICKFILE" or "GETPATHANDFILE" or "CHOOSEDATABASE" or "CHOOSEDB" or "PICKDATABASE" or "PICKDB" or "GETDBPATHANDFILE" or "CHOOSEDIR" or "CHOOSEDIRECTORY" or "PICKDIR" or "PICKDIRECTORY" or "GETPATH" or "BROWSEFORFOLDER" or "CHOOSEFOLDER" or "BROWSEFORDIR" or "BROWSEFORDIRECTORY" or "CHOOSECOLOR" or "PICKCOLOR")
                {
                    if(options.PromptAsync == null) throw Error("An interactive host is required for " + name);
                    replacement = await options.PromptAsync(new PromptRequest(name + ": " + Value(argument), name.Equals("CHOOSE",StringComparison.OrdinalIgnoreCase)?Stored("Choose Default"):""),cancellation).ConfigureAwait(false);
                    if(replacement == null)throw new OperationCanceledException(cancellation);
                }
                else if (!TryInline(name, argument, out replacement))
                {
                    if (!variables.TryGetValue(name, out replacement) || replacement.Length == 0)
                    {
                        if (ignoreBlanks) replacement = "";
                        else
                        {
                            var defaultValue = argument.StartsWith("@", StringComparison.Ordinal) ? Value(argument.Substring(1)) : argument;
                            if (options.PromptAsync == null) throw Error($"Value required for SoftVariable '{name}'. Supply a value or a prompt handler.");
                            replacement = await options.PromptAsync(new PromptRequest(name, defaultValue), cancellation).ConfigureAwait(false);
                            if (replacement == null) throw new OperationCanceledException($"Prompt cancelled: {name}", cancellation);
                            variables[name] = replacement;
                        }
                    }
                }
                text = text.Substring(0, start) + replacement + text.Substring(end + 2);
                // Included commands must run before their following variable references are expanded.
                if(name.Equals("INCLUDE",StringComparison.OrdinalIgnoreCase))return text;
                if (text.Length > options.MaxOutputCharacters) throw Error("Expanded line size limit exceeded.");
            }
        }
        private async Task<string> ChooseAsync(string name,string argument)
        {
            var decision=!name.Equals("CHOOSE",StringComparison.OrdinalIgnoreCase);var delimiter=Stored("Choose Delimiter");if(delimiter.Length==0)delimiter=";";
            var choices=decision?new[]{"Yes","No"}:argument.Split(delimiter,StringSplitOptions.None);
            if(!decision&&choices.Length<2)return argument;
            var fallback=decision?"Yes":Stored("Choose Default");if(!choices.Contains(fallback,StringComparer.OrdinalIgnoreCase))fallback=choices[0];
            var answer=await (options.PromptAsync??throw Error("An interactive host is required for "+name))(new PromptRequest(decision?argument:"Choose an item",fallback,choices),cancellation).ConfigureAwait(false);
            if(answer==null)throw new OperationCanceledException(cancellation);
            if(decision&&answer is "1" or "0")return answer;
            if(int.TryParse(answer,out var index)&&index>=1&&index<=choices.Length)answer=choices[index-1];
            var selected=choices.FirstOrDefault(c=>c.Equals(answer,StringComparison.OrdinalIgnoreCase))??throw new FormatException("Choose one of the listed values.");
            return decision?(selected=="Yes"?"1":"0"):selected;
        }

        private bool TryInline(string name, string parameters, out string result)
        {
            if (options.Host != null && options.Host.TryInline(name.ToUpperInvariant(), parameters, variables, out result)) return true;
            foreach (var provider in options.Providers)
                if (provider.TryEvaluateInline(name.ToUpperInvariant(), parameters, variables, out result)) return true;
            var args = parameters.Split(',');
            string Arg(int i, string fallback = "") => i < args.Length ? Value(args[i]) : fallback;
            int Int(int i, int fallback) => i < args.Length && int.TryParse(Arg(i), out var value) ? value : fallback;
            var value = Arg(0);
            switch (name.ToUpperInvariant())
            {
                case "INCLUDE":
                    var includeName=parameters;var includeArea=0;
                    var includeComma=includeName.IndexOf(',');
                    if(includeComma>=0&&int.TryParse(includeName[..includeComma],out includeArea))includeName=includeName[(includeComma+1)..].Trim();
                    includeName=includeName.Replace("**CURRENT CATEGORY**",library.CategoryOf(template),StringComparison.OrdinalIgnoreCase);
                    var included=library.Find(includeName,library.CategoryOf(template));
                    result=includeArea switch{1=>included.Code(CodeArea.Declarations),2=>included.Code(CodeArea.Cursor),3=>included.Code(CodeArea.Bottom),0=>included.Code(CodeArea.Cursor)+options.NewLine+"~~GotoDec"+options.NewLine+included.Code(CodeArea.Declarations)+options.NewLine+"~~GotoEnd"+options.NewLine+included.Code(CodeArea.Bottom),_=>included.Code(CodeArea.File)};
                    return true;
                case "UCASE": case "UPPERCASE": case "UPCASE": result = value.ToUpperInvariant(); return true;
                case "LCASE": case "LOWERCASE": case "LWCASE": result = value.ToLowerInvariant(); return true;
                case "LEN": case "LENGTH": case "STRLEN": case "STRINGLENGTH": result = value.Length.ToString(CultureInfo.InvariantCulture); return true;
                case "TRIM": result = value.Trim(); return true;
                case "LTRIM": case "LEFTTRIM": result = value.TrimStart(); return true;
                case "RTRIM": case "RIGHTTRIM": result = value.TrimEnd(); return true;
                case "LEFT": result = MetXStrings.Left(value, Int(1, 0)); return true;
                case "RIGHT": result = MetXStrings.Right(value, Int(1, 0)); return true;
                case "MID":
                    var offset = Math.Min(value.Length, Math.Max(0, Int(1, 1) - 1));
                    // MetX Mid's three-argument overload takes a zero-based offset; Sandy's input is one-based.
                    result = MetXStrings.Mid(value, offset, Math.Min(value.Length - offset, Math.Max(0, Int(2, value.Length)))); return true;
                case "PAD": result = value.PadRight(Math.Max(value.Length, Int(1, 50))); return true;
                case "MANGLE": result = LegacyStrings.Mangle(value); return true;
                case "REPLACE":
                    if (args.Length < 3 || Arg(1).Length == 0) throw new FormatException("Replace requires text, search, replacement.");
                    result = value.Replace(Arg(1), Arg(2)); return true;
                case "CHR": result = char.ConvertFromUtf32(Int(0, 0)); return true;
                case "ASC": result = (value.Length == 0 ? 0 : (int)value[0]).ToString(CultureInfo.InvariantCulture); return true;
                case "INSTR": result = (value.IndexOf(Arg(1), StringComparison.Ordinal) + 1).ToString(CultureInfo.InvariantCulture); return true;
                case "TOKENCOUNT": case "LTOKENCOUNT": case "LTCOUNT": case "TCOUNT":
                    var countDelimiter = Arg(1, " ");
                    ValidateDelimiter(countDelimiter);
                    result = MetXTokens.TokenCount(value, countDelimiter, StringComparison.Ordinal).ToString(CultureInfo.InvariantCulture); return true;
                case "TOKEN": case "SGETTOKEN": case "GETTOKEN": case "STOKEN": case "BEFORE": case "SBEFORE":
                case "AFTER": case "SAFTER": case "EXCEPT": case "SEXCEPT": case "EXTRACT": case "SEXTRACT":
                    var delimiter = Arg(2, " ");
                    ValidateDelimiter(delimiter);
                    var number = Int(1, name.ToUpperInvariant().EndsWith("BEFORE", StringComparison.Ordinal) ? 2 : 1);
                    var n = Math.Max(0, number - 1);
                    var command = name.ToUpperInvariant().TrimStart('S');
                    if (command != "BEFORE" && command != "AFTER" && command != "EXCEPT" && command != "EXTRACT")
                    {
                        result = MetXTokens.TokenAt(value, Math.Max(1, number), delimiter, StringComparison.Ordinal);
                        return true;
                    }
                    var tokens = Tokens(value, delimiter);
                    result = command == "BEFORE" ? string.Join(delimiter, tokens.Take(n)) : command == "AFTER" ? string.Join(delimiter, tokens.Skip(n + 1))
                        : command == "EXCEPT" ? string.Join(delimiter, tokens.Where((_, i) => i != n)) : MetXTokens.TokenAt(value, Math.Max(1, number), delimiter, StringComparison.Ordinal);
                    if (command == "EXTRACT") variables[args[0]] = string.Join(delimiter, tokens.Where((_, i) => i != n));
                    return true;
                case "FINDTOKEN": case "LFINDTOKEN": case "FIND":
                    result = (Array.IndexOf(Tokens(value, Arg(2, " ")), Arg(1)) + 1).ToString(CultureInfo.InvariantCulture); return true;
                case "NOW":
                    result = Microsoft.VisualBasic.Strings.Format(options.Clock(),parameters.Length==0?"M/d/yyyy h:nn AM/PM":parameters); return true;
                default: result = ""; return false;
            }
        }

        private static string[] Tokens(string value, string delimiter)
        {
            ValidateDelimiter(delimiter);
            // MetX AllTokens("") yields one empty token; Sandy counts an empty input as zero tokens.
            // Use AllTokens for Before/After/Except/Extract so adjacent empty tokens are retained consistently.
            return value.Length == 0 ? Array.Empty<string>() : MetXTokens.AllTokens(value, delimiter, StringSplitOptions.None).ToArray();
        }

        private static void ValidateDelimiter(string delimiter)
        {
            if (delimiter.Length == 0) throw new FormatException("Token delimiter cannot be empty.");
        }
    }
}
