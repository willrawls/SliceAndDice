namespace Sandy.Core;

public sealed class PromptRequest
{
    public PromptRequest(string name, string defaultValue) { Name = name; DefaultValue = defaultValue; }
    public string Name { get; }
    public string DefaultValue { get; }
}

public enum OutputDestination { Editor, FileAppend, FileOverwrite, Clipboard, Message, Variable, Keyboard }

public sealed class OutputBlock
{
    public OutputBlock(CodeArea area, OutputDestination destination, string target, string text)
    { Area = area; Destination = destination; Target = target; Text = text; }
    public CodeArea Area { get; }
    public OutputDestination Destination { get; }
    public string Target { get; }
    public string Text { get; }
}

public sealed class ExecutionOptions
{
    public Dictionary<string, string> Variables { get; } = new(StringComparer.OrdinalIgnoreCase);
    public Func<PromptRequest, CancellationToken, Task<string?>>? PromptAsync { get; set; }
    public IList<ISandyCommandProvider> Providers { get; } = new List<ISandyCommandProvider>();
    public ILegacyHost? Host { get; set; }
    public int MaxSteps { get; set; } = 100000;
    public int MaxIncludeDepth { get; set; } = 32;
    public int MaxExpansionSteps { get; set; } = 10000;
    public int MaxOutputCharacters { get; set; } = 10000000;
    public string NewLine { get; set; } = "\r\n";
    public Func<DateTime> Clock { get; set; } = () => DateTime.Now;
}

/// <summary>Managed Sandal contract. Providers evaluate commands without committing external side effects.</summary>
public interface ISandyCommandProvider
{
    bool TryExecute(string command, string parameters, IDictionary<string, string> variables);
    bool TryEvaluateInline(string command, string parameters, IDictionary<string, string> variables, out string result);
}

public sealed class ExecutionResult
{
    public IList<OutputBlock> Outputs { get; } = new List<OutputBlock>();
    public IDictionary<string, string> Variables { get; internal set; } = new Dictionary<string, string>();
}

public sealed class SandyExecutionException : Exception
{
    public SandyExecutionException(string template, CodeArea area, int line, string message)
        : base($"{template} [{area}, line {line}]: {message}")
    { TemplateName = template; Area = area; Line = line; }
    public string TemplateName { get; }
    public CodeArea Area { get; }
    public int Line { get; }
}
