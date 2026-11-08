namespace Sandy.Core;
public sealed record Output(CodeArea Area, string Target, string Text, OutputDestination Destination = OutputDestination.Editor);
public sealed record Prompt(string Name, string Default);
public sealed record Evaluation(IReadOnlyList<Output> Outputs, IReadOnlyDictionary<string,string> Variables);
public sealed class Engine
{
    public static readonly string[] InlineCommands = ["UCASE","UPPERCASE","UPCASE","LCASE","LOWERCASE","LWCASE","TRIM","LTRIM","RTRIM","LEN","LENGTH","STRINGLENGTH","STRLEN","CHR","ASC","REPLACE","TOKENCOUNT","TCOUNT","TOKEN","BEFORE","AFTER","EXCEPT","EXTRACT","LEFT","RIGHT","MID","PAD","FIND","FINDTOKEN","NOW"];
    public async Task<Evaluation> EvaluateAsync(Library library, Template template,
        IReadOnlyDictionary<string,string>? input = null,
        Func<Prompt,CancellationToken,Task<string?>>? prompt = null, CancellationToken cancellation = default,
        ILegacyHost? host = null)
    {
        var options = new ExecutionOptions { Host = host, PromptAsync = (p,ct) => prompt == null ? Task.FromResult<string?>(p.DefaultValue) : prompt(new Prompt(p.Name,p.DefaultValue),ct) };
        if(input != null) foreach(var pair in input) options.Variables[pair.Key] = pair.Value;
        var result = await new TemplateInterpreter().ExecuteAsync(library,template,options,cancellation);
        return new(result.Outputs.Select(o=>new Output(o.Area,o.Target,o.Text,o.Destination)).ToArray(), new Dictionary<string,string>(result.Variables));
    }
}
