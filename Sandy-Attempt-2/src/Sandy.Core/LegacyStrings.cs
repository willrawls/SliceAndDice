namespace Sandy.Core;

internal static class LegacyStrings
{
    public static string Mangle(string value)
    {
        if(!value.Contains('_'))return value;
        var result=new System.Text.StringBuilder();var upper=true;
        foreach(var character in value)
        {
            if(character=='_'){upper=true;continue;}
            if(char.IsAsciiLetter(character)){result.Append(upper?char.ToUpperInvariant(character):char.ToLowerInvariant(character));upper=false;}
            else if(char.IsAsciiDigit(character)){result.Append(character);upper=false;}
        }
        return result.ToString();
    }
    public static string Left(string value, int length) => value[..Math.Clamp(length, 0, value.Length)];
    public static string Right(string value, int length) => value[(value.Length - Math.Clamp(length, 0, value.Length))..];
    public static string Mid(string value, int start, int length) => value.Substring(start, length);
}
internal static class LegacyTokens
{
    public static int TokenCount(string value, string delimiter, StringComparison comparison) => value.Length == 0 ? 0 : value.Split(delimiter, StringSplitOptions.None).Length;
    public static string TokenAt(string value, int number, string delimiter, StringComparison comparison) => value.Split(delimiter, StringSplitOptions.None).ElementAtOrDefault(number - 1) ?? "";
    public static IEnumerable<string> AllTokens(string value, string delimiter, StringSplitOptions options) => value.Split(delimiter, options);
}

public interface ILegacyHost
{
    void BeginArea(CodeArea area);
    Task<bool> ExecuteAsync(string command, string parameters, IDictionary<string,string> variables, CancellationToken cancellation);
    bool TryInline(string command, string parameters, IDictionary<string,string> variables, out string value);
    void Write(CodeArea area, OutputDestination destination, string target, string text);
}

public sealed record ExternalEffect(string Kind, string Target, string Value = "", byte[]? Data = null, string? ResultToken = null);
public sealed class PlannedDocument(string path, string original)
{
    public string Path { get; } = path;
    public string Original { get; } = original;
    public string Text { get; set; } = original;
    public bool Changed => Text != Original;
}
