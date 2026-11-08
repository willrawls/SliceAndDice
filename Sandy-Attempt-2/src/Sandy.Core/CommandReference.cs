using System.Text.RegularExpressions;
namespace Sandy.Core;
public sealed record CommandReference(string Name,string Kind,string Aliases,string Syntax,string Summary,string Details,string Examples,string Source)
{
    public static IReadOnlyList<CommandReference> Parse(string text,string source)
    {
        var result=new List<CommandReference>();var matches=Regex.Matches(text,@"(?m)^\*\* (Command|Inline) (.+)\r?$");
        for(var i=0;i<matches.Count;i++)
        {
            var m=matches[i];var block=text[(m.Index+m.Length)..(i+1<matches.Count?matches[i+1].Index:text.Length)];
            var sections=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);var directives=Regex.Matches(block,@"(?m)^\*\* (\w+)\s*([^\r\n]*)\r?$");
            for(var j=0;j<directives.Count;j++)
            {
                var d=directives[j];var content=d.Groups[2].Value.TrimEnd('\r');var tail=block[(d.Index+d.Length)..(j+1<directives.Count?directives[j+1].Index:block.Length)];
                content+=(tail.Trim().Length>0?"\r\n"+tail.Trim():"");sections[d.Groups[1].Value]=content.Replace("** '","").Trim();
            }
            result.Add(new(m.Groups[2].Value.Trim(),m.Groups[1].Value,sections.GetValueOrDefault("Alias",""),sections.GetValueOrDefault("Syntax",""),sections.GetValueOrDefault("Short",""),sections.GetValueOrDefault("Long",sections.GetValueOrDefault("Description","")),sections.GetValueOrDefault("Example",""),source));
        }
        return result;
    }
}
