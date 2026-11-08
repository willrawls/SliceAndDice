using System.Text;
using System.Text.RegularExpressions;
namespace Sandy.Core;
public sealed record HeaderEntry(string Kind,string Name,string Declaration);
public sealed record HeaderInspectionResult(string Source,IReadOnlyList<HeaderEntry> Entries,IReadOnlyList<string> Diagnostics);
public static class HeaderInspection
{
    public static string StripComments(string source)=>Regex.Replace(source,"\"(?:\\\\.|[^\"\\\\])*\"|'(?:\\\\.|[^'\\\\])*'|/\\*[\\s\\S]*?\\*/|//[^\\r\\n]*",m=>m.Value.StartsWith("/*",StringComparison.Ordinal)||m.Value.StartsWith("//",StringComparison.Ordinal)?new string(' ',m.Length):m.Value);
    public static HeaderInspectionResult Read(string file,bool includes=false,IEnumerable<string>? includeDirectories=null)
    {
        var visited=new HashSet<string>(StringComparer.OrdinalIgnoreCase);var roots=new[]{Path.GetDirectoryName(Path.GetFullPath(file))!}.Concat(includeDirectories??[]).Select(Path.GetFullPath).ToArray();var notes=new List<string>();var bytes=0;
        string Load(string path)
        {
            path=Path.GetFullPath(path);if(!visited.Add(path))return "";if(visited.Count>128)throw new InvalidDataException("Header include limit exceeded.");var text=TextFiles.Read(path);bytes+=text.Length;if(bytes>5_000_000)throw new InvalidDataException("Header source size limit exceeded.");
            return Regex.Replace(text,@"(?m)^\s*#\s*include\s*[""<]([^"">]+)["">][^\r\n]*",match=>{
                if(!includes)return "";
                var candidates=new[]{Path.GetDirectoryName(path)!}.Concat(roots).Select(root=>Path.GetFullPath(Path.Combine(root,match.Groups[1].Value)));
                var found=candidates.FirstOrDefault(candidate=>roots.Any(root=>candidate.StartsWith(Path.TrimEndingDirectorySeparator(root)+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase))&&File.Exists(candidate));
                if(found==null){notes.Add("Include not found in selected roots: "+match.Groups[1].Value);return "";}return Load(found);
            });
        }
        var source=Load(file);var inspected=Inspect(source);return inspected with{Diagnostics=notes.Concat(inspected.Diagnostics).ToArray()};
    }
    public static HeaderInspectionResult Inspect(string source)
    {
        var entries=new List<HeaderEntry>();var text=StripComments(source);var notes=new List<string>();
        foreach(Match macro in Regex.Matches(text,@"(?m)^\s*#\s*define\s+(\w+)(\([^\r\n]*?\))?\s*([^\r\n]*)"))entries.Add(new(macro.Groups[2].Success?"Function macro":"Define",macro.Groups[1].Value,macro.Groups[3].Value.Trim()));
        foreach(Match declaration in Regex.Matches(text,@"\b(class|struct|union|enum)\s+(\w+)\s*([^;{}]*)\{"))
        {
            var start=declaration.Index+declaration.Length;var depth=1;var end=start;var quote='\0';
            for(;end<text.Length&&depth>0;end++){var ch=text[end];if(quote!='\0'){if(ch=='\\'){end++;continue;}if(ch==quote)quote='\0';continue;}if(ch is '\'' or '"'){quote=ch;continue;}if(ch=='{')depth++;else if(ch=='}')depth--;}
            if(depth!=0){notes.Add("Unclosed declaration: "+declaration.Groups[2].Value);continue;}
            var name=declaration.Groups[2].Value;entries.Add(new(declaration.Groups[1].Value,name,declaration.Groups[3].Value.Trim()));
            if(declaration.Groups[1].Value!="class")continue;
            var body=text[start..(end-1)];var scope="private";
            foreach(var raw in body.Split(';',StringSplitOptions.RemoveEmptyEntries))
            {
                var member=raw.Trim();var access=Regex.Match(member,@"^(public|private|protected)\s*:\s*");if(access.Success){scope=access.Groups[1].Value;member=member[access.Length..].Trim();}
                if(member.Length==0)continue;var method=Regex.Match(member,@"([~\w]+)\s*\(");
                entries.Add(new(method.Success?"Method":"Property",name+"."+(method.Success?method.Groups[1].Value:Regex.Match(member,@"(\w+)\s*(?:\[[^\]]*\])?$").Groups[1].Value),scope+" "+member));
            }
        }
        return new(source,entries,notes);
    }
}
