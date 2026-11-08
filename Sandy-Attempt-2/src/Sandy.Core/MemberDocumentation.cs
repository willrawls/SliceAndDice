using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System.Text.RegularExpressions;
namespace Sandy.Core;
internal static class MemberDocumentation
{
    internal static (IReadOnlyList<(string Name,string Description)> Parameters,string Returns) Describe(string source,int position)
    {
        var method=CSharpSyntaxTree.ParseText(source).GetRoot().DescendantNodes().OfType<MethodDeclarationSyntax>().FirstOrDefault(m=>m.FullSpan.Contains(position));
        if(method!=null)return(method.ParameterList.Parameters.Select(p=>(p.Identifier.ValueText,
            (p.Modifiers.Any(m=>m.ValueText=="out")?"(O) ":p.Modifiers.Any(m=>m.ValueText=="ref")?"(IO) ":"(I) ")+p.Type+(p.Default==null?"":" Defaults to "+p.Default.Value))).ToArray(),method.ReturnType.ToString()=="void"?"":method.ReturnType.ToString());
        var headers=Regex.Matches(source,@"(?im)^\s*(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+\s*\((.*?)\)(?:\s+As\s+(\w+))?[^\r\n]*");
        var header=headers.Cast<Match>().LastOrDefault(m=>m.Index<=position)??throw new InvalidOperationException("No containing method was found.");
        var parameters=new List<(string,string)>();
        foreach(var raw in header.Groups[1].Value.Split(',',StringSplitOptions.RemoveEmptyEntries))
        {
            var p=Regex.Match(raw.Trim(),@"^(Optional\s+)?(?:(ByVal|ByRef)\s+)?(\w+)(\(\))?(?:\s+As\s+(\w+))?(?:\s*=\s*(.*))?$",RegexOptions.IgnoreCase);
            if(!p.Success)throw new FormatException("Unrecognized parameter: "+raw);
            parameters.Add((p.Groups[3].Value,(p.Groups[2].Value.Equals("ByVal",StringComparison.OrdinalIgnoreCase)?"(I) ":"(IO) ")+(p.Groups[1].Success?"Opt. ":"")+(p.Groups[5].Success?p.Groups[5].Value:"Variant")+(p.Groups[4].Success?" Array":"")+(p.Groups[6].Success?" Defaults to "+p.Groups[6].Value:"")));
        }
        return(parameters,header.Groups[2].Value);
    }
}
