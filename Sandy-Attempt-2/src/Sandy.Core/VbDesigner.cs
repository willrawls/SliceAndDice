using System.Text.RegularExpressions;
namespace Sandy.Core;

internal sealed class DesignerBlock
{
    internal string Type="",Name="",Indent="";internal bool Property;internal int Start,ContentStart,EndStart,End;internal DesignerBlock? Parent;
    internal List<DesignerBlock> Children=[];internal Dictionary<string,(string Value,int Start,int Length)> Values=new(StringComparer.OrdinalIgnoreCase);
    internal string Key=>Name+(Values.TryGetValue("Index",out var index)?"("+Regex.Match(index.Value,@"^-?\d+").Value+")":"");
}
internal static class VbDesigner
{
    internal static DesignerBlock Parse(string text)
    {
        var stack=new Stack<DesignerBlock>();DesignerBlock? root=null;
        foreach(Match line in Regex.Matches(text,@"(?m)^.+(?:\r?\n|$)"))
        {
            var trimmed=line.Value.TrimEnd('\r','\n');var begin=Regex.Match(trimmed,@"^(\s*)Begin(?:Property)?\s+(.*)$",RegexOptions.IgnoreCase);
            if(begin.Success)
            {
                var property=trimmed.TrimStart().StartsWith("BeginProperty ",StringComparison.OrdinalIgnoreCase);var words=begin.Groups[2].Value.Split(' ',StringSplitOptions.RemoveEmptyEntries);
                if(words.Length<(property?1:2))throw new InvalidDataException("Invalid VB designer block.");
                var block=new DesignerBlock{Property=property,Name=words[property?0:1],Type=property?"":words[0],Indent=begin.Groups[1].Value,Start=line.Index,ContentStart=line.Index+line.Length,Parent=stack.Count>0?stack.Peek():null};
                if(block.Parent==null){if(root!=null)continue;root=block;}else block.Parent.Children.Add(block);stack.Push(block);continue;
            }
            if(Regex.IsMatch(trimmed,@"^\s*End(?:Property)?\s*$",RegexOptions.IgnoreCase)&&stack.Count>0){var block=stack.Pop();block.EndStart=line.Index;block.End=line.Index+line.Length;continue;}
            if(stack.Count>0){var field=Regex.Match(trimmed,@"^\s*(\w+)\s*=\s*(.*)$");if(field.Success)stack.Peek().Values[field.Groups[1].Value]=(field.Groups[2].Value,line.Index+field.Groups[2].Index,field.Groups[2].Length);}
        }
        if(root==null||stack.Count>0)throw new InvalidDataException("Invalid or unclosed VB6 form definition.");return root;
    }
    internal static IEnumerable<DesignerBlock> Controls(DesignerBlock root)=>root.Children.Where(c=>!c.Property).SelectMany(c=>new[]{c}.Concat(Controls(c)));
    internal static Dictionary<string,string> Properties(DesignerBlock block)
    {
        var values=block.Values.ToDictionary(p=>p.Key,p=>p.Value.Value,StringComparer.OrdinalIgnoreCase);
        foreach(var property in block.Children.Where(c=>c.Property))foreach(var pair in Properties(property))values[property.Name+"."+pair.Key]=pair.Value;
        return values;
    }
    internal static string Set(string text,string? control,string property,string literal)
    {
        var root=Parse(text);var block=control==null?root:Controls(root).Single(c=>c.Key.Equals(control,StringComparison.OrdinalIgnoreCase));var path=property.Split('.',2);
        if(path.Length==2)
        {
            var nested=block.Children.FirstOrDefault(c=>c.Property&&c.Name.Equals(path[0],StringComparison.OrdinalIgnoreCase));
            if(nested==null){var indent=block.Indent+"   ";var font=path[0].Equals("Font",StringComparison.OrdinalIgnoreCase)?" {0BE35203-8F91-11CE-9DE3-00AA004BB851}":"";text=text.Insert(block.ContentStart,indent+"BeginProperty "+path[0]+font+"\r\n"+indent+"EndProperty\r\n");return Set(text,control,property,literal);}
            block=nested;property=path[1];
            if(property.Contains('.'))throw new NotSupportedException("Designer property nesting beyond one object is not supported.");
        }
        if(!Regex.IsMatch(property,@"^\w+$"))throw new FormatException("Invalid designer property name.");
        if(property.Equals("Name",StringComparison.OrdinalIgnoreCase)&&!block.Property)
        {
            var name=literal.Trim('"');if(!Regex.IsMatch(name,@"^[A-Za-z_]\w*$"))throw new FormatException("Invalid VB designer name.");
            if(Controls(root).Any(c=>c.Name.Equals(name,StringComparison.OrdinalIgnoreCase)&&!c.Name.Equals(block.Name,StringComparison.OrdinalIgnoreCase)))throw new InvalidOperationException("A control with that name already exists.");
            var old=block.Name;var renamed=text;
            foreach(var target in (block==root?new[]{root}:Controls(root).Where(c=>c.Name.Equals(old,StringComparison.OrdinalIgnoreCase))).OrderByDescending(c=>c.Start))
            {
                var header=renamed[target.Start..target.ContentStart];var match=Regex.Match(header,@"^(\s*Begin\s+\S+\s+)(\S+)",RegexOptions.IgnoreCase);
                renamed=renamed[..(target.Start+match.Groups[2].Index)]+name+renamed[(target.Start+match.Groups[2].Index+match.Groups[2].Length)..];
            }
            if(block==root)renamed=Regex.Replace(renamed,@"(?m)^(Attribute VB_Name\s*=\s*)""[^""]*""",m=>m.Groups[1].Value+"\""+name+"\"",RegexOptions.IgnoreCase);
            return renamed;
        }
        if(block.Values.TryGetValue(property,out var current))return text[..current.Start]+literal+text[(current.Start+current.Length)..];
        return text.Insert(block.ContentStart,block.Indent+"   "+property+" = "+literal+"\r\n");
    }
    internal static string Add(string text,string type,string key,IReadOnlyDictionary<string,string> values,Func<string,string> literal)
    {
        var root=Parse(text);var match=Regex.Match(key,@"^(\w+)(?:\((-?\d+)\))?$");if(!match.Success)throw new FormatException("Control name must be a VB identifier with an optional array index.");
        var parent=values.TryGetValue("$Parent",out var container)&&container.Length>0?Controls(root).Single(c=>c.Key.Equals(container,StringComparison.OrdinalIgnoreCase)):root;var indent=parent.Indent+"   ";
        var block=indent+"Begin "+type+" "+match.Groups[1].Value+"\r\n";
        if(match.Groups[2].Success)block+=indent+"   Index = "+match.Groups[2].Value+"\r\n";
        foreach(var p in values.Where(p=>!p.Key.StartsWith('$')&&p.Key!="Index"&&!p.Key.Contains('.')))block+=indent+"   "+p.Key+" = "+literal(p.Value)+"\r\n";
        text=text.Insert(parent.EndStart,block+indent+"End\r\n");foreach(var p in values.Where(p=>p.Key.Contains('.')))text=Set(text,key,p.Key,literal(p.Value));return text;
    }
}
