using System.Text;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis.CSharp;

namespace Sandy.Core;
public sealed record SchemaField(string Name,int DaoType,int Length=0,bool Nullable=true);
public sealed record GeneratedCode(string Code,IReadOnlyList<string> Diagnostics);

public static class DatabaseGenerator
{
    public static string Identifier(string name)
    {
        var value=Regex.Replace(name,@"[^A-Za-z0-9_]","_");if(value.Length==0)value="Unnamed";if(char.IsDigit(value[0]))value="_"+value;
        return SyntaxFacts.GetKeywordKind(value)!=SyntaxKind.None?"@"+value:value;
    }
    public static string Generate(string table,IEnumerable<SchemaField> fields,string nameSpace="Generated")
    {
        var b=new StringBuilder("#nullable enable\r\nusing System;\r\n\r\nnamespace "+string.Join(".",nameSpace.Split('.').Select(Identifier))+";\r\n\r\npublic sealed class "+Identifier(table)+"\r\n{\r\n");
        foreach(var field in fields)
        {
            var type=field.DaoType switch{1=>"bool",2=>"byte",3=>"short",4=>"int",5=>"decimal",6=>"float",7=>"double",8=>"DateTime",9 or 11=>"byte[]",10 or 12=>"string",15=>"Guid",16=>"long",20=>"decimal",_=>"object"};
            var nullable=field.Nullable?"?":"";var initializer=!field.Nullable&&type=="string"?" = string.Empty;":!field.Nullable&&type=="byte[]"?" = Array.Empty<byte>();":!field.Nullable&&type=="object"?" = new();":"";
            b.Append("    public ").Append(type).Append(nullable).Append(' ').Append(Identifier(field.Name)).Append(" { get; set; }").Append(initializer).Append("\r\n");
        }
        return b.Append("}\r\n").ToString();
    }
    public static IEnumerable<SchemaField> ParseDaoFields(IEnumerable<string> fields)
    {
        foreach(var field in fields)
        {
            var m=Regex.Match(field,@"^(.*?)\s*:\s*DAO\s+(\d+)\s*\((\d+)\)$");
            if(!m.Success)throw new FormatException("Unrecognized schema field: "+field);
            yield return new(m.Groups[1].Value.Trim(),int.Parse(m.Groups[2].Value),int.Parse(m.Groups[3].Value));
        }
    }
}

public static class HeaderGenerator
{
    private static readonly Dictionary<string,string> Types=new(StringComparer.Ordinal)
    {
        ["int"]="int",["long"]="int",["unsigned int"]="uint",["unsigned long"]="uint",["short"]="short",["unsigned short"]="ushort",["char"]="byte",["unsigned char"]="byte",
        ["float"]="float",["double"]="double",["bool"]="bool",["BOOL"]="int",["BYTE"]="byte",["WORD"]="ushort",["DWORD"]="uint",["LONG"]="int",["UINT"]="uint",["HANDLE"]="IntPtr",["void*"]="IntPtr"
    };
    public static GeneratedCode Generate(string header)
    {
        var notes=new List<string>();var b=new StringBuilder("using System;\r\nusing System.Runtime.InteropServices;\r\n\r\n");
        header=Regex.Replace(header,@"/\*.*?\*/|//[^\r\n]*","",RegexOptions.Singleline);
        var defines=Regex.Matches(header,@"(?m)^\s*#define\s+(\w+)\s+([^\r\n]+)");
        if(defines.Count>0)
        {
            b.Append("public static class NativeConstants\r\n{\r\n");
            foreach(Match m in defines)
            {
                var value=m.Groups[2].Value.Trim();var kind=value.StartsWith('"')?"string":Regex.IsMatch(value,@"^-?(?:0x[0-9A-Fa-f]+|\d+)$")?"int":"";
                if(kind.Length==0){notes.Add("Unconverted macro: "+m.Groups[1].Value);continue;}
                b.Append("    public const ").Append(kind).Append(' ').Append(DatabaseGenerator.Identifier(m.Groups[1].Value)).Append(" = ").Append(value).Append(";\r\n");
            }
            b.Append("}\r\n\r\n");
        }
        foreach(Match m in Regex.Matches(header,@"(?:typedef\s+)?(struct|union|enum)\s*(\w*)\s*\{([^{}]*)\}\s*(\w*)\s*;",RegexOptions.Singleline))
        {
            var name=m.Groups[4].Value.Length>0?m.Groups[4].Value:m.Groups[2].Value;if(name.Length==0){notes.Add("Anonymous declaration needs a name.");continue;}
            if(m.Groups[1].Value=="enum") { b.Append("public enum ").Append(DatabaseGenerator.Identifier(name)).Append("\r\n{\r\n").Append(m.Groups[3].Value.Trim()).Append("\r\n}\r\n\r\n");continue; }
            var union=m.Groups[1].Value=="union";b.Append("[StructLayout(LayoutKind.").Append(union?"Explicit":"Sequential").Append(")]\r\npublic struct ").Append(DatabaseGenerator.Identifier(name)).Append("\r\n{\r\n");
            foreach(var declaration in m.Groups[3].Value.Split(';',StringSplitOptions.RemoveEmptyEntries))
            {
                if(string.IsNullOrWhiteSpace(declaration))continue;
                var field=Regex.Match(declaration.Trim(),@"^([\w\s*]+?)\s+(\w+)\s*(?:\[\s*(\d+)\s*\])?$");
                if(!field.Success){notes.Add(name+": unconverted field "+declaration.Trim());continue;}
                var nativeType=Regex.Replace(field.Groups[1].Value.Trim(),@"\s+"," ");
                if(!Types.TryGetValue(nativeType,out var type)){if(nativeType.Contains('*'))type="IntPtr";else{notes.Add(name+": unknown native type "+nativeType);continue;}}
                if(union)b.Append("    [FieldOffset(0)]\r\n");
                if(field.Groups[3].Success){b.Append("    [MarshalAs(UnmanagedType.ByValArray, SizeConst = ").Append(field.Groups[3].Value).Append(")]\r\n");type+="[]";}
                b.Append("    public ").Append(type).Append(' ').Append(DatabaseGenerator.Identifier(field.Groups[2].Value)).Append(";\r\n");
            }
            b.Append("}\r\n\r\n");
        }
        if(!Regex.IsMatch(header,@"\b(struct|union|enum)\b")&&defines.Count==0)notes.Add("No supported declarations were found.");
        return new(b.ToString(),notes);
    }
}

public static class LegacyObfuscation
{
    static LegacyObfuscation()=>Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    private static string Scramble(string value)
    {
        var chars=value.ToCharArray();for(var i=0;i+4<chars.Length;i+=4){(chars[i],chars[i+3])=(chars[i+3],chars[i]);(chars[i+1],chars[i+2])=(chars[i+2],chars[i+1]);}return new string(chars);
    }
    public static string Encrypt(string text)=>"EN* "+Scramble(string.Concat(Encoding.GetEncoding(1252).GetBytes(text).Select(b=>(255-b).ToString("X2"))));
    public static string Decrypt(string text)
    {
        if(!text.StartsWith("EN* ",StringComparison.Ordinal))return "";var hex=Scramble(text[4..]);if(hex.Length%2!=0)throw new FormatException("Invalid legacy encrypted text.");
        return Encoding.GetEncoding(1252).GetString(Enumerable.Range(0,hex.Length/2).Select(i=>(byte)((255-Convert.ToInt32(hex.Substring(i*2,2),16))%255)).ToArray());
    }
}
