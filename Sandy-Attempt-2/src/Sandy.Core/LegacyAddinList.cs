using System.Text.RegularExpressions;
namespace Sandy.Core;
public static class LegacyAddinList
{
    public static string Update(string text,string objectName,bool remove)
    {
        objectName=objectName.Trim();if(!Regex.IsMatch(objectName,@"^[A-Za-z_][A-Za-z_0-9]*(?:\.[A-Za-z_][A-Za-z_0-9]*)+$"))throw new ArgumentException("Enter a COM object name such as Project.Addin.");
        var lines=Regex.Matches(text,@"(?m)^.+(?:\r?\n|$)|^\r?\n").Select(m=>m.Value).ToArray();
        bool Match(string line){var trimmed=line.Trim();if(trimmed.StartsWith(';')||trimmed.StartsWith('#')||trimmed.StartsWith('['))return false;var equal=trimmed.IndexOf('=');return equal>0&&trimmed[..equal].Trim().Equals(objectName,StringComparison.OrdinalIgnoreCase);}
        if(remove)return string.Concat(lines.Where(l=>!Match(l)));
        if(lines.Any(Match))return text;var eol=text.Contains("\r\n",StringComparison.Ordinal)?"\r\n":text.Contains('\n')?"\n":"\r\n";
        return text+(text.Length>0&&!text.EndsWith('\n')?eol:"")+objectName+"=0"+eol;
    }
}
