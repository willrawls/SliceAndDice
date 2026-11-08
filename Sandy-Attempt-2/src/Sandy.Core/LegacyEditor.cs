using System.Text.RegularExpressions;

namespace Sandy.Core;

/// <summary>A virtual editor. Every command changes a plan, never the live IDE or filesystem.</summary>
public sealed partial class LegacyEditor : ILegacyHost
{
    public static readonly HashSet<string> Commands = new(StringComparer.OrdinalIgnoreCase)
    {
        "GOTOMODULE","GOTOCLASS","GOTOFORM","GOTOPROJECT","GOTOPROC","GOTOPROCEND","PROCTOP","PROCEND",
        "GOTODECLARATIONS","GOTODECLARTIONS","GOTODEC","GOTOENDOFFILE","GOTOENDOFMODULE","GOTOEND","ABSLINE","LINEOFFSET","OFFSET",
        "FIND","LOCATE","SEARCH","FINDINPROC","PROCFIND","PFIND","PROCLOCATE","PLOCATE","PROCSEARCH","PSEARCH",
        "DELETEPROC","COPYPROC","CUTPROC","DELETELINES","DELETELINE","READLINE","NEXTLINE","POSTFIXLINE","POSTFIX","PREFIXLINE","PREFIX",
        "GETTEXTSELECTION","GETTEXT","GETSELECTION","DELETESELECTION","REPLACELINE","REPLACEINMODULE","MODULEREPLACE",
        "FIRSTSELECTIONLINE","LASTSELECTIONLINE","ADDCONTROL","SELECTCONTROL","SETPROPERTY","SETFORMPROPERTY","FORMPROPERTY","SETDEFAULT","RESETDEFAULTS","PROCATTR","COMMENTEDPARAMETERS"
    };
    public Dictionary<string, PlannedDocument> Documents { get; } = new(StringComparer.OrdinalIgnoreCase);
    public List<ExternalEffect> Effects { get; } = [];
    public string ActivePath { get; private set; }
    public PlannedDocument Active => Documents[ActivePath];
    public int Position { get; private set; }
    public int SelectionLength { get; private set; }
    public string ProjectName { get; set; } = "SandyGenerated";
    public string Clipboard { get; set; } = "";
    public Func<string>? ClipboardReader { get; set; }
    public Func<string,string>? FileReader { get; set; }
    public Func<string,string>? DirectoryReader { get; set; }
    public Func<string,string?>? RegistryReader { get; set; }
    public Func<string,string,CancellationToken,Task<string?>>? AskAsync { get; set; }
    public string SelectedControl { get; private set; } = "";
    public Dictionary<string, Dictionary<string,string>> Controls { get; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, string> FormProperties { get; } = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, Dictionary<string,string>> defaults = new(StringComparer.OrdinalIgnoreCase);
    public LegacyEditor(string path, string text, int selectionStart = 0, int selectionLength = 0)
    {
        if (selectionStart < 0 || selectionLength < 0 || selectionStart > text.Length - selectionLength) throw new ArgumentOutOfRangeException(nameof(selectionStart));
        ActivePath = path; Documents.Add(path, new(path,text)); Position = selectionStart; SelectionLength = selectionLength; LoadDesigner();
    }
    public void BeginArea(CodeArea area)
    {
        if (area == CodeArea.Declarations) { Position = SourceNavigation.TypeBoundaries(Active.Text,Position)?.Top ?? DeclarationPoint(Active.Text); SelectionLength = 0; }
        else if (area == CodeArea.Bottom) { Position = SourceNavigation.TypeBoundaries(Active.Text,Position)?.Bottom ?? BottomPoint(Active.Text); SelectionLength = 0; }
    }
    public void Write(CodeArea area, OutputDestination destination, string target, string text)
    {
        if (destination == OutputDestination.Clipboard) { Clipboard = text; Effects.Add(new("Clipboard", "", text)); return; }
        if (destination == OutputDestination.Message) { Effects.Add(new("Message", "", text)); return; }
        if (destination == OutputDestination.Keyboard) { Effects.Add(new("SendKeys",text.Replace("\r\n","\n").Replace("\n",target.Equals("RAW",StringComparison.OrdinalIgnoreCase)?"":"{ENTER}")));return; }
        if (destination is OutputDestination.FileAppend or OutputDestination.FileOverwrite) { RecordEffect(new(destination.ToString(), PathName(target), text)); return; }
        if (destination != OutputDestination.Editor) return;
        Replace(Position, SelectionLength, text); Position += text.Length; SelectionLength = 0;
    }
    private static bool IsVisualBasic(string text) => Regex.IsMatch(text, @"(?im)^\s*(?:VERSION\s+\d|(?:Public |Private |Friend )?(?:Sub|Function|Property|Module)\s)") || text.Contains("End Class", StringComparison.OrdinalIgnoreCase);
    internal static int DeclarationPoint(string text)
    {
        if (IsVisualBasic(text))
        {
            var match = Regex.Match(text, @"(?im)^\s*(?:(?:Public|Private|Friend)\s+)?(?:Sub|Function|Property)\s");
            return match.Success ? match.Index : 0;
        }
        var type = Regex.Match(text, @"\b(?:class|struct|record|interface)\s+\w+[^\{]*\{");
        return type.Success ? type.Index + type.Length : 0;
    }
    internal static int BottomPoint(string text)
    {
        if (IsVisualBasic(text)) { var end = Regex.Match(text, @"(?im)^\s*End\s+(?:Class|Module)\b"); return end.Success ? end.Index : text.Length; }
        var type = Regex.Match(text, @"\b(?:class|struct|record|interface)\s+\w+[^\{]*\{");
        if (!type.Success) return text.Length;
        var depth = 1; for (var i = type.Index + type.Length; i < text.Length; i++) { if (text[i] == '{') depth++; if (text[i] == '}' && --depth == 0) return i; }
        return text.Length;
    }
    private void Replace(int start, int length, string value) => Active.Text = Active.Text[..start] + value + Active.Text[(start + length)..];
    private (int Start,int Length) Line(int? requested = null)
    {
        var starts = new List<int> { 0 }; for(var i=0;i<Active.Text.Length;i++) if(Active.Text[i]=='\n') starts.Add(i+1);
        var n = requested ?? starts.FindLastIndex(x=>x<=Position);
        n = Math.Clamp(n,0,starts.Count-1); var end=n+1<starts.Count?starts[n+1]:Active.Text.Length;
        return (starts[n],end-starts[n]);
    }
    private string LineText() { var line=Line();return Active.Text.Substring(line.Start,line.Length).TrimEnd('\r','\n'); }
    private (int Start,int End,string Name) Procedure(string name = "")
    {
        var vb=Regex.Matches(Active.Text,@"(?im)^\s*(?:(?:Public|Private|Friend|Protected|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\b[^\r\n]*");
        foreach(Match m in vb)
        {
            var end=Regex.Match(Active.Text[(m.Index+m.Length)..],@"(?im)^\s*End\s+(?:Sub|Function|Property)\b[^\r\n]*");
            var finish=end.Success?m.Index+m.Length+end.Index+end.Length:Active.Text.Length;
            if(name.Length==0 ? Position>=m.Index&&Position<=finish : m.Groups[1].Value.Equals(name,StringComparison.OrdinalIgnoreCase)) return(m.Index,finish,m.Groups[1].Value);
        }
        if (SourceNavigation.Member(Active.Text,Position,name) is {} syntaxMember) return syntaxMember;
        var cs=Regex.Matches(Active.Text,@"(?m)^\s*(?:(?:public|private|protected|internal|static|async|virtual|override|sealed)\s+)+[\w<>\[\],?]+\s+(\w+)\s*\([^;{}]*\)\s*\{");
        foreach(Match m in cs)
        {
            var depth=1;var finish=m.Index+m.Length;for(;finish<Active.Text.Length;finish++){if(Active.Text[finish]=='{')depth++;if(Active.Text[finish]=='}'&&--depth==0){finish++;break;}}
            if(name.Length==0 ? Position>=m.Index&&Position<=finish : m.Groups[1].Value.Equals(name,StringComparison.OrdinalIgnoreCase))return(m.Index,finish,m.Groups[1].Value);
        }
        throw new InvalidOperationException("Procedure not found: "+name);
    }
    public async Task<bool> ExecuteAsync(string command,string parameters,IDictionary<string,string> vars,CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();parameters=parameters.Trim().TrimEnd('~');
        string Value(string s)=>vars.TryGetValue(s,out var value)?value:s;
        var words=parameters.Split(' ',StringSplitOptions.RemoveEmptyEntries);int Number(string value,int fallback=0)=>int.TryParse(Value(value),out var n)?n:fallback;
        var equal=parameters.IndexOf('=');var result=equal<0?parameters:parameters[..equal].Trim();var expression=equal<0?"":parameters[(equal+1)..];
        switch(command)
        {
            case "GOTOPROJECT":
                var projectName=words.FirstOrDefault()??throw new FormatException("A project name is required.");
                if(projectName.Equals(ProjectName,StringComparison.OrdinalIgnoreCase)&&ProjectPath!=null)return true;
                var projectRoot=WorkspaceRoot.Length>0?WorkspaceRoot:System.IO.Path.GetDirectoryName(ActivePath)??"";
                var newProject=System.IO.Path.Combine(projectRoot,projectName,projectName+".csproj");
                var projectText="";try{projectText=FileReader?.Invoke(newProject)??"";}catch(FileNotFoundException){}catch(DirectoryNotFoundException){}
                var generatedProject=projectText.Length==0?"<Project Sdk=\"Microsoft.NET.Sdk\"><PropertyGroup><TargetFramework>net10.0</TargetFramework>"+(words.ElementAtOrDefault(1)?.Equals("Exe",StringComparison.OrdinalIgnoreCase)==true?"<OutputType>Exe</OutputType>":"")+"</PropertyGroup></Project>":projectText;
                AttachProject(newProject,generatedProject);Documents[ProjectPath!]=new(ProjectPath!,projectText){Text=generatedProject};WorkspaceRoot=projectRoot;return true;
            case "GOTOMODULE":case "GOTOCLASS":case "GOTOFORM":
                var name=words.FirstOrDefault()??throw new FormatException("A module name is required.");
                var existing=Documents.Keys.FirstOrDefault(x=>System.IO.Path.GetFileNameWithoutExtension(x).Equals(name,StringComparison.OrdinalIgnoreCase));
                if(existing==null)
                {
                    var extension=command=="GOTOFORM"?".frm":ActivePath.EndsWith(".cs",StringComparison.OrdinalIgnoreCase)?".cs":ActivePath.EndsWith(".vb",StringComparison.OrdinalIgnoreCase)?".vb":command=="GOTOCLASS"?".cls":".bas";
                    existing=System.IO.Path.Combine(ProjectPath!=null?System.IO.Path.GetDirectoryName(ProjectPath)!:System.IO.Path.GetDirectoryName(ActivePath)??"",name+extension);
                    var content=command=="GOTOFORM"?$"VERSION 5.00\r\nBegin VB.Form {name}\r\nEnd\r\nAttribute VB_Name = \"{name}\"\r\n":command=="GOTOCLASS"?$"public class {name}\r\n{{\r\n}}\r\n":"";
                    if(command!="GOTOFORM" && extension==".vb")content=$"Public {(command=="GOTOCLASS"?"Class":"Module")} {name}\r\n\r\nEnd {(command=="GOTOCLASS"?"Class":"Module")}\r\n";
                    else if(extension is ".cls" or ".bas")content=(extension==".cls"?"VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1\r\nEND\r\n":"")+$"Attribute VB_Name = \"{name}\"\r\nOption Explicit\r\n";
                    var original = "";
                    if (FileReader != null) { try { original = content = FileReader(existing); } catch (FileNotFoundException) { } catch (DirectoryNotFoundException) { } }
                    Documents.Add(existing,new(existing,original) { Text = content });
                }
                ActivePath=existing;Position=0;SelectionLength=0;LoadDesigner();return true;
            case "GOTODECLARATIONS":case "GOTODECLARTIONS":case "GOTODEC": Position=DeclarationPoint(Active.Text);SelectionLength=0;return true;
            case "GOTOENDOFFILE":case "GOTOENDOFMODULE":case "GOTOEND": Position=BottomPoint(Active.Text);SelectionLength=0;return true;
            case "ABSLINE": Position=Line(Math.Max(0,Number(parameters,1)-1)).Start;SelectionLength=0;return true;
            case "LINEOFFSET":case "OFFSET":var current=Active.Text[..Position].Count(c=>c=='\n');Position=Line(current+Number(parameters)).Start;SelectionLength=0;return true;
            case "FIRSTSELECTIONLINE":Position=Line().Start;SelectionLength=0;return true;
            case "LASTSELECTIONLINE":Position=Math.Min(Active.Text.Length,Position+SelectionLength);Position=Line().Start;SelectionLength=0;return true;
            case "GETTEXTSELECTION":case "GETTEXT":case "GETSELECTION":vars[parameters.Length==0?"TextSelection":parameters]=Active.Text.Substring(Position,SelectionLength);return true;
            case "DELETESELECTION":Replace(Position,SelectionLength,"");SelectionLength=0;return true;
            case "READLINE":vars[parameters.Length==0?"CurrLine":parameters]=LineText();return true;
            case "NEXTLINE":if(openFileLines!=null){if(openFileLines.Count==0)throw new EndOfStreamException();vars[parameters.Length==0?"CurrLine":parameters]=openFileLines.Dequeue();return true;}var old=Line();Position=Math.Min(Active.Text.Length,old.Start+old.Length);vars[parameters.Length==0?"CurrLine":parameters]=LineText();return true;
            case "DELETELINE":case "DELETELINES":var first=Line();var end=Line(Active.Text[..Position].Count(c=>c=='\n')+Math.Max(1,Number(parameters,1))-1);Replace(first.Start,end.Start+end.Length-first.Start,"");Position=first.Start;SelectionLength=0;return true;
            case "REPLACELINE":var line=Line();var ending=Active.Text.Substring(line.Start,line.Length).EndsWith('\n')?"\r\n":"";Replace(line.Start,line.Length,Value(parameters)+ending);Position=line.Start;return true;
            case "PREFIXLINE":case "PREFIX":var prefix=Line();Replace(prefix.Start,0,Value(parameters));Position+=Value(parameters).Length;return true;
            case "POSTFIXLINE":case "POSTFIX":var suffix=Line();Replace(suffix.Start+LineText().Length,0,Value(parameters));return true;
            case "REPLACEINMODULE":case "MODULEREPLACE":if(vars.TryGetValue("ToFind",out var find)&&find.Length>0)Active.Text=Active.Text.Replace(find,Var(vars,"ReplaceWith"));Position=Math.Min(Position,Active.Text.Length);return true;
            case "FIND":case "LOCATE":case "SEARCH":case "FINDINPROC":case "PROCFIND":case "PFIND":case "PROCLOCATE":case "PLOCATE":case "PROCSEARCH":case "PSEARCH":
                var index=Active.Text.IndexOf(Value(parameters),Position,StringComparison.OrdinalIgnoreCase);vars["Found"]=index<0?"0":"1";if(index>=0){Position=index;SelectionLength=Value(parameters).Length;}return true;
            case "GOTOPROC":case "GOTOPROCEND":case "PROCTOP":case "PROCEND":
                (int Start,int End,string Name) proc;
                try { proc=Procedure(command is "PROCTOP" or "PROCEND"?"":words.FirstOrDefault()??""); }
                catch (InvalidOperationException) when (command is "GOTOPROC" or "GOTOPROCEND" && words.Length>0)
                {
                    var memberName=words[0];var memberKind=words.ElementAtOrDefault(1)??"Sub";var vb=ActivePath.EndsWith(".vb",StringComparison.OrdinalIgnoreCase)||ActivePath.EndsWith(".bas",StringComparison.OrdinalIgnoreCase)||ActivePath.EndsWith(".cls",StringComparison.OrdinalIgnoreCase)||ActivePath.EndsWith(".frm",StringComparison.OrdinalIgnoreCase);
                    var declaration=vb?$"\r\nPrivate {(memberKind.Equals("Function",StringComparison.OrdinalIgnoreCase)?"Function":memberKind.Equals("Get",StringComparison.OrdinalIgnoreCase)?"Property Get":"Sub")} {memberName}()\r\n\r\nEnd {(memberKind.Equals("Function",StringComparison.OrdinalIgnoreCase)?"Function":memberKind.Equals("Get",StringComparison.OrdinalIgnoreCase)?"Property":"Sub")}\r\n":$"\r\nprivate void {memberName}()\r\n{{\r\n}}\r\n";
                    var at=SourceNavigation.TypeBoundaries(Active.Text,Position)?.Bottom??Active.Text.Length;Replace(at,0,declaration);proc=Procedure(memberName);
                }
                var endPoint=command is "GOTOPROCEND" or "PROCEND";
                Position=endPoint?Math.Max(proc.Start,Active.Text.LastIndexOf('\n',Math.Max(0,proc.End-1))+1):proc.Start;
                var offset=words.LastOrDefault();var shift=offset!=null&&int.TryParse(offset,out var requestedShift)?requestedShift:endPoint?0:1;
                if(!endPoint&&shift==0)shift=1;
                if(!endPoint && ActivePath.EndsWith(".cs",StringComparison.OrdinalIgnoreCase)) {
                    var brace=Active.Text.IndexOf('{',proc.Start,proc.End-proc.Start);
                    if(brace>=0)Position=brace;
                }
                Position=Line(Active.Text[..Position].Count(c=>c=='\n')+shift).Start;
                vars["Proc Name"]=proc.Name;vars["Proc Type"]=words.ElementAtOrDefault(1)??"Proc";
                SelectionLength=0;return true;
            case "DELETEPROC":case "COPYPROC":case "CUTPROC":var member=Procedure(words.FirstOrDefault()??"");Clipboard=Active.Text[member.Start..member.End];if(command!="COPYPROC"){Replace(member.Start,member.End-member.Start,"");Position=member.Start;}return true;
            case "GETCLIPBOARDTEXT":case "GETCLIPBOARD":case "GETCLIP":vars[parameters.Length==0?"Clipboard":parameters]=Effects.Any(e=>e.Kind=="Clipboard")?Clipboard:ClipboardReader?.Invoke()??Clipboard;return true;
            case "SETCLIPBOARDTEXT":case "SETCLIPBOARD":case "SETCLIP":Clipboard=Value(parameters);Effects.Add(new("Clipboard","",Clipboard));return true;
            case "IGNOREREADONLY":case "WATCHREADONLY":case "CLOSE":case "CLOSECODE":case "CLOSEWINDOW":case "HIDE":case "HIDECODE":case "HIDEWINDOW":case "CLOSEALL":case "CLOSEWINDOWS":case "CLOSEALLWINDOWS":return true;
            default:return await ExecuteSupplementAsync(command,parameters,vars,ct);
        }
    }
    public bool TryInline(string command,string parameters,IDictionary<string,string> vars,out string value)
    {
        value=command switch {"MODULELIST"=>string.Join(";",Documents.Keys.Select(System.IO.Path.GetFileNameWithoutExtension)),"PROJECTLIST"=>ProjectName,"MODULECONTENTS" or "CONTENTS" or "MEMBERS"=>Active.Text,_=>""};
        return command is "MODULELIST" or "PROJECTLIST" or "MODULECONTENTS" or "CONTENTS" or "MEMBERS" || TrySupplementInline(command,parameters,vars,out value);
    }
}
