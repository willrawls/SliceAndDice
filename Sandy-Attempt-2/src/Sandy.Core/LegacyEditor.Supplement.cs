using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Sandy.Core;
public sealed partial class LegacyEditor
{
    private readonly Dictionary<string,string[]> itemLists=new(StringComparer.OrdinalIgnoreCase);
    private string directory = "";
    private readonly Dictionary<string,string?> plannedFiles=new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string,byte[]> plannedBinaryFiles=new(StringComparer.OrdinalIgnoreCase);
    public Func<Uri,CancellationToken,Task<byte[]>>? FetchAsync { get; set; }
    private Queue<string>? openFileLines;
    private readonly Dictionary<string,string?> plannedSettings=new(StringComparer.OrdinalIgnoreCase);
    private string? ReadSetting(string name)=>plannedSettings.TryGetValue(name,out var value)?value:RegistryReader?.Invoke(name);
    private string ReadFile(string path)
    {
        var name=PathName(path);if(plannedBinaryFiles.TryGetValue(name,out var bytes))return TextFiles.Decode(bytes);if(plannedFiles.TryGetValue(name,out var text))return text??throw new FileNotFoundException("Planned file was deleted.",name);
        return (FileReader??throw new NotSupportedException("No file reader is available."))(name);
    }
    private void RecordEffect(ExternalEffect effect)
    {
        if(effect.Kind=="FileOverwrite") { if(effect.Data!=null)plannedBinaryFiles[PathName(effect.Target)]=effect.Data;else {plannedBinaryFiles.Remove(PathName(effect.Target));plannedFiles[PathName(effect.Target)]=effect.Value;} }
        else if(effect.Kind=="FileAppend" && (FileReader!=null || plannedFiles.ContainsKey(PathName(effect.Target)))) { string old;try{old=ReadFile(effect.Target);}catch(FileNotFoundException){old="";}catch(DirectoryNotFoundException){old="";} plannedFiles[PathName(effect.Target)]=old+effect.Value; }
        else if(effect.Kind=="DeleteFile")plannedFiles[PathName(effect.Target)]=null;
        Effects.Add(effect);
    }
    private string Var(IDictionary<string,string> vars,string name,string fallback="") => vars.TryGetValue(name,out var value)?value:fallback;
    private async Task<bool> ExecuteSupplementAsync(string command,string parameters,IDictionary<string,string> vars,CancellationToken ct)
    {
        string Value(string s)=>Var(vars,s,s);
        var args=parameters.Split(',').Select(x=>Value(x.Trim())).ToArray();
        string Arg(int i)=>i<args.Length?args[i]:"";
        var equal=parameters.IndexOf('=');var key=equal<0?parameters:parameters[..equal].Trim();var expression=equal<0?"":parameters[(equal+1)..].Trim();
        switch(command)
        {
            case "COMMENTEDPARAMETERS":
                var documentation=MemberDocumentation.Describe(Active.Text,Position);var docArgs=parameters.Split(' ',StringSplitOptions.RemoveEmptyEntries);
                if(docArgs.Length>0) {
                    var itemDelimiter=docArgs.ElementAtOrDefault(1)??"$$";var valueDelimiter=docArgs.ElementAtOrDefault(2)??"||";
                    vars[docArgs[0]]=string.Join(itemDelimiter,documentation.Parameters.Select(p=>p.Name+valueDelimiter+p.Description));vars["Function Returns"]=documentation.Returns;
                }
                else {
                    var prefix=ActivePath.EndsWith(".cs",StringComparison.OrdinalIgnoreCase)?"// ":"' ";
                    var text=string.Join("\r\n",documentation.Parameters.Select(p=>prefix+p.Name.PadRight(30)+p.Description));
                    if(documentation.Parameters.Count==0)text=prefix+"None";
                    if(documentation.Returns.Length>0)text+="\r\n"+prefix+"Returns "+documentation.Returns;
                    Write(CodeArea.Cursor,OutputDestination.Editor,"",text+"\r\n");
                }
                return true;
            case "URL":case "GETWEBPAGE":case "GETPAGE":case "FETCHURL":case "GETURL":
                vars[key]=TextFiles.Decode(await (FetchAsync??throw new NotSupportedException("No HTTP reader is available in this host."))(new Uri(Value(expression)),ct));return true;
            case "GETFILEURL":case "FILEURL":case "GETFILEFROMPAGE":case "GETREMOTEFILE":case "FETCHFILE":
                var download=await (FetchAsync??throw new NotSupportedException("No HTTP reader is available in this host."))(new Uri(Value(expression)),ct);
                RecordEffect(new("FileOverwrite",PathName(Value(key)),Data:download));return true;
            case "CHANGEDRIVE":case "DRIVE":case "CHDRIVE":case "CHANGEDIRECTORY":case "CHANGEDIR":case "CHDIR":case "CD":directory=Arg(0);return true;
            case "FILECONTENTS":case "GETFILE":case "FILETOSOFTVAR":case "FILETOVAR":vars[key]=ReadFile(Value(expression));return true;
            case "OPENFORINPUT":case "OPENFILE":openFileLines=new Queue<string>(ReadFile(Arg(0)).Replace("\r\n","\n").TrimEnd('\n').Split('\n'));return true;
            case "OPENFOROUTPUT":RecordEffect(new("FileOverwrite",PathName(Arg(0)),""));return true;
            case "LINEINPUTFROMFILE":case "READLINEFROMFILE":case "READFILELINE":case "FILEREADLINE":if(openFileLines==null||openFileLines.Count==0)throw new EndOfStreamException();vars[Arg(0)]=openFileLines.Dequeue();return true;
            case "CLOSEFILE":openFileLines=null;return true;
            case "COPYFILE":case "FILECOPY":RecordEffect(new("FileOverwrite",PathName(equal>=0?Value(key):Arg(1)),ReadFile(equal>=0?Value(expression):Arg(0))));return true;
            case "CHECKFORFILE":case "PAUSEFORFILE":case "WAITFORFILE":case "WAITFOR":Effects.Add(new("WaitForFile",PathName(Arg(0))));return true;
            case "PAUSE":case "WAIT":Effects.Add(new("Pause",Arg(0)));return true;
            case "CREATEFILE":case "MAKEFILE":case "MKFILE":RecordEffect(new("FileOverwrite",PathName(Arg(0)),Arg(1)));return true;
            case "APPENDTOFILE":case "APPENDFILE":RecordEffect(new("FileAppend",PathName(Arg(0)),Arg(1)));return true;
            case "VARTOFILE":case "PUTFILE":case "SAVETOFILE":RecordEffect(new("FileOverwrite",PathName(Value(key)),Value(expression)));return true;
            case "APPENDVARTOFILE":case "APPENDVAR":RecordEffect(new("FileAppend",PathName(Value(key)),Value(expression)));return true;
            case "DELETEFILE":case "DELFILE":case "DEL":case "KILL":plannedBinaryFiles.Remove(PathName(Arg(0)));RecordEffect(new("DeleteFile",PathName(Arg(0))));return true;
            case "CREATEFOLDER":case "CREATEDIRECTORY":case "MAKEDIR":case "MKDIR":case "MD":Effects.Add(new("CreateDirectory",PathName(Arg(0))));return true;
            case "DELETEFOLDER":case "DELETEDIRECTORY":case "DELDIR":case "DELETEDIR":Effects.Add(new("DeleteDirectory",PathName(Arg(0))));return true;
            case "RENAME":case "REN":case "RENAMEFILE":Effects.Add(new("Rename",PathName(Arg(0)),PathName(Arg(1))));return true;
            case "RUNPROGRAM":case "SHELL":case "EXECUTE":Effects.Add(new("Process",Arg(0),Arg(1)));vars["SHELLID"]="planned";return true;
            case "ACTIVATE":case "APPACTIVATE":case "ACTIVATEAPPLICATION":case "SWITCHTOAPP":Effects.Add(new("Activate",Arg(0)));return true;
            case "SENDKEYS":case "SENDKEYSWITHPAUSE":case "SENDKEYSANDWAIT":case "SENDKEYSANDPAUSE":case "SENDKEYSWITHWAIT":Effects.Add(new("SendKeys",Arg(0)));return true;
            case "SAVESETTING":case "SAVEREGISTRYSETTING":case "SETREGKEY":plannedSettings[Arg(0)+"/"+Arg(1)]=Arg(2);Effects.Add(new("Setting",Arg(0)+"/"+Arg(1),Arg(2)));return true;
            case "GETSETTING":case "GETREGISTRYSETTING":case "GETREGKEY":vars[Arg(0)]=ReadSetting(Arg(1)+"/"+Arg(2))??Arg(3);return true;
            case "DELETESETTING":case "DELETEREGISTRYSETTING":case "DELETEREGKEY":case "DELREGKEY":plannedSettings[Arg(0)+"/"+Arg(1)]=null;Effects.Add(new("DeleteSetting",Arg(0)+"/"+Arg(1)));return true;
            case "DECIDE":case "YESNO":case "YESNOCANCEL":case "ASKUSER":case "CHOOSE":
                if(AskAsync==null)throw new NotSupportedException("This command needs an interactive prompt host.");
                vars[key]=await AskAsync(Value(expression),command=="CHOOSE"?Var(vars,"Choose Default"):"1",ct)??throw new OperationCanceledException(ct);return true;
            case "AUTOSPACE":vars[key]=Regex.Replace(Value(expression),"([a-z0-9])([A-Z])","$1 $2");return true;
            case "MANGLE":vars[key]=LegacyStrings.Mangle(Value(expression));return true;
            case "REVERSETOKENS":case "FLIPFLOP":
                var flip=Regex.Match(parameters,"^(\\S+)\\s+(\\S+)\\s+\"([^\"]+)\"$");
                if(!flip.Success)throw new FormatException("ReverseTokens needs SourceVariable TargetVariable \"delimiter\".");
                var flipped=new StringBuilder();var input=Var(vars,flip.Groups[1].Value);var separator=flip.Groups[3].Value;
                foreach(var segment in Regex.Split(input,"(?<=\\n)"))
                {
                    var content=segment.TrimEnd('\r','\n');var eol=segment[content.Length..];var at=content.IndexOf(separator,StringComparison.Ordinal);
                    if(at<0){flipped.Append(content).Append(eol);continue;}
                    var left=content[..at];var indent=left[..(left.Length-left.TrimStart().Length)];
                    flipped.Append(indent).Append(content[(at+separator.Length)..]).Append(separator).Append(left.TrimStart()).Append(eol);
                }
                vars[flip.Groups[2].Value]=flipped.ToString();return true;
            case "STRINGLENGTH":case "LEN":case "LENGTH":case "STRLEN":vars[key]=Value(expression).Length.ToString(CultureInfo.InvariantCulture);return true;
            case "PARSE":case "SPLIT":case "TOKENIZE":
                var listName=expression.Length==0?"Default":key;
                var items=Value(expression.Length==0?key:expression).Split(Var(vars,"ItemDelimiter",Var(vars,"Delimiter",";")),StringSplitOptions.None);
                itemLists[listName]=items;vars["ItemCount"]=items.Length.ToString();vars["ItemsLeft"]=items.Length.ToString();vars["Current ItemList"]=listName;
                vars["Next Item in "+listName]="1";vars["Next Item"]="1";vars["Current Item in "+listName]="0";vars["Current Item"]="0";return true;
            case "RESETGLOBALS":vars.Clear();return true;
            case "SETDEFAULT":
                var split=parameters.IndexOf(' ');if(split<0)throw new FormatException("SetDefault needs ControlType Property=Value.");
                var kind=parameters[..split];var pair=parameters[(split+1)..].Split('=',2);if(pair.Length!=2)throw new FormatException("SetDefault needs Property=Value.");
                if(!defaults.TryGetValue(kind,out var props))defaults[kind]=props=new(StringComparer.OrdinalIgnoreCase);props[pair[0]]=pair[1];return true;
            case "RESETDEFAULTS":defaults.Clear();return true;
            case "ADDCONTROL":
                var pieces=parameters.Split(' ',2,StringSplitOptions.RemoveEmptyEntries);if(pieces.Length!=2)throw new FormatException("AddControl requires a control type and name.");
                SelectedControl=pieces[1];Controls[SelectedControl]=defaults.TryGetValue(pieces[0],out var d)?new(d,StringComparer.OrdinalIgnoreCase):new(StringComparer.OrdinalIgnoreCase);
                Controls[SelectedControl]["$Type"]=pieces[0];UpdateDesigner();return true;
            case "SELECTCONTROL":if(!Controls.ContainsKey(parameters))throw new InvalidOperationException("Control not found: "+parameters);SelectedControl=parameters;return true;
            case "SETPROPERTY":if(!Controls.TryGetValue(SelectedControl,out var control))throw new InvalidOperationException("Select a control first.");control[key]=Value(expression);UpdateDesigner();return true;
            case "SETFORMPROPERTY":case "FORMPROPERTY":FormProperties[key]=Value(expression);UpdateDesigner();return true;
            case "PROCATTR":
                var procedure=Procedure();var attribute=parameters.Split(' ',2);if(attribute.Length!=2)throw new FormatException("ProcAttr requires an attribute and value.");
                var attr=attribute[0].ToUpperInvariant() switch{"ID"=>"VB_UserMemId","HIDDEN"=>"VB_MemberFlags","DESC"=>"VB_Description",_=>throw new NotSupportedException("Unknown procedure attribute.")};
                var attrValue=attribute[1].Equals("Default",StringComparison.OrdinalIgnoreCase)?"0":attribute[1];
                Replace(procedure.Start,0,$"Attribute {procedure.Name}.{attr} = {attrValue}\r\n");return true;
            case "ADDFILE":case "INCLUDEFILE":case "INCLUDE":case "ADDFILEREFERENCE":case "ADDFILEREF":AddProjectItem(parameters,"Compile");return true;
            case "ADDREFERENCE":case "ADDCOMPONENT":case "ADDREF":AddProjectItem(parameters,"Reference");return true;
            case "EDITMESSAGE":
                if(AskAsync==null)throw new NotSupportedException("EditMessage needs an interactive host.");
                vars["EditedMessage"]=await AskAsync("Edit message",Value(parameters),ct)??throw new OperationCanceledException(ct);return true;
            default:return false;
        }
    }
    private string PathName(string value)=>System.IO.Path.IsPathRooted(value)?value:System.IO.Path.Combine(directory.Length>0?directory:System.IO.Path.GetDirectoryName(ActivePath)??"",value);
    private static string PropertyLiteral(string value)
    {
        if(value.StartsWith('"') || Regex.IsMatch(value,@"^(?:-?\d+(?:\.\d+)?|&H[0-9A-F]+&?|True|False)$",RegexOptions.IgnoreCase))return value;
        return "\""+value.Replace("\"","\"\"")+"\"";
    }
    private void LoadDesigner()
    {
        Controls.Clear(); FormProperties.Clear(); SelectedControl="";
        if(!ActivePath.EndsWith(".frm",StringComparison.OrdinalIgnoreCase))return;
        var stack=new Stack<Dictionary<string,string>>();
        foreach(var line in Active.Text.Replace("\r\n","\n").Split('\n'))
        {
            var begin=Regex.Match(line,@"^\s*Begin (?<type>\S+) (?<name>\w+)\s*$",RegexOptions.IgnoreCase);
            if(begin.Success)
            {
                var properties=stack.Count==0?FormProperties:new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
                if(stack.Count>0){properties["$Type"]=begin.Groups["type"].Value;properties["$Parent"]=Controls.FirstOrDefault(c=>ReferenceEquals(c.Value,stack.Peek())).Key??"";Controls[begin.Groups["name"].Value]=properties;}
                stack.Push(properties);continue;
            }
            if(line.Trim().Equals("End",StringComparison.OrdinalIgnoreCase)){if(stack.Count>0)stack.Pop();continue;}
            var property=Regex.Match(line,@"^\s*(\w+)\s*=\s*(.*?)\s*$");
            if(stack.Count>0&&property.Success)stack.Peek()[property.Groups[1].Value]=property.Groups[2].Value;
        }
    }
    private void UpdateDesigner()
    {
        if(!ActivePath.EndsWith(".frm",StringComparison.OrdinalIgnoreCase))throw new NotSupportedException("VB6 control definitions require a .frm document. They cannot be applied to a C#/Blazor designer.");
        var root=Regex.Match(Active.Text,@"(?im)^Begin VB.Form \w+[^\r\n]*\r?\n");
        if(!root.Success)throw new InvalidDataException("The document has no VB6 form definition.");
        foreach(var pair in FormProperties)SetProperty(root.Index+root.Length,"   ",pair.Key,pair.Value);
        foreach(var item in Controls)
        {
            var start=Regex.Match(Active.Text,@"(?im)^(?<indent> +)Begin \S+ "+Regex.Escape(item.Key)+@"\s*\r?$");
            if(!start.Success)
            {
                var rootEnd=Regex.Match(Active.Text,@"(?m)^End\r?$");if(!rootEnd.Success)throw new InvalidDataException("The form definition is not closed.");
                var block=new StringBuilder("   Begin "+item.Value["$Type"]+" "+item.Key+"\r\n");
                foreach(var pair in item.Value.Where(x=>!x.Key.StartsWith('$')))block.Append("      ").Append(pair.Key).Append(" = ").Append(PropertyLiteral(pair.Value)).Append("\r\n");
                block.Append("   End\r\n");Replace(rootEnd.Index,0,block.ToString());
            }
            else foreach(var pair in item.Value.Where(x=>!x.Key.StartsWith('$')))SetProperty(start.Index+start.Length,start.Groups["indent"].Value+"   ",pair.Key,pair.Value);
        }
        Position=Math.Min(Position,Active.Text.Length);
    }
    private void SetProperty(int afterBegin,string indent,string key,string value)
    {
        var closing=Regex.Match(Active.Text[afterBegin..],@"(?m)^"+Regex.Escape(indent.Length>=3?indent[..^3]:"")+@"End\r?$");
        var end=closing.Success?afterBegin+closing.Index:Active.Text.Length;
        var pattern=@"(?m)^"+Regex.Escape(indent+key)+@"\s*=.*\r?$";
        var existing=Regex.Match(Active.Text[afterBegin..end],pattern);
        var line=indent+key+" = "+PropertyLiteral(value);
        if(existing.Success)Replace(afterBegin+existing.Index,existing.Length,line);
        else Replace(afterBegin,0,"\r\n"+line+"\r\n");
    }
    internal IEnumerable<string> ExpandControlBlocks(IReadOnlyList<string> lines,bool byFrame)
    {
        var blocks=new Dictionary<string,StringBuilder>(StringComparer.OrdinalIgnoreCase);string type="";
        foreach(var line in lines)
        {
            if(Regex.IsMatch(line,@"^\*\*[\w.]+\s*$")){type=line[2..].Trim();blocks.TryAdd(type,new());}
            else if(type.Length>0)blocks[type].Append(line).Append("\r\n");
        }
        IEnumerable<KeyValuePair<string,Dictionary<string,string>>> controls=Controls;
        if(byFrame)controls=Controls.Where(c=>c.Value["$Type"].Split('.').Last().Equals("Frame",StringComparison.OrdinalIgnoreCase))
            .SelectMany(frame=>new[]{frame}.Concat(Controls.Where(c=>c.Value.GetValueOrDefault("$Parent")==frame.Key)));
        foreach(var control in controls)
        {
            var kind=control.Value["$Type"].Split('.').Last();
            if(!blocks.TryGetValue(kind,out var body)&&!blocks.TryGetValue(control.Value["$Type"],out body))continue;
            yield return Regex.Replace(body.ToString(),@"\*\*(?:Control\.)?(\w+)\*\*",match=>{
                var property=match.Groups[1].Value;
                if(property.Equals("Name",StringComparison.OrdinalIgnoreCase))return control.Key;
                if(property.Equals("sName",StringComparison.OrdinalIgnoreCase))return control.Key.Length>3?control.Key[3..]:"";
                return control.Value.GetValueOrDefault(property,"").Trim('"');
            },RegexOptions.IgnoreCase);
        }
    }
    private bool TrySupplementInline(string command,string parameters,IDictionary<string,string> vars,out string value)
    {
        value="";switch(command)
        {
            case "NEXTITEM":case "ITEMTOPROCESS":case "NEXTITEMTOPROCESS":case "NEXTTOKEN":
                var listName=parameters.Length==0?Var(vars,"Current ItemList"):Var(vars,parameters,parameters);
                if(!itemLists.TryGetValue(listName,out var items))return true;
                var index=int.TryParse(Var(vars,"Next Item in "+listName),out var n)?n:0;if(index<1||index>items.Length)return true;
                value=items[index-1];vars["Current Item"]=vars["Current Item in "+listName]=index.ToString();
                vars["Next Item"]=vars["Next Item in "+listName]=index==items.Length?"0":(index+1).ToString();vars["ItemsLeft"]=(items.Length-index).ToString();return true;
            case "FILECONTENTS":case "GETFILE":case "INCLUDEFILE":value=ReadFile(Var(vars,parameters,parameters));return true;
            case "GETSETTING":case "GETREGISTRYSETTING":case "GETREGKEY":var parts=parameters.Split(',').Select(p=>Var(vars,p.Trim(),p.Trim())).ToArray();value=ReadSetting(string.Join("/",parts.Take(2)))??parts.ElementAtOrDefault(2)??"";return true;
            case "TESTEOF":case "EOF":case "ENDOFFILE":case "ISENDOFFILE":case "TESTENDOFFILE":value=openFileLines?.Count==0?"True":"False";return true;
            case "FILEEXISTS":case "EXISTENCE":case "DOESFILEEXIST":try{ReadFile(Var(vars,parameters,parameters));value="1";}catch(FileNotFoundException){value="0";}catch(DirectoryNotFoundException){value="0";}return true;
            case "LISTFILES":case "FILES":case "ENUMFILES":case "LISTDIRS":case "LISTDIRECTORIES":case "DIRECTORIES":case "DIRS":case "ENUMDIRS":case "MATCHFILE":case "DIR":case "FINDFIRST":case "FIRSTFILE":value=DirectoryReader?.Invoke(parameters)??"";return true;
            case "MANGLE":value=LegacyStrings.Mangle(Var(vars,parameters,parameters));return true;
            default:return false;
        }
    }
}
