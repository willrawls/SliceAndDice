using System.Text.Json;
using Sandy.Core;

if(args.Length!=2)throw new ArgumentException("Usage: Sandy.Audit library.xml report.json");
var library=Library.Load(args[0]);var rows=new List<object>();var failures=new Dictionary<string,int>();var passed=0;
foreach(var template in library.Templates)
{
    using var deadline=new CancellationTokenSource(TimeSpan.FromMilliseconds(250));
    var editor=new LegacyEditor(Path.GetFullPath("SandyAudit.frm"),"VERSION 5.00\r\nBegin VB.Form Sample\r\n   Begin VB.TextBox txtName\r\n      Text = \"Example\"\r\n   End\r\nEnd\r\nAttribute VB_Name = \"Sample\"\r\nOption Explicit\r\nPrivate Sub Form_Load()\r\nEnd Sub\r\n"){
        FileReader=path=>throw new FileNotFoundException("Audit fixture has no external file.",path), DirectoryReader=_=>"",Clipboard="Example\r\nNext line",AskAsync=(_,fallback,_)=>Task.FromResult<string?>(fallback)
    };
    try
    {
        var result=await new Engine().EvaluateAsync(library,template,prompt:(_,_)=>Task.FromResult<string?>("1"),cancellation:deadline.Token,host:editor);
        rows.Add(new{template.Id,template.Name,Status="Evaluated",Outputs=result.Outputs.Count,Effects=editor.Effects.Count});passed++;
    }
    catch(Exception error)
    {
        var message=error.Message;var kind=error is OperationCanceledException?"CancelledOrTimeLimit":message.Contains("Unsupported SoftCommand:")?"UnsupportedCommand":"NeedsFixtureOrCompatibilityWork";
        failures[kind]=failures.GetValueOrDefault(kind)+1;
        rows.Add(new{template.Id,template.Name,Status=kind,Error=message});
    }
}
var report=new{Note="Synthetic fixture probe only. Successful evaluation does not establish legacy parity, correct output or live host behavior. No effects were committed.",Evaluated=passed,Total=rows.Count,Failures=failures,Templates=rows};
Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(args[1]))!);
File.WriteAllText(args[1],JsonSerializer.Serialize(report,new JsonSerializerOptions{WriteIndented=true}));
Console.WriteLine(JsonSerializer.Serialize(new{Evaluated=passed,Total=rows.Count,Failures=failures}));
