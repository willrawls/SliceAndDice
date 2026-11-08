using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Runtime.Serialization;
using Microsoft.VisualStudio.Extensibility.UI;
using Sandy.Core;

namespace Sandy.VisualStudio;

[DataContract]
public sealed class Workspace : NotifyPropertyChangedObject
{
    private string path = Path.Combine(Path.GetDirectoryName(typeof(Workspace).Assembly.Location)!, "Libraries", "SliceAndDice.xml");
    private string filter = "", variables = "", preview = "", status = "Load a Category.xml export. Use the desktop forms to import MDB libraries and save an XML working copy.";
    private TemplateRow? selected;
    private Library? library;
    internal SemaphoreSlim Gate { get; } = new(1, 1);
    internal Template? CurrentTemplate => library?.Templates.SingleOrDefault(x => x["TemplateID"] == Selected?.Id);
    public Workspace()
    {
        LoadCommand = new AsyncCommand(async (_, ct) => await GuardAsync(async () => { var loaded = Library.Load(LibraryPath); library = loaded; Selected = null; Refresh(); Status = $"Loaded {loaded.Templates.Count()} templates."; await Task.CompletedTask; }, ct));
        PreviewCommand = new AsyncCommand(async (_, ct) => await GuardAsync(async () => { var result = await EvaluateAsync(null, ct); Preview = string.Join("\r\n\r\n", result.Outputs.Select(x => $"[{x.Area}] {x.Target}\r\n{x.Text}")); Status = "Preview uses defaults. Insertion prompts for missing values."; }, ct));
        DesktopCommand = new AsyncCommand(async (_, ct) => await GuardAsync(async () =>
        {
            var executable = Path.Combine(Path.GetDirectoryName(typeof(Workspace).Assembly.Location)!, "Desktop", "Sandy.Desktop.exe");
            if (!File.Exists(executable)) throw new FileNotFoundException("The desktop payload is missing from the extension.", executable);
            var start = new ProcessStartInfo(executable) { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden };
            if (File.Exists(LibraryPath)) start.ArgumentList.Add(LibraryPath);
            Process.Start(start); await Task.CompletedTask;
        }, ct));
    }
    [DataMember] public string LibraryPath { get => path; set => SetProperty(ref path, value); }
    [DataMember] public string Filter { get => filter; set { SetProperty(ref filter, value); Refresh(); } }
    [DataMember] public string Variables { get => variables; set => SetProperty(ref variables, value); }
    [DataMember] public string Preview { get => preview; set => SetProperty(ref preview, value); }
    [DataMember] public string Status { get => status; set => SetProperty(ref status, value); }
    [DataMember] public TemplateRow? Selected { get => selected; set => SetProperty(ref selected, value); }
    [DataMember] public ObservableCollection<TemplateRow> Templates { get; } = [];
    [DataMember] public AsyncCommand LoadCommand { get; }
    [DataMember] public AsyncCommand PreviewCommand { get; }
    [DataMember] public AsyncCommand DesktopCommand { get; }
    private void Refresh()
    {
        Templates.Clear(); if (library == null) return;
        foreach (var t in library.Templates.Where(t => t.Name.Contains(Filter, StringComparison.OrdinalIgnoreCase))) Templates.Add(new() { Id = t["TemplateID"], Name = t.Name });
    }
    internal Task<Evaluation> EvaluateAsync(Func<Prompt, CancellationToken, Task<string?>>? prompt, CancellationToken cancellation)
    {
        var template = CurrentTemplate ?? throw new InvalidOperationException("Select a template.");
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in Variables.Replace("\r\n", "\n").Split('\n').Where(x => x.Trim().Length > 0))
        {
            var equal = line.IndexOf('='); if (equal < 1) throw new FormatException("SoftVariables must be Name=Value, one per line.");
            values[line[..equal].Trim()] = line[(equal + 1)..];
        }
        return new Engine().EvaluateAsync(library!, template, values, prompt, cancellation);
    }
    private async Task GuardAsync(Func<Task> action, CancellationToken ct)
    {
        if (!await Gate.WaitAsync(0, ct)) return;
        try { await action(); } catch (OperationCanceledException) { Status = "Cancelled."; } catch (Exception ex) { Status = ex.Message; } finally { Gate.Release(); }
    }
}

[DataContract]
public sealed class TemplateRow
{
    [DataMember] public string Id { get; set; } = "";
    [DataMember] public string Name { get; set; } = "";
}
