using System.Collections.ObjectModel;
using System.Runtime.Serialization;
using System.Text.Json;
using Microsoft.VisualStudio.Extensibility.UI;
using Sandy.Core;

namespace Sandy.VisualStudio;

[DataContract]
public sealed class SandySession : NotifyPropertyChangedObject
{
    private string libraryPath = Path.Combine(Path.GetDirectoryName(typeof(SandySession).Assembly.Location)!, "Libraries", "CSharp.xml");
    private string filter = "";
    private string variables = "";
    private string preview = "";
    private string status = "Load a Category.xml library or the included C# examples.";
    private TemplateItem? selected;
    private bool isBusy;
    private string cursorCode = "";
    private string declarationsCode = "";
    private string bottomCode = "";
    private string fileCode = "";
    private string filename = "";
    private readonly Dictionary<string, string[]> drafts = new();
    internal SemaphoreSlim Gate { get; } = new(1, 1);
    internal SandyLibrary? Library { get; private set; }
    internal SandyTemplate? SelectedTemplate => Library?.Templates.FirstOrDefault(t => t.Id == Selected?.Id);

    public SandySession()
    {
        LoadCommand = new AsyncCommand(LoadAsync);
        PreviewCommand = new AsyncCommand(PreviewAsync);
        SaveCopyCommand = new AsyncCommand(SaveCopyAsync);
        UpdateTemplateCommand = new AsyncCommand(UpdateTemplateAsync);
    }

    [DataMember] public string LibraryPath { get => libraryPath; set => SetProperty(ref libraryPath, value); }
    [DataMember] public string Filter { get => filter; set { SetProperty(ref filter, value); Refresh(); } }
    [DataMember] public string Variables { get => variables; set => SetProperty(ref variables, value); }
    [DataMember] public string Preview { get => preview; set => SetProperty(ref preview, value); }
    [DataMember] public string Status { get => status; set => SetProperty(ref status, value); }
    [DataMember] public bool IsBusy { get => isBusy; private set => SetProperty(ref isBusy, value); }
    internal void SetBusy(bool value) => IsBusy = value;
    [DataMember] public string CursorCode { get => cursorCode; set => SetProperty(ref cursorCode, value); }
    [DataMember] public string DeclarationsCode { get => declarationsCode; set => SetProperty(ref declarationsCode, value); }
    [DataMember] public string BottomCode { get => bottomCode; set => SetProperty(ref bottomCode, value); }
    [DataMember] public string FileCode { get => fileCode; set => SetProperty(ref fileCode, value); }
    [DataMember] public string Filename { get => filename; set => SetProperty(ref filename, value); }
    [DataMember] public ObservableCollection<TemplateItem> Templates { get; } = new();
    [DataMember] public TemplateItem? Selected
    {
        get => selected;
        set
        {
            if (selected != null)
                drafts[selected.Id] = [CursorCode, DeclarationsCode, BottomCode, FileCode, Filename];
            SetProperty(ref selected, value);
            CursorCode = SelectedTemplate?.GetCode(TemplateArea.Cursor) ?? "";
            DeclarationsCode = SelectedTemplate?.GetCode(TemplateArea.Declarations) ?? "";
            BottomCode = SelectedTemplate?.GetCode(TemplateArea.Bottom) ?? "";
            FileCode = SelectedTemplate?.GetCode(TemplateArea.File) ?? "";
            Filename = SelectedTemplate?.Get("Filename") ?? "";
            if (value != null && drafts.TryGetValue(value.Id, out var draft))
            {
                CursorCode = draft[0]; DeclarationsCode = draft[1]; BottomCode = draft[2]; FileCode = draft[3]; Filename = draft[4];
            }
            Preview = SelectedTemplate == null ? "" : string.Join("\n\n", Enum.GetValues<TemplateArea>()
                .Select(a => $"[{a}]\n{SelectedTemplate.GetCode(a)}"));
        }
    }
    [DataMember] public IAsyncCommand LoadCommand { get; }
    [DataMember] public IAsyncCommand PreviewCommand { get; }
    [DataMember] public IAsyncCommand SaveCopyCommand { get; }
    [DataMember] public IAsyncCommand UpdateTemplateCommand { get; }

    internal ExecutionOptions CreateOptions()
    {
        var options = new ExecutionOptions();
        foreach (var line in Variables.Replace("\r\n", "\n").Split('\n').Where(l => !string.IsNullOrWhiteSpace(l)))
        {
            var equals = line.IndexOf('=');
            if (equals < 1) throw new FormatException("SoftVariables use one Name=Value assignment per line.");
            options.Variables[line[..equals].Trim()] = line[(equals + 1)..];
        }
        return options;
    }

    private async Task LoadAsync(object? parameter, CancellationToken ct)
    {
        if (!await Gate.WaitAsync(0, ct)) return;
        IsBusy = true;
        try
        {
            var path = LibraryPath;
            var loaded = await Task.Run(() => LibraryXml.Load(path), ct);
            Selected = null;
            drafts.Clear();
            Library = loaded;
            Refresh();
            Status = $"Loaded {loaded.Templates.Count()} templates in {loaded.Categories.Count} categories.";
        }
        catch (Exception ex) { Status = ex.Message; }
        finally { IsBusy = false; Gate.Release(); }
    }

    private void Refresh()
    {
        Templates.Clear();
        if (Library == null) return;
        foreach (var template in Library.Templates.Where(t => t.Name.Contains(Filter, StringComparison.OrdinalIgnoreCase)).OrderBy(t => t.Name))
            Templates.Add(new TemplateItem { Id = template.Id, Name = template.Name });
    }

    private async Task PreviewAsync(object? parameter, CancellationToken ct)
    {
        if (!await Gate.WaitAsync(0, ct)) return;
        IsBusy = true;
        try
        {
            var library = Library ?? throw new InvalidOperationException("Load a library first.");
            var template = SelectedTemplate ?? throw new InvalidOperationException("Select a template first.");
            var options = CreateOptions();
            // For preview, use template defaults. Insertion prompts interactively through the shell.
            options.PromptAsync = (p, _) => Task.FromResult<string?>(p.DefaultValue);
            var result = await Task.Run(() => new TemplateEngine().ExecuteAsync(library, template, options, ct), ct);
            Preview = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true, Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() } });
            Status = "Preview only. For insertion, focus the code editor and choose Extensions > Insert Sandy Template.";
        }
        catch (Exception ex) { Status = ex.Message; Preview = ex.Message; }
        finally { IsBusy = false; Gate.Release(); }
    }

    private async Task SaveCopyAsync(object? parameter, CancellationToken ct)
    {
        if (!await Gate.WaitAsync(0, ct)) return;
        IsBusy = true;
        try
        {
            var library = Library ?? throw new InvalidOperationException("Load a library first.");
            var directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SandyCSharp10", "Libraries");
            Directory.CreateDirectory(directory);
            var path = Path.Combine(directory, $"Library-{Guid.NewGuid():N}.xml");
            LibraryXml.Save(library, path);
            LibraryPath = path;
            Status = "Saved a library copy: " + path;
        }
        catch (Exception ex) { Status = ex.Message; }
        finally { IsBusy = false; Gate.Release(); }
    }

    private async Task UpdateTemplateAsync(object? parameter, CancellationToken ct)
    {
        if (!await Gate.WaitAsync(0, ct)) return;
        IsBusy = true;
        try
        {
            var template = SelectedTemplate ?? throw new InvalidOperationException("Select a template first.");
            TemplateEditor.UpdateCode(template, new Dictionary<TemplateArea, string>
            {
                [TemplateArea.Cursor] = CursorCode, [TemplateArea.Declarations] = DeclarationsCode,
                [TemplateArea.Bottom] = BottomCode, [TemplateArea.File] = FileCode
            }, Filename, DateTime.UtcNow);
            Status = "Updated in memory. Save a library copy to keep your changes. Original library files are preserved.";
        }
        catch (Exception ex) { Status = ex.Message; }
        finally { IsBusy = false; Gate.Release(); }
    }
}

[DataContract]
public sealed class TemplateItem
{
    [DataMember] public string Id { get; set; } = "";
    [DataMember] public string Name { get; set; } = "";
}
