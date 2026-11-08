using Microsoft.Extensions.DependencyInjection;
using Microsoft.VisualStudio.Extensibility;
using Microsoft.VisualStudio.Extensibility.Commands;
using Microsoft.VisualStudio.Extensibility.Editor;
using Microsoft.VisualStudio.Extensibility.Shell;
using Microsoft.VisualStudio.Extensibility.ToolWindows;
using Microsoft.VisualStudio.Extensibility.UI;
using Microsoft.VisualStudio.RpcContracts.RemoteUI;
using Microsoft.VisualStudio.RpcContracts.Documents;

namespace Sandy.VisualStudio;

[VisualStudioContribution]
public sealed class SandyExtension : Microsoft.VisualStudio.Extensibility.Extension
{
    public override ExtensionConfiguration ExtensionConfiguration => new()
    {
        Metadata = new("WilliamRawls.Sandy.Attempt2", ExtensionAssemblyVersion, "William Rawls", "Sandy — Attempt 2", "Sandy templates, desktop forms and Visual Studio insertion") { DotnetTargetVersions = ["net10.0"] }
    };
    protected override void InitializeServices(IServiceCollection services) { base.InitializeServices(services); services.AddSingleton<Workspace>(); }
}

[VisualStudioContribution]
public sealed class OpenCommand : Command
{
    public override CommandConfiguration CommandConfiguration => new("%Sandy2.Open%") { Placements = [CommandPlacement.KnownPlacements.ExtensionsMenu],Shortcuts=[new(ModifierKey.ControlShiftLeftAlt,Key.S)] };
    public override Task ExecuteCommandAsync(IClientContext context, CancellationToken cancellationToken) => Extensibility.Shell().ShowToolWindowAsync<SandyWindow>(true, cancellationToken);
}

[VisualStudioContribution]
public sealed class InsertCommand(Workspace workspace) : Command
{
    public override CommandConfiguration CommandConfiguration => new("%Sandy2.Insert%")
    {
        Placements = [CommandPlacement.KnownPlacements.ExtensionsMenu,CommandPlacement.VsctParent(new Guid("D309F791-903F-11D0-9EFC-00A0C911004F"),0x01C0,0x0600)],
        Shortcuts=[new(ModifierKey.ControlShiftLeftAlt,Key.I)]
    };
    public override async Task ExecuteCommandAsync(IClientContext context, CancellationToken cancellationToken)
    {
        if (!await workspace.Gate.WaitAsync(0, cancellationToken)) return;
        try
        {
            var template = workspace.CurrentTemplate ?? throw new InvalidOperationException("Load a library and select a template in Sandy first.");
            using var view = await context.GetActiveTextViewAsync(cancellationToken);
            if (view == null) return;
            var documentPath = view.Document.Uri.IsFile ? view.Document.Uri.LocalPath : throw new InvalidOperationException("Save the document before applying Sandy templates.");
            var original = view.Document.Text.CopyToString();
            var snapshots = new Dictionary<string,ITextDocumentSnapshot>(StringComparer.OrdinalIgnoreCase) { [documentPath] = view.Document };
            foreach(var open in await Extensibility.Documents().GetOpenDocumentsAsync(cancellationToken))
            {
                if(!open.Moniker.IsFile || snapshots.ContainsKey(open.Moniker.LocalPath))continue;
                if(Path.GetExtension(open.Moniker.LocalPath).ToLowerInvariant() is not (".cs" or ".vb" or ".bas" or ".cls" or ".frm" or ".txt" or ".razor" or ".xml" or ".json" or ".csproj" or ".vbproj"))continue;
                var snapshot = await open.AsTextDocumentAsync(Extensibility,cancellationToken);
                if(snapshot != null) snapshots.Add(open.Moniker.LocalPath,snapshot);
            }
            var editor = new Core.LegacyEditor(documentPath, original, view.Selection.Extent.Start.Offset, view.Selection.Extent.Length)
            {
                FileReader = Core.TextFiles.Read,
                RegistryReader = Sandy.Windows.WindowsEffects.ReadSetting,
                ClipboardReader = Sandy.Windows.WindowsEffects.ReadClipboard,
                FetchAsync = Core.WebContent.FetchAsync,
                PathReader = Core.LegacyFileSystem.Enumerate,
                DirectoryReader = pattern => string.Join(";", Directory.GetFileSystemEntries(Path.GetDirectoryName(pattern) is {Length:>0} dir ? dir : Path.GetDirectoryName(documentPath)!, Path.GetFileName(pattern))),
                AskAsync = async (question, fallback, ct) => await Extensibility.Shell().ShowPromptAsync(question, new InputPromptOptions { DefaultText = fallback, Title = template.Name }, ct)
            };
            if(Core.LegacyEditor.FindProject(documentPath) is {} project)editor.AttachProject(project,snapshots.TryGetValue(project,out var projectSnapshot)?projectSnapshot.Text.CopyToString():Core.TextFiles.Read(project));
            foreach(var snapshot in snapshots.Where(s=>!s.Key.Equals(documentPath,StringComparison.OrdinalIgnoreCase)))
                editor.Documents.TryAdd(snapshot.Key,new Core.PlannedDocument(snapshot.Key,snapshot.Value.Text.CopyToString()));
            var result = await workspace.EvaluateAsync((p, ct) => Extensibility.Shell().ShowPromptAsync(p.Name+(p.Choices==null?"":"\n"+string.Join("\n",p.Choices.Select((choice,index)=>(index+1)+". "+choice))), new InputPromptOptions { DefaultText = p.Default, Title = template.Name }, ct), cancellationToken, editor);
            Core.Library? revisedLibrary=null;string? libraryOutput=null;
            if(result.LibraryMutation is {} libraryMutation)
            {
                revisedLibrary=workspace.PrepareMutation(libraryMutation);
                if(libraryMutation.HasChanges)
                {
                    libraryOutput=await Extensibility.Shell().ShowPromptAsync("Save the edited Sandy library as a new XML working copy",new InputPromptOptions{DefaultText=Path.Combine(editor.WorkspaceRoot.Length>0?editor.WorkspaceRoot:Path.GetDirectoryName(documentPath)!,"Sandy-library-"+Guid.NewGuid().ToString("N")+".xml")},cancellationToken);
                    if(libraryOutput==null)return;
                    libraryOutput=Path.GetFullPath(libraryOutput);
                    if(File.Exists(libraryOutput)||snapshots.ContainsKey(libraryOutput))throw new InvalidOperationException("Choose a new library filename.");
                    editor.Documents.Add(libraryOutput,new(libraryOutput,""){Text=revisedLibrary.Serialize()});
                }
            }
            Sandy.Windows.WindowsEffects.Validate(editor.Effects);
            var changed = editor.Documents.Values.Where(d => d.Changed).ToArray();
            var files = new Core.FileCommit(editor.WorkspaceRoot.Length>0?editor.WorkspaceRoot:Path.GetDirectoryName(documentPath)!);
            foreach (var document in changed.Where(d => !snapshots.ContainsKey(d.Path))) files.AddDocument(document);
            foreach (var effect in editor.Effects.Where(Sandy.Windows.WindowsEffects.IsFile)) {
                if (snapshots.ContainsKey(Path.GetFullPath(effect.Target))) throw new InvalidOperationException("File output targets an open editor. Use editor insertion for this document.");
                files.AddEffect(effect);
            }
            var summary = string.Join("\n", changed.Select(d => $"{Path.GetFileName(d.Path)}: {d.Original.Length} → {d.Text.Length} characters")) + "\n" + string.Join("\n", editor.Effects.Select(e => e.Kind + ": " + e.Target));
            if (!await Extensibility.Shell().ShowPromptAsync("Apply Sandy template '" + template.Name + "'?\n\n" + summary, PromptOptions.OKCancel, cancellationToken)) return;
            files.Commit();
            try
            {
                var editorChanges = changed.Where(d=>snapshots.ContainsKey(d.Path)).ToArray();
                if (editorChanges.Length > 0)
                {
                    var outcome = await Extensibility.Editor().EditAsync(edit => {
                        foreach(var document in editorChanges) { var snapshot = snapshots[document.Path]; snapshot.AsEditable(edit).Replace(snapshot.Text,document.Text); }
                    }, cancellationToken);
                    if (!outcome.Succeeded) throw new InvalidOperationException("Editor changed while the template was running: " + outcome.Message);
                }
            }
            catch { files.Rollback(); throw; }
            if(revisedLibrary!=null)workspace.AcceptMutation(revisedLibrary,result.LibraryMutation!,libraryOutput);
            workspace.Status = "Applied " + changed.Length + " document plan(s). File backups: " + (files.JournalPath ?? "none");
            var appliedActions=Core.DeferredActions.Apply(editor.Effects.Where(e=>!Sandy.Windows.WindowsEffects.IsFile(e)),Sandy.Windows.WindowsEffects.Apply);
            foreach (var effect in appliedActions.Where(e => e.Kind == "Message")) await Extensibility.Shell().ShowPromptAsync(effect.Value, PromptOptions.OK, cancellationToken);
            foreach(var effect in editor.Effects.Where(e=>e.Kind is "EditorClose" or "EditorCloseAll"))
            {
                var openDocuments=await Extensibility.Documents().GetOpenDocumentsAsync(cancellationToken);
                foreach(var open in openDocuments.Where(d=>effect.Kind=="EditorCloseAll"?d.Moniker.IsFile:d.Moniker.IsFile&&d.Moniker.LocalPath.Equals(effect.Target,StringComparison.OrdinalIgnoreCase)))
                    await open.CloseAsync(SaveDocumentOption.PromptSave,Extensibility,cancellationToken);
            }
        }
        catch (OperationCanceledException) { workspace.Status = "Cancelled."; }
        catch (Exception ex) { workspace.Status = ex.Message; await Extensibility.Shell().ShowPromptAsync(ex.Message, PromptOptions.ErrorConfirm, cancellationToken); }
        finally { workspace.Gate.Release(); }
    }
}

[VisualStudioContribution]
public sealed class SandyWindow(Workspace workspace) : ToolWindow
{
    private SandyView? view;
    public override ToolWindowConfiguration ToolWindowConfiguration => new() { Placement = ToolWindowPlacement.DocumentWell };
    public override Task InitializeAsync(CancellationToken cancellationToken) { Title = "Sandy — Attempt 2"; return Task.CompletedTask; }
    public override Task<IRemoteUserControl> GetContentAsync(CancellationToken cancellationToken) => Task.FromResult<IRemoteUserControl>(view ??= new SandyView(workspace));
}
internal sealed class SandyView(object model) : RemoteUserControl(model);
