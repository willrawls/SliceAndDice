using Microsoft.Extensions.DependencyInjection;
using Microsoft.VisualStudio.Extensibility;
using Microsoft.VisualStudio.Extensibility.Commands;
using Microsoft.VisualStudio.Extensibility.Editor;
using Microsoft.VisualStudio.Extensibility.Shell;
using Microsoft.VisualStudio.Extensibility.ToolWindows;
using Microsoft.VisualStudio.Extensibility.UI;
using Microsoft.VisualStudio.RpcContracts.RemoteUI;

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
    public override CommandConfiguration CommandConfiguration => new("%Sandy2.Open%") { Placements = [CommandPlacement.KnownPlacements.ExtensionsMenu] };
    public override Task ExecuteCommandAsync(IClientContext context, CancellationToken cancellationToken) => Extensibility.Shell().ShowToolWindowAsync<SandyWindow>(true, cancellationToken);
}

[VisualStudioContribution]
public sealed class InsertCommand(Workspace workspace) : Command
{
    public override CommandConfiguration CommandConfiguration => new("%Sandy2.Insert%")
    {
        Placements = [CommandPlacement.KnownPlacements.ExtensionsMenu],
        VisibleWhen = ActivationConstraint.ClientContext(ClientContextKey.Shell.ActiveEditorContentType, ".+")
    };
    public override async Task ExecuteCommandAsync(IClientContext context, CancellationToken cancellationToken)
    {
        if (!await workspace.Gate.WaitAsync(0, cancellationToken)) return;
        try
        {
            var template = workspace.CurrentTemplate ?? throw new InvalidOperationException("Load a library and select a template in Sandy first.");
            using var view = await context.GetActiveTextViewAsync(cancellationToken);
            if (view == null) return;
            var result = await workspace.EvaluateAsync((p, ct) => Extensibility.Shell().ShowPromptAsync(p.Name, new InputPromptOptions { DefaultText = p.Default, Title = template.Name }, ct), cancellationToken);
            if (result.Outputs.Any(x => x.Area != Core.CodeArea.Cursor)) throw new NotSupportedException("This template has declarations, bottom or file output. Inspect those outputs in Preview; this extension currently inserts cursor-only templates.");
            var text = string.Concat(result.Outputs.Select(x => x.Text));
            if (text.Length == 0) return;
            if (!await Extensibility.Shell().ShowPromptAsync($"Insert {text.Length} characters from {template.Name}?\n\n{text[..Math.Min(text.Length, 1800)]}", PromptOptions.OKCancel, cancellationToken)) return;
            var outcome = await Extensibility.Editor().EditAsync(edit => view.Document.AsEditable(edit).Replace(view.Selection.Extent, text), cancellationToken);
            workspace.Status = "Editor result: " + outcome;
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
