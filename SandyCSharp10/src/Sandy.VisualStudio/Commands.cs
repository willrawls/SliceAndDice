using Microsoft.VisualStudio.Extensibility;
using Microsoft.VisualStudio.Extensibility.Commands;
using Microsoft.VisualStudio.Extensibility.Editor;
using Microsoft.VisualStudio.Extensibility.Shell;
using Sandy.Core;

namespace Sandy.VisualStudio;

[VisualStudioContribution]
public sealed class OpenSandyCommand : Command
{
    public override CommandConfiguration CommandConfiguration => new("%Sandy.Open%")
    {
        Placements = [CommandPlacement.KnownPlacements.ExtensionsMenu]
    };

    public override Task ExecuteCommandAsync(IClientContext context, CancellationToken cancellationToken)
        => Extensibility.Shell().ShowToolWindowAsync<SandyToolWindow>(true, cancellationToken);
}

[VisualStudioContribution]
public sealed class InsertTemplateCommand(SandySession session) : Command
{
    public override CommandConfiguration CommandConfiguration => new("%Sandy.Insert%")
    {
        Placements = [CommandPlacement.KnownPlacements.ExtensionsMenu],
        VisibleWhen = ActivationConstraint.ClientContext(ClientContextKey.Shell.ActiveEditorContentType, ".+")
    };

    public override async Task ExecuteCommandAsync(IClientContext context, CancellationToken cancellationToken)
    {
        if (!await session.Gate.WaitAsync(0, cancellationToken)) return;
        session.SetBusy(true);
        try
        {
            var library = session.Library;
            var template = session.SelectedTemplate;
            if (library == null || template == null)
            {
                await Extensibility.Shell().ShowToolWindowAsync<SandyToolWindow>(true, cancellationToken);
                await Extensibility.Shell().ShowPromptAsync("Select a template in Sandy, return to the code editor, and run Insert Sandy Template.", PromptOptions.OK, cancellationToken);
                return;
            }
            // Capture the editor snapshot before any prompts. Mutations are version-checked by VS.
            using var view = await context.GetActiveTextViewAsync(cancellationToken);
            if (view == null) return;
            var options = session.CreateOptions();
            options.PromptAsync = async (prompt, ct) => await Extensibility.Shell().ShowPromptAsync(
                prompt.Name, new InputPromptOptions { DefaultText = prompt.DefaultValue, Title = template.Name }, ct);
            var result = await new TemplateEngine().ExecuteAsync(library, template, options, cancellationToken);
            // Until language-aware adapters exist, never silently put declarations or file output at the cursor.
            if (result.Outputs.Any(o => o.Destination != OutputDestination.Editor || o.Area != TemplateArea.Cursor))
                throw new NotSupportedException("This template produces declarations, bottom, file, clipboard, or message output. Preview it in Sandy; this first extension build inserts cursor-only output.");
            var text = string.Concat(result.Outputs.Select(o => o.Text));
            if (text.Length == 0) return;
            if (!await Extensibility.Shell().ShowPromptAsync($"Insert {text.Length} characters from '{template.Name}' at the current selection?\n\n{text[..Math.Min(text.Length, 2000)]}", PromptOptions.OKCancel, cancellationToken)) return;
            var mutation = await Extensibility.Editor().EditAsync(batch =>
                view.Document.AsEditable(batch).Replace(view.Selection.Extent, text), cancellationToken);
            session.Status = "Editor edit result: " + mutation;
        }
        catch (OperationCanceledException) { session.Status = "Cancelled. No editor changes were applied."; }
        catch (Exception ex)
        {
            session.Status = ex.Message;
            await Extensibility.Shell().ShowPromptAsync(ex.Message, PromptOptions.ErrorConfirm, cancellationToken);
        }
        finally { session.SetBusy(false); session.Gate.Release(); }
    }
}
