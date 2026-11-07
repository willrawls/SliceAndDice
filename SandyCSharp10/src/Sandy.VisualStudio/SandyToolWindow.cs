using Microsoft.VisualStudio.Extensibility;
using Microsoft.VisualStudio.Extensibility.ToolWindows;
using Microsoft.VisualStudio.Extensibility.UI;
using Microsoft.VisualStudio.RpcContracts.RemoteUI;

namespace Sandy.VisualStudio;

[VisualStudioContribution]
public sealed class SandyToolWindow(SandySession session) : ToolWindow
{
    private SandyControl? control;
    public override ToolWindowConfiguration ToolWindowConfiguration => new() { Placement = ToolWindowPlacement.DocumentWell };
    public override Task InitializeAsync(CancellationToken cancellationToken)
    {
        Title = "Slice and Dice — Sandy";
        control = new SandyControl(session);
        return Task.CompletedTask;
    }
    public override Task<IRemoteUserControl> GetContentAsync(CancellationToken cancellationToken)
        => Task.FromResult<IRemoteUserControl>(control ??= new SandyControl(session));
}

internal sealed class SandyControl(object dataContext) : RemoteUserControl(dataContext);
