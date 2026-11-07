using Microsoft.Extensions.DependencyInjection;
using Microsoft.VisualStudio.Extensibility;

namespace Sandy.VisualStudio;

[VisualStudioContribution]
public sealed class SandyExtension : Extension
{
    public override ExtensionConfiguration ExtensionConfiguration => new()
    {
        Metadata = new("WilliamRawls.SandyCSharp10", ExtensionAssemblyVersion, "William Rawls",
            "Slice and Dice — Sandy", "Programmable templates and SoftCommands for Visual Studio")
        {
            DotnetTargetVersions = ["net10.0"]
        }
    };

    protected override void InitializeServices(IServiceCollection services)
    {
        base.InitializeServices(services);
        services.AddSingleton<SandySession>();
    }
}
