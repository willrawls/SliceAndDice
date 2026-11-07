using Sandy.Core;

namespace Sandy.Core.Tests;

public class EditingAndSafetyTests
{
    [Fact]
    public void SuiteLoadsTheRequestedCoreTarget()
    {
        var framework = System.Reflection.CustomAttributeExtensions.GetCustomAttribute<System.Runtime.Versioning.TargetFrameworkAttribute>(typeof(TemplateEngine).Assembly)!.FrameworkName;
#if STANDARD21_CORE
        Assert.Equal(".NETStandard,Version=v2.1", framework);
#elif STANDARD_CORE
        Assert.Equal(".NETStandard,Version=v2.0", framework);
#else
        Assert.Equal(".NETCoreApp,Version=v10.0", framework);
#endif
    }

    [Fact]
    public void UpdatesPreserveIdentityAndUnknownFieldsIncrementRevisionAndHonorLocks()
    {
        var (_, template) = EngineTests.Library("before");
        template.Fields["Future"] = "retain";
        template.Fields["RevisionCount"] = "7";
        var change = new Dictionary<TemplateArea, string> { [TemplateArea.Cursor] = "after" };
        TemplateEditor.UpdateCode(template, change, "", new DateTime(2026, 11, 7, 0, 0, 0, DateTimeKind.Utc));
        Assert.Equal("1", template.Id); Assert.Equal("retain", template.Get("Future"));
        Assert.Equal("8", template.Get("RevisionCount")); Assert.Equal("after", template.GetCode(TemplateArea.Cursor));
        TemplateEditor.UpdateCode(template, change, "", DateTime.UtcNow);
        Assert.Equal("8", template.Get("RevisionCount"));
        template.Fields["Locked"] = "-1";
        Assert.Throws<InvalidOperationException>(() => TemplateEditor.UpdateCode(template, new Dictionary<TemplateArea, string> { [TemplateArea.Cursor] = "blocked" }, "", DateTime.UtcNow));
        Assert.Equal("after", template.GetCode(TemplateArea.Cursor));
    }

    [Fact]
    public async Task HistoricalAppendMarkersReplacementAndIncrementAmountWorkTogether()
    {
        var (library, template) = EngineTests.Library("~~Set X=William M Rawls\n~~Set ToFind= M \n~~Set ReplaceWith=$\n~~Replace Y=X\n~~Append Y=EOL\n~~Append Y=$SP$done$9$\n~~Increment Count 5\n~~Decrement Count 2\n%%Y%% %%Count%%");
        var result = await new TemplateEngine().ExecuteAsync(library, template);
        Assert.Equal("William$Rawls\r\n done\t 3\r\n", result.Outputs.Single().Text);
        Assert.Equal("\r\n", result.Variables["EOL"]);
    }

    [Fact]
    public async Task ExplicitCommandReferencesAndScriptRegionsAreProcessed()
    {
        var (library, template) = EngineTests.Library("~##~BeginScript\nSet Name=World\n_i1_If 1\n!!Hello %%Name%%\n_i1_EndIf\n~##~EndScript\n~##~BeginComments\nignored\n~##~EndComments\n");
        var result = await new TemplateEngine().ExecuteAsync(library, template);
        Assert.Equal("Hello World\r\n", result.Outputs.Single().Text);
    }

    [Fact]
    public async Task IncludeCyclesAndOutputExplosionHaveUsefulDiagnostics()
    {
        var (library, template) = EngineTests.Library("~##~Include 2,Main");
        var exception = await Assert.ThrowsAsync<SandyExecutionException>(() => new TemplateEngine().ExecuteAsync(library, template));
        Assert.Contains("Cyclic", exception.Message);
        template.Fields["memoCodeAtCursor"] = "~~String X 500 \"a\"\n%%X%%";
        await Assert.ThrowsAsync<SandyExecutionException>(() => new TemplateEngine().ExecuteAsync(library, template, new ExecutionOptions { MaxOutputCharacters = 100 }));
    }

    [Fact]
    public async Task ManagedSandalCanProvideBothCommandAndInlineFunction()
    {
        var (library, template) = EngineTests.Library("~~Example Set\n%%Example::Hello%%");
        var options = new ExecutionOptions(); options.Providers.Add(new ExampleProvider());
        var result = await new TemplateEngine().ExecuteAsync(library, template, options);
        Assert.Equal("Hello Sandy\r\n", result.Outputs.Single().Text);
    }

    private sealed class ExampleProvider : ISandyCommandProvider
    {
        public bool TryExecute(string command, string parameters, IDictionary<string, string> variables)
        {
            if (command != "EXAMPLE") return false;
            variables["PluginValue"] = "Sandy";
            return true;
        }
        public bool TryEvaluateInline(string command, string parameters, IDictionary<string, string> variables, out string result)
        {
            result = command == "EXAMPLE" ? parameters + " " + variables["PluginValue"] : "";
            return command == "EXAMPLE";
        }
    }
}
