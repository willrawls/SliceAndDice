using Sandy.Core;

namespace Sandy.Core.Tests;

public class EngineTests
{
    internal static (SandyLibrary, SandyTemplate) Library(string code)
    {
        var library = new SandyLibrary();
        var category = new SandyCategory(); category.Fields["CategoryID"] = "1"; category.Fields["CategoryName"] = "Tests";
        var template = new SandyTemplate();
        template.Fields["TemplateID"] = "1"; template.Fields["CategoryID"] = "1";
        template.Fields["TemplateName"] = "Tests - Main"; template.Fields["ShortTemplateName"] = "Main";
        template.Fields["memoCodeAtCursor"] = code;
        category.Templates.Add(template); library.Categories.Add(category);
        return (library, template);
    }

    private static async Task<ExecutionResult> Execute(string source, ExecutionOptions? options = null)
    {
        var (library, template) = Library(source);
        return await new TemplateEngine().ExecuteAsync(library, template, options);
    }

    [Fact]
    public async Task PromptsReuseValuesAndResolveIndirectDefaultsAcrossAllFourAreas()
    {
        var (library, template) = Library("%%Name::@Suggested%% %%Name%%\n");
        template.Fields["memoCodeAtTop"] = "%%Name%% declarations";
        template.Fields["memoCodeAtBottom"] = "%%Name%% bottom";
        template.Fields["memoCodeToFile"] = "%%Name%% file";
        template.Fields["Filename"] = "%%Name%%.cs";
        var requests = new List<PromptRequest>();
        var options = new ExecutionOptions(); options.Variables["Suggested"] = "Widget";
        options.PromptAsync = (p, _) => { requests.Add(p); return Task.FromResult<string?>(p.DefaultValue); };
        var result = await new TemplateEngine().ExecuteAsync(library, template, options);
        Assert.Single(requests); Assert.Equal("Widget", requests[0].DefaultValue);
        Assert.Equal(new[] { TemplateArea.Cursor, TemplateArea.Declarations, TemplateArea.Bottom, TemplateArea.File }, result.Outputs.Select(o => o.Area));
        Assert.Equal("Widget Widget\r\n", result.Outputs[0].Text);
        Assert.Equal("Widget.cs", result.Outputs[3].Target);
    }

    [Fact]
    public async Task NestedConditionalsSkipUnreachablePromptsAndLoopsReadUpdatedVariables()
    {
        var result = await Execute("~~Set N=3\n~~StartCoda\n~~If N>1\n~~If N=3\nthree\n~~Else\ntwo\n~~EndIf\n~~Else\none\n~~EndIf\n~~Decrement N\n~~Coda N\n~~If 0\n%%Unreachable%%\n~~EndIf");
        Assert.Equal("three\r\ntwo\r\none\r\n", Assert.Single(result.Outputs).Text);
    }

    [Fact]
    public async Task IncludesDefaultToDeclarationsAndAllowExplicitCursorArea()
    {
        var (library, template) = Library("~##~Include Shared\n~##~Include 2,Shared\n");
        var shared = new SandyTemplate(); shared.Fields["TemplateID"] = "2"; shared.Fields["TemplateName"] = "Tests - Shared";
        shared.Fields["ShortTemplateName"] = "Shared"; shared.Fields["memoCodeAtTop"] = "top\n"; shared.Fields["memoCodeAtCursor"] = "cursor\n";
        library.Categories[0].Templates.Add(shared);
        var result = await new TemplateEngine().ExecuteAsync(library, template);
        Assert.Equal("top\r\ncursor\r\n", Assert.Single(result.Outputs).Text);
    }

    [Fact]
    public async Task VariableOutputFlushesBeforeSubsequentExpansionAndRoutesPreserveOrder()
    {
        var result = await Execute("~~GotoSoftVar Text\nhello\n~~ResumeInsertion\n~~Block Text\n~~GotoClipboard\nclip\n~~GotoNewFile output.cs\nfile\n");
        Assert.Equal("hello\r\n", result.Variables["Text"]);
        Assert.Equal(new[] { OutputDestination.Editor, OutputDestination.Clipboard, OutputDestination.FileOverwrite }, result.Outputs.Select(o => o.Destination));
        Assert.Equal("hello\r\n", result.Outputs[0].Text);
        Assert.Equal("output.cs", result.Outputs[2].Target);
    }

    [Fact]
    public async Task UnknownCommandsFailWithTemplateAreaAndLineInsteadOfBeingInserted()
    {
        var ex = await Assert.ThrowsAsync<SandyExecutionException>(() => Execute("before\n~~AddControl button\n"));
        Assert.Equal(2, ex.Line); Assert.Equal(TemplateArea.Cursor, ex.Area);
        Assert.Contains("ADDCONTROL", ex.Message);
    }

    [Fact]
    public async Task EndlessLoopsAndRecursiveExpansionAreBounded()
    {
        await Assert.ThrowsAsync<SandyExecutionException>(() => Execute("~~StartCoda\n~~Coda 1", new ExecutionOptions { MaxSteps = 5 }));
        var options = new ExecutionOptions { MaxExpansionSteps = 5 }; options.Variables["X"] = "%%X%%";
        await Assert.ThrowsAsync<SandyExecutionException>(() => Execute("%%X%%", options));
    }

    [Fact]
    public async Task MalformedBlocksAreRejectedEvenInUnreachableBranches()
    {
        await Assert.ThrowsAsync<SandyExecutionException>(() => Execute("~~If 0\n~~Else\n~~Else\n~~EndIf"));
        await Assert.ThrowsAsync<SandyExecutionException>(() => Execute("~~If 0\n~~StartCoda\n~~EndIf\n~~Coda 0"));
    }

    [Fact]
    public async Task CancelledPromptDoesNotReturnPartialEdits()
    {
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => Execute("partial\n%%Name%%", new ExecutionOptions { PromptAsync = (_, _) => Task.FromResult<string?>(null) }));
        using var cts = new CancellationTokenSource(); cts.Cancel();
        var (library, template) = Library("text");
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new TemplateEngine().ExecuteAsync(library, template, cancellationToken: cts.Token));
    }

    [Fact]
    public async Task ConcurrentRunsNeverShareVariables()
    {
        var (library, template) = Library("%%Name%%");
        var engine = new TemplateEngine();
        var runs = Enumerable.Range(0, 20).Select(async n =>
        {
            var options = new ExecutionOptions(); options.Variables["Name"] = n.ToString();
            return (await engine.ExecuteAsync(library, template, options)).Outputs.Single().Text;
        });
        Assert.Equal(20, (await Task.WhenAll(runs)).Distinct().Count());
    }

    [Theory]
    [InlineData("1+2*3", "7")]
    [InlineData("(1+2)*3", "9")]
    [InlineData("2^3^2", "512")]
    [InlineData("3<>2 AND 4>=4", "1")]
    [InlineData("\"a\" & \"b\"", "ab")]
    [InlineData("9 MOD 4", "1")]
    public void EvaluatesWithoutExecutingCode(string expression, string expected)
        => Assert.Equal(expected, SandyExpression.Evaluate(expression, new Dictionary<string, string>()));

    [Fact]
    public async Task InlineStringOperationsUseOneBasedPositionsAndVariableArguments()
    {
        var result = await Execute("~~Set Name=William M Rawls\n%%Token::Name,3%% %%Mid::Name,9,1%% %%Before::Name,3%% %%After::Name%% %%Except::Name,2%% %%TokenCount::Name%%");
        Assert.Equal("Rawls M William M M Rawls William Rawls 3\r\n", Assert.Single(result.Outputs).Text);
    }
}
