using Sandy.Core;

namespace Sandy.Tests;

public sealed class EngineTests
{
    private static (Library Library, Template Template) Fixture(string code)
    {
        var library = new Library(); var category = new Category(); category.Fields["CategoryID"] = "1"; category.Fields["CategoryName"] = "Test";
        library.Categories.Add(category);
        var template = new Template(); template["TemplateID"] = "1"; template["CategoryID"] = "1"; template["TemplateName"] = "Test - One"; template["ShortTemplateName"] = "One"; template[Template.CodeField(CodeArea.Cursor)] = code;
        category.Templates.Add(template); return (library, template);
    }
    [Fact] public async Task DefaultsAndRepeatedPromptsResolveOnce()
    {
        var (library, template) = Fixture("%%Name::Title%% %%Name%%"); var calls = 0;
        var result = await new Engine().EvaluateAsync(library, template, prompt: (p, _) => { calls++; Assert.Equal("Title", p.Default); return Task.FromResult<string?>("Chosen"); });
        Assert.Equal("Chosen Chosen", Assert.Single(result.Outputs).Text); Assert.Equal(1, calls);
    }
    [Fact] public async Task VariablesAreCaseInsensitiveAndDefaultsCanReferenceVariables()
    {
        var (library, template) = Fixture("%%NAME%% %%Other::@name%%");
        var result = await new Engine().EvaluateAsync(library, template, new Dictionary<string, string> { ["name"] = "William" });
        Assert.Equal("William William", Assert.Single(result.Outputs).Text);
    }
    [Fact] public async Task CancelledPromptProducesNoResult()
    {
        var (library, template) = Fixture("Before %%Question%% after");
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new Engine().EvaluateAsync(library, template, prompt: (_, _) => Task.FromResult<string?>(null)));
    }
    [Fact] public async Task AssignmentAndInlineCommandsFollowSourceSyntax()
    {
        var (library, template) = Fixture("~~Set Name=William\r\n%%UCASE::Name%% %%LEFT::Name,4%% %%LEN::Name%%");
        var result = await new Engine().EvaluateAsync(library, template);
        Assert.Equal("WILLIAM Will 7", Assert.Single(result.Outputs).Text);
    }
    [Fact] public async Task UnknownCommandsAreRejectedWithContext()
    {
        var (library, template) = Fixture("Hello\n~~RegistryWrite Dangerous");
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new Engine().EvaluateAsync(library, template));
        Assert.Contains("Test - One, Cursor, line 2", error.Message);
    }
    [Fact] public async Task RecursiveVariablesAreBounded()
    {
        var (library, template) = Fixture("%%A%%");
        await Assert.ThrowsAsync<InvalidOperationException>(() => new Engine().EvaluateAsync(library, template, new Dictionary<string, string> { ["A"] = "%%A%%" }));
    }
    [Fact] public async Task RecursiveIncludesAreBounded()
    {
        var (library, template) = Fixture("~##~Include 2,One");
        await Assert.ThrowsAsync<InvalidOperationException>(() => new Engine().EvaluateAsync(library, template));
    }
    [Fact] public async Task SeparateAreasNeverMixAndFileOutputDoesNotWrite()
    {
        var (library, template) = Fixture("Cursor"); template[Template.CodeField(CodeArea.Declarations)] = "Top"; template[Template.CodeField(CodeArea.File)] = "File"; template["Filename"] = "NeverWrite.txt";
        var result = await new Engine().EvaluateAsync(library, template);
        Assert.Equal(3, result.Outputs.Count); Assert.Equal(CodeArea.Cursor, result.Outputs[0].Area); Assert.Equal(CodeArea.Declarations, result.Outputs[1].Area);
        Assert.Equal("NeverWrite.txt", result.Outputs[2].Target);
    }
    [Fact] public async Task CancellationStopsEvenWithoutPrompts()
    {
        var (library, template) = Fixture("Hello");
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new Engine().EvaluateAsync(library, template, cancellation: new CancellationToken(true)));
    }
    [Fact] public void LockedTemplateRejectsChangesAtomically()
    {
        var (library, template) = Fixture("Original"); template["Locked"] = "-1";
        Assert.Throws<InvalidOperationException>(() => library.Update(template, new Dictionary<CodeArea, string> { [CodeArea.Cursor] = "Edited" }, ""));
        Assert.Equal("Original", template.Code(CodeArea.Cursor));
    }
    [Fact] public void RevisionOverflowRejectsChangesAtomically()
    {
        var (library, template) = Fixture("Original"); template["RevisionCount"] = int.MaxValue.ToString();
        Assert.Throws<OverflowException>(() => library.Update(template, new Dictionary<CodeArea, string> { [CodeArea.Cursor] = "Edited" }, ""));
        Assert.Equal("Original", template.Code(CodeArea.Cursor));
    }
    [Fact] public void SaveAndLoadPreserveUnknownFieldsAndLineEndings()
    {
        var (library, template) = Fixture("A\r\nB\r\n"); template["HistoricalField"] = "<original>&\r\n";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xml");
        try { library.Save(path); var loaded = Library.Load(path).Templates.Single(); Assert.Equal(template.Fields, loaded.Fields); }
        finally { File.Delete(path); }
    }
    [Fact] public void DuplicateIdsAreRejected()
    {
        var (library, template) = Fixture("Original"); library.Categories[0].Templates.Add(template);
        Assert.Throws<InvalidDataException>(library.Validate);
    }
    [Fact] public void UnchangedCodeDoesNotBumpRevision()
    {
        var (library, template) = Fixture("Original"); template["RevisionCount"] = "7";
        library.Update(template, new Dictionary<CodeArea, string> { [CodeArea.Cursor] = "Original" }, "");
        Assert.Equal("7", template["RevisionCount"]);
    }
    [Fact] public void DtdIsRejected()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xml");
        try { File.WriteAllText(path, "<!DOCTYPE dataroot [<!ENTITY x 'bad'>]><dataroot>&x;</dataroot>"); Assert.Throws<System.Xml.XmlException>(() => Library.Load(path)); }
        finally { File.Delete(path); }
    }
}
