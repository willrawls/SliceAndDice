using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.Extensions.DependencyInjection;
using Sandy.Blazor;
using Sandy.Blazor.Components;
using Sandy.Core;

namespace Sandy.Tests;
public sealed class WorkspaceTests
{
    private static Workspace Create()
    {
        var state = new Workspace(); state.NewLibrary(); state.AddCategory("Test"); state.AddTemplate("One"); state.Current!.Code[CodeArea.Cursor] = "original"; state.Update(); state.AddTemplate("Two"); return state;
    }
    [Fact] public void SwitchingKeepsDraftWithoutChangingCommittedTemplate()
    {
        var state = Create(); var first = state.Library.Templates.First(); var second = state.Library.Templates.Last();
        state.Select(first); state.Current!.Code[CodeArea.Cursor] = "draft"; state.Select(second); state.Select(first);
        Assert.Equal("draft", state.Current!.Code[CodeArea.Cursor]); Assert.Equal("original", first.Code(CodeArea.Cursor)); Assert.True(state.IsDirty);
    }
    [Fact] public void RenamingBumpsRevisionAndNoOpDoesNot()
    {
        var state = Create(); var t = state.Selected!; state.Update(); var revision = int.Parse(t["RevisionCount"]);
        state.Current!.Name = "Renamed"; state.Update(); Assert.Equal(revision + 1, int.Parse(t["RevisionCount"]));
        state.Update(); Assert.Equal(revision + 1, int.Parse(t["RevisionCount"]));
    }
    [Fact] public void FailedSaveDoesNotPartiallyCommitEarlierDrafts()
    {
        var state = Create(); var first = state.Library.Templates.First(); var second = state.Selected!;
        second["Locked"] = "-1"; state.Load(state.Library, null); state.Select(first); state.Current!.Code[CodeArea.Cursor] = "pending";
        state.Select(second); state.Current!.Code[CodeArea.Cursor] = "blocked";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xml");
        try { Assert.Throws<InvalidOperationException>(() => state.Save(path)); Assert.False(File.Exists(path)); Assert.Equal("original", first.Code(CodeArea.Cursor)); }
        finally { File.Delete(path); }
    }
    [Fact] public void SavingAppliesAllDraftsAndReloadsCleanly()
    {
        var state = Create(); var first = state.Library.Templates.First(); state.Select(first); state.Current!.Code[CodeArea.Cursor] = "saved";
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xml");
        try { state.Save(path); Assert.False(state.IsDirty); Assert.Equal("saved", state.Library.Templates.First().Code(CodeArea.Cursor)); Assert.Throws<InvalidOperationException>(() => state.Save(path)); }
        finally { File.Delete(path); }
    }
    [Fact] public void FileWriteFailureKeepsModelAndDrafts()
    {
        var state = Create(); var first = state.Library.Templates.First(); state.Select(first); state.Current!.Code[CodeArea.Cursor] = "pending";
        Assert.Throws<DirectoryNotFoundException>(() => state.Save(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString(), "copy.xml")));
        Assert.Equal("original", first.Code(CodeArea.Cursor)); Assert.Equal("pending", state.Current.Code[CodeArea.Cursor]);
    }
    [Fact] public async Task RazorComponentsRenderTheWorkspaceAndEveryAuxiliaryForm()
    {
        var state = Create(); var services = new ServiceCollection(); services.AddLogging(); services.AddSingleton(state); services.AddSingleton<IDesktopServices, TestDesktop>();
        using var provider = services.BuildServiceProvider(); await using var renderer = new HtmlRenderer(provider, provider.GetRequiredService<Microsoft.Extensions.Logging.ILoggerFactory>());
        await renderer.Dispatcher.InvokeAsync(async () =>
        {
            var workbench = await renderer.RenderComponentAsync<Workbench>(); Assert.Contains("code-editor", workbench.ToHtmlString());
            Assert.Equal(21, FormCatalog.All.Count);
            foreach (var form in FormCatalog.All.Where(f => f.ClassName != "frmMain"))
            {
                var result = await renderer.RenderComponentAsync<LegacyDialog>(ParameterView.FromDictionary(new Dictionary<string, object?> { ["Definition"] = form }));
                Assert.Contains("legacy-fields", result.ToHtmlString());
            }
        });
    }
    private sealed class TestDesktop : IDesktopServices
    {
        public string? OpenLibrary() => null;
        public string? SaveCopy() => null;
        public Library ReadLibrary(string path) => Library.Load(path);
        public void OpenUrl(string url) { }
    }
}
