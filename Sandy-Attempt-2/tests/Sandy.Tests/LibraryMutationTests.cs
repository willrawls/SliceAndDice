using Sandy.Core;
namespace Sandy.Tests;
public sealed class LibraryMutationTests
{
    internal static (Library Library,Template Template) Fixture(string source)
    {
        var library=new Library();var category=new Category();category.Fields["CategoryID"]="1";category.Fields["CategoryName"]="Scripts";library.Categories.Add(category);
        var template=new Template();template["TemplateID"]="1";template["CategoryID"]="1";template["TemplateName"]="Scripts - Run";template["ShortTemplateName"]="Run";template["memoCodeAtCursor"]=source;category.Templates.Add(template);return(library,template);
    }
    [Fact]public async Task NewTemplateChangesStayIsolatedUntilCommit()
    {
        var(l,t)=Fixture("~~GotoSandyTemplate Generated - Example\r\n~~SetCodeAreaText AtCursor=hello\r\n~~SetCodeArea Decl=private int n;");
        var r=await new Engine().EvaluateAsync(l,t);Assert.Single(l.Templates);var mutation=Assert.IsType<LibraryMutation>(r.LibraryMutation);
        var result=mutation.Commit(l);Assert.Equal(2,result.Templates.Count());Assert.Equal("hello",result.Find("Example","Generated").Code(CodeArea.Cursor));Assert.Equal("1",result.Find("Example","Generated")["RevisionCount"]);
    }
    [Fact]public async Task CancellationDoesNotMutateLibrary()
    {
        var(l,t)=Fixture("~~GotoSandyTemplate Generated - Example\r\n~~SetCodeArea Cursor=hello\r\n~~Cancel");var before=l.Serialize();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(()=>new Engine().EvaluateAsync(l,t));Assert.Equal(before,l.Serialize());
    }
    [Fact]public void ConcurrentLibraryChangesRejectStalePlan()
    {
        var(l,t)=Fixture("");var mutation=new LibraryMutation(l,t);mutation.SetCode("Cursor","new");t["Favorite"]="-1";Assert.Throws<InvalidOperationException>(()=>mutation.Commit(l));Assert.Equal("",t.Code(CodeArea.Cursor));
    }
    [Fact]public void LockedTargetCannotBeChanged()
    {
        var(l,t)=Fixture("");t["Locked"]="-1";var mutation=new LibraryMutation(l,t);Assert.Throws<InvalidOperationException>(()=>mutation.SetCode("Cursor","new"));
    }
    [Fact]public void SerializationKeepsAuxiliaryRecordsAndCarriageReturnsInsideRoot()
    {
        var(l,_)=Fixture("a\r\nb\r\n");l.AuxiliaryRecords.Add(System.Xml.Linq.XElement.Parse("<SystemInfo><Value>kept</Value></SystemInfo>"));
        var xml=System.Xml.Linq.XDocument.Parse(l.Serialize());Assert.Equal("kept",xml.Root!.Element("SystemInfo")!.Element("Value")!.Value);Assert.Equal("a\r\nb\r\n",xml.Descendants("memoCodeAtCursor").Single().Value);
    }
    [Fact]public async Task BlazorApplyPreservesUnrelatedDrafts()
    {
        var(l,t)=Fixture("~~GotoSandyTemplate Generated - Example\r\n~~SetCodeArea Cursor=hello");var state=new Sandy.Blazor.Workspace();state.Load(l,null);
        await state.PreviewAsync();var plan=state.PreviewPlan!;state.Current!.ShortName="Unrelated draft";state.CompleteApplication(plan);
        Assert.Equal("Example",state.Current!.ShortName);Assert.True(state.IsDirty);state.Select(state.Library.Templates.Single(x=>x.Id==t.Id));Assert.Equal("Unrelated draft",state.Current!.ShortName);
    }
}
