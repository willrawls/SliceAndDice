using Sandy.Core;
namespace Sandy.Tests;
public sealed class LibraryPatchTests
{
    [Fact]public void PatchPreservesMovesDeletionsUnknownFieldsAndOrder()
    {
        var(before,t)=LibraryMutationTests.Fixture("original");var after=before.Clone();var c=new Category();c.Fields["CategoryID"]="2";c.Fields["CategoryName"]="Moved";after.Categories.Insert(0,c);
        var moved=after.Templates.Single();after.Categories[1].Templates.Remove(moved);c.Templates.Add(moved);moved["CategoryID"]="2";moved["Unknown field"]="kept";
        var patch=LibraryPatch.FromJson(LibraryPatch.Create(before,after).ToJson());var applied=patch.Apply(before);Assert.Equal(after.Serialize(),applied.Serialize());Assert.Equal("original",t.Code(CodeArea.Cursor));
    }
    [Fact]public void TamperedAndStalePatchesAreAtomic()
    {
        var(before,_)=LibraryMutationTests.Fixture("original");var after=before.Clone();after.Templates.Single()["memoCodeAtCursor"]="changed";var patch=LibraryPatch.Create(before,after);
        patch.Templates[0].Fields!["memoCodeAtCursor"]="tampered";Assert.Throws<InvalidDataException>(()=>patch.Apply(before));Assert.Equal("original",before.Templates.Single().Code(CodeArea.Cursor));
        before.Templates.Single()["Favorite"]="-1";Assert.Throws<InvalidOperationException>(()=>patch.Apply(before));
    }
    [Fact]public void ProtectedTemplateCannotBeRemovedThroughPatch()
    {
        var(before,t)=LibraryMutationTests.Fixture("");t["Undeletable"]="-1";var after=before.Clone();after.Categories[0].Templates.Clear();Assert.Throws<InvalidOperationException>(()=>LibraryPatch.Create(before,after).Apply(before));
    }
    [Fact]public void NavigationDropsForwardHistoryAfterSelectingAnotherTemplate()
    {
        var state=new Sandy.Blazor.Workspace();state.NewLibrary();state.AddCategory("Test");state.AddTemplate("One");var one=state.Selected!;state.AddTemplate("Two");var two=state.Selected!;
        state.Navigate(-1);Assert.Same(one,state.Selected);Assert.True(state.CanForward);state.AddTemplate("Three");Assert.False(state.CanForward);state.Navigate(-1);Assert.Same(one,state.Selected);
    }
    [Fact]public void ImportsReassignIdsAndMovesPreserveDrafts()
    {
        var(l,t)=LibraryMutationTests.Fixture("original");var state=new Sandy.Blazor.Workspace();state.Load(l,null);state.Current!.Code[CodeArea.Cursor]="draft";state.Import(l);
        Assert.Equal(2,state.Library.Templates.Select(t=>t.Id).Distinct().Count());state.Select(state.Library.Templates.First());Assert.Equal("draft",state.Current!.Code[CodeArea.Cursor]);
        state.AddCategory("Moved");var category=state.SelectedCategory!;state.MoveTemplate(state.Selected!.Id,category.Fields["CategoryID"]);Assert.Same(category,state.SelectedCategory);Assert.Equal("draft",state.Current!.Code[CodeArea.Cursor]);
    }
    [Fact]public async Task LegacyGeneratorUsesCollectionMemberAndPropertyTemplatesInOrder()
    {
        var(l,_)=LibraryMutationTests.Fixture("");var category=l.Categories[0];category.Templates.Clear();var n=0;
        foreach(var pair in new Dictionary<string,string>{{"Collection, No Parent, No Child","Collection %%Table Name%%"},{"Collection Member, Terminal","Member %%Object Name%%"},{"Property - Long","Field %%Field Name%%"},{"Property - String","Field %%Property Name%%"}})
        {var t=new Template();t["TemplateID"]=(++n).ToString();t["CategoryID"]="1";t["TemplateName"]="Scripts - "+pair.Key;t["ShortTemplateName"]=pair.Key;t["memoCodeAtCursor"]=pair.Value+"\r\n";category.Templates.Add(t);}
        var design=new DatabaseDesign();var table=design.AddTable("People");table.Fields.Add(new(){Name="PeopleID",DaoType=4,PrimaryKey=true,Nullable=false});table.Fields.Add(new(){Name="Name",DaoType=10});
        var files=await LegacyDatabaseGenerator.GenerateAsync(design,l,"Scripts",[table]);Assert.Equal("Collection People\r\nMember People\r\nField PeopleID\r\nField Name\r\n",files["Generated.bas"]);
    }
}
