using Sandy.Core;
namespace Sandy.Tests;
public sealed class FileCommitTests
{
    [Theory][InlineData("utf16")][InlineData("utf8bom")][InlineData("legacy")]
    public void SourceEncodingAndBomSurviveEditing(string format)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        var encoding=format=="utf16"?System.Text.Encoding.Unicode:format=="utf8bom"?new System.Text.UTF8Encoding(true):System.Text.Encoding.GetEncoding(1252);
        var bytes=encoding.GetPreamble().Concat(encoding.GetBytes("café\r\n")).ToArray();
        Assert.Equal("café\r\n",TextFiles.Decode(bytes));Assert.Equal(bytes,TextFiles.Encode(TextFiles.Decode(bytes),bytes));
        Assert.Equal(encoding.GetPreamble().Concat(encoding.GetBytes("café edited\r\n")),TextFiles.Encode("café edited\r\n",bytes));
    }
    [Fact] public void WritesAndRollbackRestoreExactOriginalBytes()
    {
        var root=Path.Combine(Path.GetTempPath(),"Sandy-"+Guid.NewGuid());Directory.CreateDirectory(root);
        try
        {
            var path=Path.Combine(root,"source.txt");var bytes=new byte[]{0xEF,0xBB,0xBF,65,13,10};File.WriteAllBytes(path,bytes);
            var plan=new FileCommit(root);plan.AddDocument(new PlannedDocument(path,"A\r\n"){Text="Edited\r\n"});plan.AddEffect(new("FileOverwrite","new.txt","New"));plan.Commit();
            Assert.Equal("Edited\r\n",File.ReadAllText(path));Assert.True(File.Exists(Path.Combine(root,"new.txt")));Assert.True(File.Exists(Path.Combine(plan.JournalPath!,"journal.json")));
            plan.Rollback();Assert.Equal(bytes,File.ReadAllBytes(path));Assert.False(File.Exists(Path.Combine(root,"new.txt")));
        }
        finally{Directory.Delete(root,true);}
    }
    [Fact] public void RejectsEscapingPathsBeforeWriting()
    {
        var plan=new FileCommit(Path.GetTempPath());Assert.Throws<InvalidOperationException>(()=>plan.AddEffect(new("FileOverwrite","../escaped.txt","No")));
    }
    [Fact] public void ConcurrentFileChangesPreventCommit()
    {
        var root=Path.Combine(Path.GetTempPath(),"Sandy-"+Guid.NewGuid());Directory.CreateDirectory(root);
        try
        {
            var path=Path.Combine(root,"source.txt");File.WriteAllText(path,"Original");var plan=new FileCommit(root);plan.AddEffect(new("FileOverwrite",path,"Planned"));File.WriteAllText(path,"Concurrent");
            Assert.Throws<IOException>(plan.Commit);Assert.Equal("Concurrent",File.ReadAllText(path));
        }
        finally{Directory.Delete(root,true);}
    }
    [Fact] public async Task CSharpPlacementIgnoresBracesInStringsAndComments()
    {
        var lib=new Library();var c=new Category();c.Fields["CategoryID"]="1";c.Fields["CategoryName"]="T";lib.Categories.Add(c);
        var t=new Template();t["TemplateID"]="1";t["CategoryID"]="1";t["TemplateName"]="Test";t["memoCodeAtBottom"]="void Added() {}\r\n";c.Templates.Add(t);
        var editor=new LegacyEditor("test.cs","class C { string x = \"}\"; /* } */\r\n}\r\n");await new Engine().EvaluateAsync(lib,t,host:editor);
        Assert.EndsWith("void Added() {}\r\n}\r\n",editor.Active.Text);
    }
}
