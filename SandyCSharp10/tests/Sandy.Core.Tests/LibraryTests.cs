using System.Text;
using Sandy.Core;

namespace Sandy.Core.Tests;

public class LibraryTests
{
    private static string RepositoryRoot
    {
        get
        {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null && !Directory.Exists(Path.Combine(directory.FullName, "MdbGallery"))) directory = directory.Parent;
            return directory?.FullName ?? throw new DirectoryNotFoundException("Cannot locate repository test corpus.");
        }
    }

    [Theory]
    [InlineData("SliceAndDice 2000", 27, 550)]
    [InlineData("SandyClipboard", 4, 36)]
    [InlineData("SandyColdFusion", 16, 171)]
    [InlineData("SandySqlSp", 24, 328)]
    [InlineData("OldSliceAndDice", 23, 313)]
    [InlineData("GTESliceAndDice", 21, 380)]
    [InlineData("Copy of SandyBeneMO", 26, 419)]
    public void HistoricalXmlLibrariesRetainAllFieldsThroughRoundtrip(string name, int categories, int templates)
    {
        var path = Path.Combine(RepositoryRoot, "MdbGallery", "BrokenDown", name, "Category.xml");
        var library = LibraryXml.Load(path);
        Assert.Equal(categories, library.Categories.Count);
        Assert.Equal(templates, library.Templates.Count());
        var saved = Path.GetTempFileName();
        try
        {
            LibraryXml.Save(library, saved);
            var restored = LibraryXml.Load(saved);
            Assert.Equal(library.Categories.Select(c => c.Fields.ToArray()), restored.Categories.Select(c => c.Fields.ToArray()));
            Assert.Equal(library.Templates.Select(t => t.Fields.ToArray()), restored.Templates.Select(t => t.Fields.ToArray()));
        }
        finally { File.Delete(saved); }
    }

    [Fact]
    public void CsvProfilePreservesMultilineBodiesAndRejectsUnverifiedSchema()
    {
        var folder = Path.Combine(RepositoryRoot, "MdbGallery", "BrokenDown", "SliceAndDice 2000");
        using var categories = File.OpenText(Path.Combine(folder, "Category.csv"));
        using var templates = File.OpenText(Path.Combine(folder, "Template.csv"));
        var library = LibraryCsv.Read(categories, templates);
        Assert.Equal(550, library.Templates.Count());
        Assert.Contains("\r\n", library.Find("Basic - Public Function").GetCode(TemplateArea.Cursor));
        Assert.Contains("%%New Function Name%%", library.Find("Basic - Public Function").GetCode(TemplateArea.Cursor));
        using var categoryReader = new StringReader("1,Basic,0,0,0,,,,");
        using var badSchema = new StringReader(string.Join(",", Enumerable.Repeat("0", 22)));
        Assert.Throws<InvalidDataException>(() => LibraryCsv.Read(categoryReader, badSchema));
    }

    [Fact]
    public async Task HistoricalPublicFunctionTemplateRetainsPromptAndOutputBehavior()
    {
        var path = Path.Combine(RepositoryRoot, "MdbGallery", "BrokenDown", "SliceAndDice 2000", "Category.xml");
        var library = LibraryXml.Load(path);
        var options = new ExecutionOptions();
        options.Variables["New Function Name"] = "FindWidget";
        options.Variables["Parameters (If Any)"] = "ByVal id As Long";
        options.Variables["Return Type"] = "Boolean";
        var result = await new TemplateEngine().ExecuteAsync(library, library.Find("Basic - Public Function"), options);
        Assert.Equal("Public Function FindWidget(ByVal id As Long) As Boolean\r\n    \r\nEnd Function\r\n", Assert.Single(result.Outputs).Text);
    }

    [Fact]
    public void XmlCannotResolveEntitiesAndRejectsInvalidCategoryRelationships()
    {
        using var entity = new MemoryStream(Encoding.UTF8.GetBytes("<!DOCTYPE dataroot [<!ENTITY x SYSTEM 'file:///never-read'>]><dataroot>&x;</dataroot>"));
        Assert.Throws<System.Xml.XmlException>(() => LibraryXml.Read(entity));
        using var mismatch = new MemoryStream(Encoding.UTF8.GetBytes("<dataroot><Category><CategoryID>1</CategoryID><Template><TemplateID>1</TemplateID><CategoryID>2</CategoryID></Template></Category></dataroot>"));
        Assert.Throws<InvalidDataException>(() => LibraryXml.Read(mismatch));
    }

    [Fact]
    public void EditedCrlfCodeAndUnknownMetadataSurviveExport()
    {
        var (library, template) = EngineTests.Library("line1\r\nline2\r\n");
        template.Fields["FutureField"] = "value";
        var path = Path.GetTempFileName();
        try
        {
            LibraryXml.Save(library, path);
            var restored = LibraryXml.Load(path).Templates.Single();
            Assert.Equal(template.GetCode(TemplateArea.Cursor), restored.GetCode(TemplateArea.Cursor));
            Assert.Equal("value", restored.Get("FutureField"));
        }
        finally { File.Delete(path); }
    }
}
