using System.Xml;
using System.Xml.Linq;

namespace Sandy.Core;

public enum TemplateArea { Cursor, Declarations, Bottom, File }

/// <summary>Original field names and values are retained, including fields unknown to this version.</summary>
public sealed class SandyTemplate
{
    public Dictionary<string, string> Fields { get; } = new(StringComparer.OrdinalIgnoreCase);
    public string Get(string name) => Fields.TryGetValue(name, out var value) ? value : "";
    public string Id => Get("TemplateID");
    public string Name => Get("TemplateName");
    public string ShortName => Get("ShortTemplateName");
    public string CategoryId => Get("CategoryID");
    public string GetCode(TemplateArea area) => Get(FieldFor(area));
    public static string FieldFor(TemplateArea area) => area switch
    {
        TemplateArea.Cursor => "memoCodeAtCursor",
        TemplateArea.Declarations => "memoCodeAtTop",
        TemplateArea.Bottom => "memoCodeAtBottom",
        _ => "memoCodeToFile"
    };
}

public sealed class SandyCategory
{
    public Dictionary<string, string> Fields { get; } = new(StringComparer.OrdinalIgnoreCase);
    public string Id => Fields.TryGetValue("CategoryID", out var id) ? id : "";
    public string Name => Fields.TryGetValue("CategoryName", out var name) ? name : "";
    public List<SandyTemplate> Templates { get; } = new();
}

public sealed class SandyLibrary
{
    public string Name { get; set; } = "";
    public List<SandyCategory> Categories { get; } = new();
    public IEnumerable<SandyTemplate> Templates => Categories.SelectMany(c => c.Templates);

    public SandyTemplate Find(string name, string? currentCategory = null)
    {
        var matches = Templates.Where(t => t.Name.Equals(name, StringComparison.OrdinalIgnoreCase)).ToList();
        if (matches.Count == 0)
            matches = Categories.Where(c => currentCategory == null || c.Name.Equals(currentCategory, StringComparison.OrdinalIgnoreCase))
                .SelectMany(c => c.Templates).Where(t => t.ShortName.Equals(name, StringComparison.OrdinalIgnoreCase)).ToList();
        if (matches.Count != 1)
            throw new InvalidOperationException(matches.Count == 0 ? $"Template not found: {name}" : $"Template name is ambiguous: {name}");
        return matches[0];
    }

    public string CategoryOf(SandyTemplate template) => Categories.First(c => c.Templates.Contains(template)).Name;
}

public static class LibraryXml
{
    public static SandyLibrary Load(string path)
    {
        using var stream = System.IO.File.OpenRead(path);
        return Read(stream, System.IO.Path.GetFileNameWithoutExtension(path));
    }

    public static SandyLibrary Read(Stream stream, string name = "Imported")
    {
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        var root = XDocument.Load(reader).Root ?? throw new InvalidDataException("Missing library root.");
        if (root.Name.LocalName != "dataroot") throw new InvalidDataException("Expected an Access Category.xml export (dataroot).");
        var library = new SandyLibrary { Name = name };
        foreach (var element in root.Elements().Where(e => e.Name.LocalName == "Category"))
        {
            var category = new SandyCategory();
            foreach (var field in element.Elements().Where(e => e.Name.LocalName != "Template"))
                category.Fields.Add(field.Name.LocalName, field.Value);
            if (category.Id.Length == 0 || library.Categories.Any(c => c.Id == category.Id))
                throw new InvalidDataException($"Missing or duplicate category ID: {category.Id}");
            foreach (var item in element.Elements().Where(e => e.Name.LocalName == "Template"))
            {
                var template = new SandyTemplate();
                foreach (var field in item.Elements()) template.Fields.Add(field.Name.LocalName, field.Value);
                if (template.Id.Length == 0 || library.Templates.Any(t => t.Id == template.Id) || category.Templates.Any(t => t.Id == template.Id))
                    throw new InvalidDataException($"Missing or duplicate template ID: {template.Id}");
                if (template.CategoryId != category.Id) throw new InvalidDataException($"Template {template.Id} has mismatched CategoryID.");
                category.Templates.Add(template);
            }
            library.Categories.Add(category);
        }
        return library;
    }

    public static void Save(SandyLibrary library, string path)
    {
        var root = new XElement("dataroot", library.Categories.Select(c => new XElement("Category",
            c.Fields.Select(f => new XElement(f.Key, f.Value)),
            c.Templates.Select(t => new XElement("Template", t.Fields.Select(f => new XElement(f.Key, f.Value)))))));
        // Entitize CRs so importing this file preserves CRLFs in edited template bodies.
        using var writer = XmlWriter.Create(path, new XmlWriterSettings
        {
            Encoding = new System.Text.UTF8Encoding(false), Indent = true, NewLineHandling = NewLineHandling.Entitize
        });
        new XDocument(root).Save(writer);
    }
}
