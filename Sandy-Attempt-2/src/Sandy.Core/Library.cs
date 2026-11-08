using System.Xml;
using System.Xml.Linq;

namespace Sandy.Core;

public enum CodeArea { Declarations, Cursor, Bottom, File }

public sealed class Template
{
    public Dictionary<string, string> Fields { get; } = new(StringComparer.OrdinalIgnoreCase);
    public string this[string field] { get => Fields.GetValueOrDefault(field, ""); set => Fields[field] = value; }
    public string Name => this["TemplateName"];
    public string Id => this["TemplateID"];
    public bool Locked => Flag("Locked");
    public bool Flag(string field) => this[field] is "1" or "-1" || this[field].Equals("true", StringComparison.OrdinalIgnoreCase);
    public static string CodeField(CodeArea area) => area switch
    {
        CodeArea.Declarations => "memoCodeAtTop", CodeArea.Cursor => "memoCodeAtCursor",
        CodeArea.Bottom => "memoCodeAtBottom", _ => "memoCodeToFile"
    };
    public string Code(CodeArea area) => this[CodeField(area)];
    public override string ToString() => Name;
}

public sealed class Category
{
    public Dictionary<string, string> Fields { get; } = new(StringComparer.OrdinalIgnoreCase);
    public string Name => Fields.GetValueOrDefault("CategoryName", "");
    public List<Template> Templates { get; } = [];
    public override string ToString() => Name;
}

public sealed class Library
{
    public Library Clone()
    {
        var copy=new Library();copy.AuxiliaryRecords.AddRange(AuxiliaryRecords.Select(e=>new XElement(e)));
        foreach(var category in Categories){var c=new Category();foreach(var p in category.Fields)c.Fields[p.Key]=p.Value;foreach(var t in category.Templates){var clone=new Template();foreach(var p in t.Fields)clone[p.Key]=p.Value;c.Templates.Add(clone);}copy.Categories.Add(c);}
        return copy;
    }
    public string Serialize()
    {
        Validate();var builder=new System.Text.StringBuilder();
        using(var writer=XmlWriter.Create(builder,new XmlWriterSettings{OmitXmlDeclaration=true,Indent=true,NewLineHandling=NewLineHandling.Entitize}))Document().Save(writer);
        return builder.ToString();
    }
    private XDocument Document()
    {
        var root=new XElement("dataroot",Categories.Select(c=>new XElement("Category",c.Fields.Select(p=>new XElement(XmlConvert.EncodeLocalName(p.Key),p.Value)),c.Templates.Select(t=>new XElement("Template",t.Fields.Select(p=>new XElement(XmlConvert.EncodeLocalName(p.Key),p.Value)))))));
        root.Add(AuxiliaryRecords.Select(e=>new XElement(e)));return new(root);
    }
    public List<Category> Categories { get; } = [];
    public List<XElement> AuxiliaryRecords { get; } = [];
    public IEnumerable<Template> Templates => Categories.SelectMany(x => x.Templates);

    public static Library Load(string filename)
    {
        using var reader = XmlReader.Create(filename, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        var root = XDocument.Load(reader).Root ?? throw new InvalidDataException("Missing XML root.");
        if (root.Name.LocalName != "dataroot") throw new InvalidDataException("Choose an Access Category.xml export, not a COM type-library XML file.");
        var result = new Library();
        result.AuxiliaryRecords.AddRange(root.Elements().Where(e=>e.Name.LocalName!="Category").Select(e=>new XElement(e)));
        foreach (var element in root.Elements().Where(x => x.Name.LocalName == "Category"))
        {
            var category = new Category();
            foreach (var field in element.Elements().Where(x => x.Name.LocalName != "Template")) category.Fields.Add(XmlConvert.DecodeName(field.Name.LocalName), field.Value);
            foreach (var item in element.Elements().Where(x => x.Name.LocalName == "Template"))
            {
                var template = new Template();
                foreach (var field in item.Elements()) template.Fields.Add(XmlConvert.DecodeName(field.Name.LocalName), field.Value);
                category.Templates.Add(template);
            }
            result.Categories.Add(category);
        }
        result.Validate();
        return result;
    }

    public void Validate()
    {
        var categoryIds = new HashSet<string>();
        var templateIds = new HashSet<string>();
        foreach (var category in Categories)
        {
            var id = category.Fields.GetValueOrDefault("CategoryID", "");
            if (id.Length == 0 || !categoryIds.Add(id)) throw new InvalidDataException("Missing or duplicate CategoryID: " + id);
            foreach (var template in category.Templates)
                if (template["CategoryID"] != id || template["TemplateID"].Length == 0 || !templateIds.Add(template["TemplateID"]))
                    throw new InvalidDataException("Invalid template identity: " + template.Name);
        }
    }

    public void Save(string filename)
    {
        Validate();
        var document = Document();
        // Complete the new document before replacing an existing working copy.
        var temporary = filename + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var writer = XmlWriter.Create(temporary, new XmlWriterSettings { Indent = true, NewLineHandling = NewLineHandling.Entitize, Encoding = new System.Text.UTF8Encoding(false) }))
                document.Save(writer);
            System.IO.File.Move(temporary, filename, true);
        }
        finally { if (System.IO.File.Exists(temporary)) System.IO.File.Delete(temporary); }
    }

    public string CategoryOf(Template template) => Categories.Single(c => c.Templates.Contains(template) ||
        template["CategoryID"].Length > 0 && c.Fields.GetValueOrDefault("CategoryID") == template["CategoryID"]).Name;

    public Template Find(string name, string? currentCategory = null)
    {
        var found = Templates.Where(t => t.Name.Equals(name, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (found.Length == 0) found = Categories.Where(c => currentCategory == null || c.Name.Equals(currentCategory, StringComparison.OrdinalIgnoreCase)).SelectMany(c => c.Templates).Where(t => t["ShortTemplateName"].Equals(name, StringComparison.OrdinalIgnoreCase)).ToArray();
        return found.Length == 1 ? found[0] : throw new InvalidOperationException("Missing or ambiguous template: " + name);
    }

    public void Update(Template template, IReadOnlyDictionary<CodeArea, string> code, string filename)
    {
        if (!Templates.Contains(template)) throw new InvalidOperationException("Template belongs to another library.");
        if (code.All(x => template.Code(x.Key) == x.Value) && template["Filename"] == filename) return;
        if (template.Locked) throw new InvalidOperationException("The template is locked.");
        var revision = checked(int.Parse(template["RevisionCount"].Length == 0 ? "0" : template["RevisionCount"], System.Globalization.CultureInfo.InvariantCulture) + 1);
        foreach (var entry in code) template[Template.CodeField(entry.Key)] = entry.Value;
        template["Filename"] = filename;
        template["RevisionCount"] = revision.ToString(System.Globalization.CultureInfo.InvariantCulture);
        template["DateModified"] = DateTime.Now.ToString("O");
    }
}
