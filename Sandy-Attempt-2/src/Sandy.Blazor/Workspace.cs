using Sandy.Core;
namespace Sandy.Blazor;

public interface IDesktopServices
{
    string? OpenLibrary();
    string? SaveCopy();
    Library ReadLibrary(string path);
    void OpenUrl(string url);
    IReadOnlyDictionary<string, string[]> ReadSchema(string path) => throw new NotSupportedException("Database schema inspection is not available in this host.");
}

public sealed class Draft
{
    public string Name { get; set; } = "";
    public string ShortName { get; set; } = "";
    public string Filename { get; set; } = "";
    public Dictionary<CodeArea, string> Code { get; } = [];
    public bool Favorite { get; set; }
    public bool Locked { get; set; }
    public bool Undeletable { get; set; }
    public static Draft From(Template t)
    {
        var draft = new Draft { Name = t.Name, ShortName = t["ShortTemplateName"], Filename = t["Filename"], Favorite = t.Flag("Favorite"), Locked = t.Locked, Undeletable = t.Flag("Undeletable") };
        foreach (var area in Enum.GetValues<CodeArea>()) draft.Code[area] = t.Code(area);
        return draft;
    }
    public bool Matches(Template t) => Name == t.Name && ShortName == t["ShortTemplateName"] && Filename == t["Filename"] && Favorite == t.Flag("Favorite") && Locked == t.Locked && Undeletable == t.Flag("Undeletable") && Code.All(x => t.Code(x.Key) == x.Value);
}

public sealed class Workspace
{
    private readonly Dictionary<Template, Draft> drafts = [];
    private bool modelDirty;
    public Library Library { get; private set; } = new();
    public Template? Selected { get; private set; }
    public Category? SelectedCategory { get; private set; }
    public Draft? Current => Selected == null ? null : drafts[Selected];
    public string? SourcePath { get; private set; }
    public bool IsDirty => modelDirty || drafts.Any(x => !x.Value.Matches(x.Key));
    public string Variables { get; set; } = "";
    public string Filter { get; set; } = "";
    public string Preview { get; private set; } = "";
    public string Status { get; set; } = "Ready";
    public bool IsBusy { get; private set; }
    public void Load(Library library, string? path)
    {
        library.Validate(); Library = library; SourcePath = path == null ? null : Path.GetFullPath(path);
        drafts.Clear(); Selected = null; SelectedCategory = library.Categories.FirstOrDefault(); modelDirty = false; Preview = "";
        if (library.Templates.FirstOrDefault() is { } first) Select(first);
        Status = $"{library.Categories.Count} categories · {library.Templates.Count()} templates";
    }
    public void Select(Template template)
    {
        if (!Library.Templates.Contains(template)) throw new InvalidOperationException("Template belongs to another library.");
        Selected = template; SelectedCategory = Library.Categories.Single(c => c.Templates.Contains(template));
        drafts.TryAdd(template, Draft.From(template));
    }
    private static string NextId(IEnumerable<string> ids) => checked(ids.Select(x => int.TryParse(x, out var n) ? n : 0).DefaultIfEmpty().Max() + 1).ToString();
    public void NewLibrary() => Load(new Library(), null);
    public void AddCategory(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Enter a category name.");
        var category = new Category(); category.Fields["CategoryID"] = NextId(Library.Categories.Select(c => c.Fields["CategoryID"])); category.Fields["CategoryName"] = name.Trim();
        Library.Categories.Add(category); SelectedCategory = category; modelDirty = true;
    }
    public void AddTemplate(string name)
    {
        if (SelectedCategory == null) throw new InvalidOperationException("Create a category first.");
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Enter a template name.");
        var t = new Template(); t["CategoryID"] = SelectedCategory.Fields["CategoryID"]; t["TemplateID"] = NextId(Library.Templates.Select(x => x["TemplateID"]));
        t["ShortTemplateName"] = name.Trim(); t["TemplateName"] = SelectedCategory.Name + " - " + name.Trim(); t["RevisionCount"] = "0";
        SelectedCategory.Templates.Add(t); Select(t); modelDirty = true;
    }
    private static void Apply(Template t, Draft d)
    {
        if (d.Matches(t)) return;
        if (t.Locked && d.Locked) throw new InvalidOperationException("Unlock the working template before editing.");
        var revision = checked(int.Parse(t["RevisionCount"].Length == 0 ? "0" : t["RevisionCount"]) + 1);
        foreach (var area in Enum.GetValues<CodeArea>()) t[Template.CodeField(area)] = d.Code[area];
        t["TemplateName"] = d.Name; t["ShortTemplateName"] = d.ShortName; t["Filename"] = d.Filename;
        t["Favorite"] = d.Favorite ? "-1" : "0"; t["Locked"] = d.Locked ? "-1" : "0"; t["Undeletable"] = d.Undeletable ? "-1" : "0";
        t["RevisionCount"] = revision.ToString(); t["DateModified"] = DateTime.Now.ToString("O");
    }
    public void Update()
    {
        if (Selected == null || Current == null) return;
        if (!Current.Matches(Selected)) { Apply(Selected, Current); modelDirty = true; }
        Status = "Template updated. Save a copy to keep your changes.";
    }
    public void Save(string path)
    {
        if (SourcePath != null && string.Equals(Path.GetFullPath(path), SourcePath, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Choose a new filename to preserve the imported source.");
        // Validate and apply every draft to a separate model before writing or committing.
        var copy = new Library();
        foreach (var category in Library.Categories)
        {
            var c = new Category(); foreach (var pair in category.Fields) c.Fields.Add(pair.Key, pair.Value);
            foreach (var template in category.Templates)
            {
                var t = new Template(); foreach (var pair in template.Fields) t[pair.Key] = pair.Value;
                if (drafts.TryGetValue(template, out var d)) Apply(t, d);
                c.Templates.Add(t);
            }
            copy.Categories.Add(c);
        }
        copy.Save(path);
        var selectedId = Selected?["TemplateID"];
        Load(copy, path);
        if (copy.Templates.FirstOrDefault(t => t["TemplateID"] == selectedId) is { } selected) Select(selected);
        Status = "Saved working copy: " + path;
    }
    public void DeleteSelected()
    {
        if (Selected == null) return;
        if (Selected.Locked || Selected.Flag("Undeletable")) throw new InvalidOperationException("Template is locked or undeletable.");
        var template = Selected; SelectedCategory!.Templates.Remove(template); drafts.Remove(template); Selected = null; modelDirty = true;
        if (Library.Templates.FirstOrDefault() is { } first) Select(first);
    }
    public async Task PreviewAsync(Func<Prompt, CancellationToken, Task<string?>>? prompt = null, CancellationToken cancellation = default)
    {
        if (IsBusy) return;
        IsBusy = true;
        try
        {
            var t = Selected ?? throw new InvalidOperationException("Select a template."); var d = Current!;
            var working = new Template(); foreach (var pair in t.Fields) working[pair.Key] = pair.Value;
            foreach (var area in Enum.GetValues<CodeArea>()) working[Template.CodeField(area)] = d.Code[area]; working["Filename"] = d.Filename;
            var variables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in Variables.Replace("\r\n", "\n").Split('\n').Where(x => x.Trim().Length > 0))
            { var equal = line.IndexOf('='); if (equal < 1) throw new FormatException("Use one Name=Value per line."); variables[line[..equal].Trim()] = line[(equal + 1)..]; }
            var result = await new Engine().EvaluateAsync(Library, working, variables, prompt, cancellation);
            Preview = string.Join("\r\n\r\n", result.Outputs.Select(o => $"[{o.Area}] {o.Target}\r\n{o.Text}")); Status = "Preview complete. No output was applied.";
        }
        finally { IsBusy = false; }
    }
}
