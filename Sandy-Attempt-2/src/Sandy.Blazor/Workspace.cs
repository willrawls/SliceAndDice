using Sandy.Core;
namespace Sandy.Blazor;

public interface IDesktopServices
{
    string? OpenLibrary();
    string? SaveCopy();
    Library ReadLibrary(string path);
    void OpenUrl(string url);
    IReadOnlyDictionary<string, string[]> ReadSchema(string path) => throw new NotSupportedException("Database schema inspection is not available in this host.");
    string? PickFile(string filter) => OpenLibrary();
    string? SaveText(string suggestedName) => SaveCopy();
    IReadOnlyDictionary<string,string[]> ReadOdbcSchema(string connectionString) => throw new NotSupportedException("ODBC is not available in this host.");
    string? ReadSetting(string key) => null;
    string ReadClipboard() => "";
    void ApplyPlan(LegacyEditor plan) => throw new NotSupportedException("This host cannot apply execution plans.");
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
    public string EditorPath { get; set; } = Path.Combine(Environment.CurrentDirectory, "SandyPreview.cs");
    public string EditorText { get; set; } = "";
    public int Cursor { get; set; }
    public int SelectionLength { get; set; }
    public LegacyEditor? PreviewPlan { get; private set; }
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
    public void SelectCategory(Category category) { if(!Library.Categories.Contains(category))throw new InvalidOperationException("Category belongs to another library.");SelectedCategory=category; }
    public void RenameCategory(string name)
    {
        var category=SelectedCategory??throw new InvalidOperationException("Select a category.");
        if(string.IsNullOrWhiteSpace(name))throw new ArgumentException("Enter a category name.");
        var old=category.Name;
        if(category.Templates.Any(t=>(drafts.TryGetValue(t,out var d)?d.Locked:t.Locked)&&
            (drafts.TryGetValue(t,out var existing)?existing.Name:t.Name).StartsWith(old+" - ",StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException("Unlock templates before renaming their category.");
        category.Fields["CategoryName"]=name.Trim();
        foreach(var t in category.Templates) {
            if(!drafts.TryGetValue(t,out var d))drafts[t]=d=Draft.From(t);
            if(d.Name.StartsWith(old+" - ",StringComparison.OrdinalIgnoreCase))d.Name=name.Trim()+d.Name[old.Length..];
        }
        modelDirty=true;
    }
    public void DeleteCategory()
    {
        var category=SelectedCategory??throw new InvalidOperationException("Select a category.");
        if(category.Templates.Any(t=>t.Locked||t.Flag("Undeletable")||drafts.TryGetValue(t,out var d)&&(d.Locked||d.Undeletable)))throw new InvalidOperationException("Category contains locked or undeletable templates.");
        foreach(var t in category.Templates)drafts.Remove(t);
        Library.Categories.Remove(category);SelectedCategory=Library.Categories.FirstOrDefault();Selected=null;modelDirty=true;
        if(SelectedCategory?.Templates.FirstOrDefault() is {} first)Select(first);
    }
    public void DuplicateTemplate()
    {
        var original=Selected??throw new InvalidOperationException("Select a template.");var draft=Current!;
        var clone=new Template();foreach(var pair in original.Fields)clone[pair.Key]=pair.Value;
        clone["TemplateID"]=NextId(Library.Templates.Select(t=>t.Id));clone["Locked"]="0";clone["Undeletable"]="0";
        clone["TemplateName"]=draft.Name+" (copy)";clone["ShortTemplateName"]=draft.ShortName+" (copy)";clone["Filename"]=draft.Filename;
        foreach(var area in Enum.GetValues<CodeArea>())clone[Template.CodeField(area)]=draft.Code[area];
        Library.Categories.Single(c=>c.Templates.Contains(original)).Templates.Add(clone);Select(clone);modelDirty=true;
    }
    public IReadOnlyList<SearchHit> Search(string find,string replacement="",string scope="Library",bool matchCase=false,bool wholeWord=false,bool pattern=false,bool replace=false)
    {
        if(find.Length==0)throw new ArgumentException("Enter text to find.");
        var expression=pattern?find:System.Text.RegularExpressions.Regex.Escape(find);
        if(wholeWord)expression=@"\b(?:"+expression+@")\b";
        var regex=new System.Text.RegularExpressions.Regex(expression,matchCase?System.Text.RegularExpressions.RegexOptions.None:System.Text.RegularExpressions.RegexOptions.IgnoreCase,TimeSpan.FromMilliseconds(100));
        var templates=scope=="Template"?(Selected==null?Array.Empty<Template>():new[]{Selected}):scope=="Category"?SelectedCategory?.Templates??[]:Library.Templates;
        var hits=new List<SearchHit>();var changes=new List<(Draft Draft,CodeArea Area,string Text)>();
        foreach(var template in templates)
        {
            if(!drafts.TryGetValue(template,out var draft))drafts[template]=draft=Draft.From(template);
            foreach(var area in Enum.GetValues<CodeArea>())
            {
                var text=draft.Code[area];var matches=regex.Matches(text);if(matches.Count==0)continue;
                foreach(System.Text.RegularExpressions.Match match in matches) { if(hits.Count>=10000)throw new InvalidOperationException("Search exceeds 10,000 matches; narrow the scope."); hits.Add(new(template,area,match.Index,match.Value)); }
                if(replace) {
                    if(draft.Locked)throw new InvalidOperationException("Unlock matching template before replacing: "+template.Name);
                    changes.Add((draft,area,pattern?regex.Replace(text,replacement):regex.Replace(text,_=>replacement)));
                }
            }
        }
        foreach(var change in changes)change.Draft.Code[change.Area]=change.Text;
        Status=$"{hits.Count} occurrence(s) {(replace?"replaced":"found")}.";return hits;
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
        copy.AuxiliaryRecords.AddRange(Library.AuxiliaryRecords.Select(e=>new System.Xml.Linq.XElement(e)));
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
        if (Selected.Locked || Selected.Flag("Undeletable") || Current is { Locked:true } or { Undeletable:true }) throw new InvalidOperationException("Template is locked or undeletable.");
        var template = Selected; SelectedCategory!.Templates.Remove(template); drafts.Remove(template); Selected = null; modelDirty = true;
        if (Library.Templates.FirstOrDefault() is { } first) Select(first);
    }
    public async Task PreviewAsync(Func<Prompt, CancellationToken, Task<string?>>? prompt = null, CancellationToken cancellation = default, IDesktopServices? services = null)
    {
        if (IsBusy) return;
        IsBusy = true;
        PreviewPlan = null; Preview = "";
        try
        {
            var t = Selected ?? throw new InvalidOperationException("Select a template."); var d = Current!;
            var working = new Template(); foreach (var pair in t.Fields) working[pair.Key] = pair.Value;
            foreach (var area in Enum.GetValues<CodeArea>()) working[Template.CodeField(area)] = d.Code[area]; working["Filename"] = d.Filename;
            var variables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in Variables.Replace("\r\n", "\n").Split('\n').Where(x => x.Trim().Length > 0))
            { var equal = line.IndexOf('='); if (equal < 1) throw new FormatException("Use one Name=Value per line."); variables[line[..equal].Trim()] = line[(equal + 1)..]; }
            PreviewPlan = null;
            var editor = new LegacyEditor(Path.GetFullPath(EditorPath), EditorText, Cursor, SelectionLength) {
                FileReader = TextFiles.Read,
                FetchAsync = WebContent.FetchAsync,
                RegistryReader = services == null ? null : services.ReadSetting,
                ClipboardReader = services == null ? null : services.ReadClipboard,
                DirectoryReader = pattern => string.Join(";", Directory.GetFileSystemEntries(Path.GetDirectoryName(pattern) is {Length:>0} dir ? dir : Path.GetDirectoryName(EditorPath)!, Path.GetFileName(pattern))),
                AskAsync = prompt == null ? null : (question, fallback, ct) => prompt(new Prompt(question, fallback), ct)
            };
            if(LegacyEditor.FindProject(EditorPath) is {} project)editor.AttachProject(project,TextFiles.Read(project));
            var result = await new Engine().EvaluateAsync(Library, working, variables, prompt, cancellation, editor);
            Preview = string.Join("\r\n\r\n", editor.Documents.Values.Where(d => d.Changed).Select(d => $"[Document] {d.Path}\r\n{d.Text}")
                .Concat(editor.Effects.Select(e => $"[{e.Kind}] {e.Target}\r\n{e.Value}")));
            if (Preview.Length == 0) Preview = string.Join("\r\n\r\n", result.Outputs.Select(o => $"[{o.Area}] {o.Target}\r\n{o.Text}"));
            PreviewPlan = editor; Status = "Preview complete. No output was applied.";
        }
        finally { IsBusy = false; }
    }
    public void CompleteApplication(LegacyEditor plan)
    {
        if(!ReferenceEquals(PreviewPlan,plan))throw new InvalidOperationException("Preview plan was replaced.");
        EditorText=plan.Documents.Values.First().Text;Cursor=Math.Min(Cursor,EditorText.Length);SelectionLength=0;PreviewPlan=null;
        Status="Applied plan. File backups are in .sandy-history beside the project/source file.";
    }
}
public sealed record SearchHit(Template Template,CodeArea Area,int Offset,string Text);
