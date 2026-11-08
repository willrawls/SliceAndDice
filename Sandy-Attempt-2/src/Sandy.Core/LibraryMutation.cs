using System.Security.Cryptography;
using System.Text;
namespace Sandy.Core;

/// <summary>Isolated template-library edits. A host must commit them after successful evaluation.</summary>
public sealed class LibraryMutation
{
    private readonly string baseline;
    private readonly HashSet<string> revised = [];
    public Library Working { get; }
    public Template Selected { get; private set; }
    public HashSet<string> ChangedTemplateIds { get; } = new(StringComparer.Ordinal);
    public bool HasChanges => Fingerprint(Working) != baseline;
    public LibraryMutation(Library library,Template selected)
    {
        baseline=Fingerprint(library);Working=library.Clone();Selected=Working.Templates.Single(t=>t.Id==selected.Id);
    }
    public void Select(string name)
    {
        name=name.Trim();if(name.Length==0)throw new FormatException("GotoSandyTemplate requires a template name.");
        var split=name.IndexOf(" - ",StringComparison.Ordinal);
        var categoryName=split<0?Working.CategoryOf(Selected):name[..split];var shortName=split<0?name:name[(split+3)..];
        var category=Working.Categories.SingleOrDefault(c=>c.Name.Equals(categoryName,StringComparison.OrdinalIgnoreCase));
        if(category==null){category=new();category.Fields["CategoryID"]=Next(Working.Categories.Select(c=>c.Fields["CategoryID"]));category.Fields["CategoryName"]=categoryName;Working.Categories.Add(category);}
        var matches=category.Templates.Where(t=>t["ShortTemplateName"].Equals(shortName,StringComparison.OrdinalIgnoreCase)||t.Name.Equals(name,StringComparison.OrdinalIgnoreCase)).ToArray();
        if(matches.Length>1)throw new InvalidOperationException("Ambiguous Sandy template: "+name);
        if(matches.Length==1){Selected=matches[0];return;}
        Selected=new();Selected["TemplateID"]=Next(Working.Templates.Select(t=>t.Id));Selected["CategoryID"]=category.Fields["CategoryID"];
        Selected["TemplateName"]=categoryName+" - "+shortName;Selected["ShortTemplateName"]=shortName;Selected["RevisionCount"]="0";
        category.Templates.Add(Selected);ChangedTemplateIds.Add(Selected.Id);
    }
    public void SetCode(string area,string text)
    {
        var target=area.Trim().ToUpperInvariant() switch {
            "0" or "ATTOP" or "AT TOP" or "TOP" or "(DECLARATIONS)" or "DECLARATIONS" or "DECL"=>CodeArea.Declarations,
            "1" or "ATCURSOR" or "AT CURSOR" or "CURSOR"=>CodeArea.Cursor,
            "2" or "ATBOTTOM" or "AT BOTTOM" or "BOTTOM"=>CodeArea.Bottom,
            "3" or "INFILE" or "IN FILE" or "FILE"=>CodeArea.File,
            _=>throw new FormatException("Unknown Sandy code area: "+area)
        };
        if(Selected.Code(target)==text)return;
        if(Selected.Locked)throw new InvalidOperationException("The target Sandy template is locked: "+Selected.Name);
        if(revised.Add(Selected.Id))Selected["RevisionCount"]=checked(int.Parse(Selected["RevisionCount"].Length==0?"0":Selected["RevisionCount"])+1).ToString();
        ChangedTemplateIds.Add(Selected.Id);
        Selected[Template.CodeField(target)]=text;Selected["DateModified"]=DateTime.Now.ToString("O");
    }
    public Library Commit(Library current)
    {
        if(Fingerprint(current)!=baseline)throw new InvalidOperationException("Library changed after evaluation. Preview again before applying library edits.");
        Working.Validate();return Working.Clone();
    }
    private static string Next(IEnumerable<string> values)=>checked(values.Select(v=>int.TryParse(v,out var n)?n:0).DefaultIfEmpty().Max()+1).ToString();
    public static string Fingerprint(Library library)=>Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(library.Serialize())));
}
