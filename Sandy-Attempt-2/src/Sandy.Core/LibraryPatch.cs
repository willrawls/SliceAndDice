using System.Text.Json;
using System.Xml.Linq;
namespace Sandy.Core;

public sealed record LibraryRecordChange(string Id,Dictionary<string,string>? Fields);
public sealed class LibraryPatch
{
    public int FormatVersion { get; set; }=1;
    public string BaseHash { get; set; }="";
    public string TargetHash { get; set; }="";
    public List<LibraryRecordChange> Categories { get; set; }=[];
    public List<LibraryRecordChange> Templates { get; set; }=[];
    public List<string> CategoryOrder { get; set; }=[];
    public Dictionary<string,List<string>> TemplateOrder { get; set; }=[];
    public List<string> AuxiliaryRecords { get; set; }=[];
    public static LibraryPatch Create(Library before,Library after)
    {
        before.Validate();after.Validate();
        return new(){BaseHash=LibraryMutation.Fingerprint(before),TargetHash=LibraryMutation.Fingerprint(after),
            Categories=Diff(before.Categories.ToDictionary(c=>c.Fields["CategoryID"],c=>c.Fields),after.Categories.ToDictionary(c=>c.Fields["CategoryID"],c=>c.Fields)),
            Templates=Diff(before.Templates.ToDictionary(t=>t.Id,t=>t.Fields),after.Templates.ToDictionary(t=>t.Id,t=>t.Fields)),
            CategoryOrder=after.Categories.Select(c=>c.Fields["CategoryID"]).ToList(),TemplateOrder=after.Categories.ToDictionary(c=>c.Fields["CategoryID"],c=>c.Templates.Select(t=>t.Id).ToList()),AuxiliaryRecords=after.AuxiliaryRecords.Select(e=>e.ToString(SaveOptions.DisableFormatting)).ToList()};
    }
    private static List<LibraryRecordChange> Diff(Dictionary<string,Dictionary<string,string>> before,Dictionary<string,Dictionary<string,string>> after)
        =>before.Keys.Union(after.Keys).Where(id=>!before.TryGetValue(id,out var b)||!after.TryGetValue(id,out var a)||b.Count!=a.Count||b.Any(p=>!a.TryGetValue(p.Key,out var value)||value!=p.Value))
            .Select(id=>new LibraryRecordChange(id,after.TryGetValue(id,out var fields)?new(fields,StringComparer.OrdinalIgnoreCase):null)).ToList();
    public Library Apply(Library before)
    {
        if(FormatVersion!=1)throw new InvalidDataException("Unsupported Sandy patch version.");
        if(LibraryMutation.Fingerprint(before)!=BaseHash)throw new InvalidOperationException("The patch base does not match this working library.");
        if(Categories.Select(c=>c.Id).Distinct().Count()!=Categories.Count||Templates.Select(t=>t.Id).Distinct().Count()!=Templates.Count)throw new InvalidDataException("Duplicate patch records.");
        var result=before.Clone();var categories=result.Categories.ToDictionary(c=>c.Fields["CategoryID"]);var templates=result.Templates.ToDictionary(t=>t.Id);
        foreach(var change in Templates)
        {
            if(templates.TryGetValue(change.Id,out var old))
            {
                if(change.Fields==null&&(old.Locked||old.Flag("Undeletable")))throw new InvalidOperationException("Patch would remove a protected template: "+old.Name);
                if(change.Fields!=null&&old.Locked&&change.Fields.GetValueOrDefault("Locked") is not ("0" or "false" or "False" or ""))throw new InvalidOperationException("Patch would edit a locked template: "+old.Name);
            }
            if(change.Fields==null){templates.Remove(change.Id);continue;}
            var t=new Template();foreach(var pair in change.Fields)t[pair.Key]=pair.Value;if(t.Id!=change.Id)throw new InvalidDataException("Template patch identity mismatch.");templates[change.Id]=t;
        }
        foreach(var change in Categories)
        {
            if(change.Fields==null){categories.Remove(change.Id);continue;}
            var c=new Category();foreach(var pair in change.Fields)c.Fields[pair.Key]=pair.Value;if(c.Fields.GetValueOrDefault("CategoryID")!=change.Id)throw new InvalidDataException("Category patch identity mismatch.");categories[change.Id]=c;
        }
        if(CategoryOrder.Distinct().Count()!=categories.Count||CategoryOrder.Count!=categories.Count||CategoryOrder.Any(id=>!categories.ContainsKey(id)))throw new InvalidDataException("Invalid category order in patch.");
        result.Categories.Clear();var placed=new HashSet<string>();
        foreach(var id in CategoryOrder)
        {
            var c=categories[id];c.Templates.Clear();
            if(!TemplateOrder.TryGetValue(id,out var order))throw new InvalidDataException("Missing template order in patch.");
            foreach(var templateId in order){if(!templates.TryGetValue(templateId,out var t)||t["CategoryID"]!=id||!placed.Add(templateId))throw new InvalidDataException("Invalid template ownership/order in patch.");c.Templates.Add(t);}
            result.Categories.Add(c);
        }
        if(placed.Count!=templates.Count)throw new InvalidDataException("Patch leaves orphan templates.");
        result.AuxiliaryRecords.Clear();result.AuxiliaryRecords.AddRange(AuxiliaryRecords.Select(xml=>XElement.Parse(xml)));result.Validate();
        if(LibraryMutation.Fingerprint(result)!=TargetHash)throw new InvalidDataException("Patch result checksum mismatch.");return result;
    }
    public string ToJson()=>JsonSerializer.Serialize(this,new JsonSerializerOptions{WriteIndented=true});
    public static LibraryPatch FromJson(string text)=>JsonSerializer.Deserialize<LibraryPatch>(text)??throw new InvalidDataException("Missing patch.");
}
