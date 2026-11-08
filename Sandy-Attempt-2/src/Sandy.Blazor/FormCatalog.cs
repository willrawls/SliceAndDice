using System.Text.Json;
namespace Sandy.Blazor;

public sealed class FormDefinition
{
    public string Type { get; set; } = "";
    public string Name { get; set; } = "";
    public string ClassName { get; set; } = "";
    public string Source { get; set; } = "";
    public Dictionary<string, JsonElement> Properties { get; set; } = [];
    public List<FormDefinition> Children { get; set; } = [];
    public string Get(string key, string fallback = "") => Properties.TryGetValue(key, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString()! : fallback;
    public bool Flag(string key, bool fallback = false) => int.TryParse(Get(key), out var value) ? value != 0 : fallback;
    public string Key => Name + (Properties.ContainsKey("Index") ? "_" + Get("Index") : "");
    public string Caption => Get("Caption", Name).Replace("&", "");
    public IEnumerable<FormDefinition> Descendants => Children.SelectMany(x => new[] { x }.Concat(x.Descendants));
    public string[] Items => Properties.TryGetValue("Items", out var items) ? items.EnumerateArray().Select(x => x.GetString() ?? "").ToArray() : [];
}
public static class FormCatalog
{
    public static IReadOnlyList<FormDefinition> All { get; } = Load();
    private static List<FormDefinition> Load()
    {
        using var stream = typeof(FormCatalog).Assembly.GetManifestResourceStream("Sandy.Blazor.catalog.json")!;
        return JsonSerializer.Deserialize<List<FormDefinition>>(stream, new JsonSerializerOptions { PropertyNameCaseInsensitive = true })!;
    }
}
