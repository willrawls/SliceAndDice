using System.Globalization;

namespace Sandy.Core;

public static class TemplateEditor
{
    public static bool IsLocked(SandyTemplate template)
        => template.Get("Locked") == "1" || template.Get("Locked") == "-1" || template.Get("Locked").Equals("True", StringComparison.OrdinalIgnoreCase);

    /// <summary>Replace code in memory after explicit Save, retaining all other historical fields.</summary>
    public static void UpdateCode(SandyTemplate template, IReadOnlyDictionary<TemplateArea, string> code, string filename, DateTime modified)
    {
        if (IsLocked(template)) throw new InvalidOperationException("This template is locked. Unlock a working copy before editing.");
        var changed = code.Any(c => template.GetCode(c.Key) != c.Value) || template.Get("Filename") != filename;
        if (!changed) return;
        var revision = int.TryParse(template.Get("RevisionCount"), out var existing) ? existing : 1;
        var nextRevision = checked(revision + 1);
        foreach (var pair in code) template.Fields[SandyTemplate.FieldFor(pair.Key)] = pair.Value;
        template.Fields["Filename"] = filename;
        template.Fields["RevisionCount"] = nextRevision.ToString(CultureInfo.InvariantCulture);
        template.Fields["DateModified"] = modified.ToString("O", CultureInfo.InvariantCulture);
    }
}
