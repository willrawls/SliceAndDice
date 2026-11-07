using System.Text;

namespace Sandy.Core;

/// <summary>Explicit Access 18-column export profile; never guesses field order from record width.</summary>
public static class LibraryCsv
{
    private static readonly string[] CategoryFields = { "CategoryID", "CategoryName", "CategoryType", "View", "Arrange", "ColumnWidth", "DateCreated", "DateModified", "memoAttributes" };
    private static readonly string[] TemplateFields = { "TemplateID", "CategoryID", "TemplateName", "ShortTemplateName", "Filename", "Undeletable", "Locked", "IncludeInMenu", "memoCodeAtCursor", "memoCodeAtTop", "memoCodeAtBottom", "memoCodeToFile", "DateCreated", "DateModified", "memoAttributes", "Favorite", "RevisionCount", "TimerInsertion" };

    public static SandyLibrary Read(TextReader categories, TextReader templates, string name = "Imported")
    {
        var library = new SandyLibrary { Name = name };
        foreach (var row in ReadRecords(categories))
        {
            if (row.Count != CategoryFields.Length) throw new InvalidDataException($"Expected 9 category columns; found {row.Count}.");
            var category = new SandyCategory();
            for (var i = 0; i < row.Count; i++) category.Fields[CategoryFields[i]] = row[i];
            if (category.Id.Length == 0 || library.Categories.Any(c => c.Id == category.Id)) throw new InvalidDataException("Missing or duplicate category ID.");
            library.Categories.Add(category);
        }
        foreach (var row in ReadRecords(templates))
        {
            if (row.Count != TemplateFields.Length) throw new InvalidDataException($"Expected the legacy 18-column template profile; found {row.Count}. Use Category.xml or an explicit verified schema for newer exports.");
            var template = new SandyTemplate();
            for (var i = 0; i < row.Count; i++) template.Fields[TemplateFields[i]] = row[i];
            var category = library.Categories.SingleOrDefault(c => c.Id == template.CategoryId)
                ?? throw new InvalidDataException($"Missing category {template.CategoryId} for template {template.Id}.");
            if (template.Id.Length == 0 || library.Templates.Any(t => t.Id == template.Id)) throw new InvalidDataException("Missing or duplicate template ID.");
            category.Templates.Add(template);
        }
        return library;
    }

    private static IEnumerable<List<string>> ReadRecords(TextReader reader)
    {
        var record = new List<string>();
        var field = new StringBuilder();
        var quoted = false;
        var closedQuote = false;
        int next;
        while ((next = reader.Read()) >= 0)
        {
            var c = (char)next;
            if (quoted)
            {
                if (c != '"') field.Append(c);
                else if (reader.Peek() == '"') { reader.Read(); field.Append('"'); }
                else { quoted = false; closedQuote = true; }
                continue;
            }
            if (closedQuote && c != ',' && c != '\r' && c != '\n') throw new InvalidDataException("Unexpected character after CSV closing quote.");
            if (c == '"')
            {
                if (field.Length != 0) throw new InvalidDataException("Unexpected quote inside unquoted CSV field.");
                quoted = true;
            }
            else if (c == ',') { record.Add(field.ToString()); field.Clear(); closedQuote = false; }
            else if (c == '\r' || c == '\n')
            {
                if (c == '\r' && reader.Peek() == '\n') reader.Read();
                record.Add(field.ToString()); field.Clear(); closedQuote = false;
                yield return record; record = new List<string>();
            }
            else field.Append(c);
        }
        if (quoted) throw new InvalidDataException("Unclosed CSV quoted field.");
        if (field.Length > 0 || record.Count > 0 || closedQuote) { record.Add(field.ToString()); yield return record; }
    }
}
