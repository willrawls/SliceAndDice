using System.IO;
using System.Runtime.InteropServices;
using Sandy.Core;

namespace Sandy.Desktop;

/// <summary>Late-bound, read-only DAO import. No registration or writes to historical MDBs.</summary>
internal static class AccessLibrary
{
    private static dynamic CreateEngine()
    {
        var type = Type.GetTypeFromProgID("DAO.DBEngine.120") ?? Type.GetTypeFromProgID("DAO.DBEngine.36")
            ?? throw new NotSupportedException("MDB import requires a matching-bitness Microsoft Access Database Engine (DAO). Export Category and its related Template table to Category.xml with Access, or open an existing Category.xml export.");
        return Activator.CreateInstance(type)!;
    }
    private static void Release(object? value) { if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value); }
    private static Dictionary<string, string> Fields(dynamic row)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        dynamic fields = row.Fields;
        try
        {
            for (var i = 0; i < (int)fields.Count; i++)
            {
                dynamic field = fields.Item(i);
                try { result[(string)field.Name] = field.Value is DBNull ? "" : Convert.ToString(field.Value, System.Globalization.CultureInfo.InvariantCulture) ?? ""; }
                finally { Release(field); }
            }
        }
        finally { Release(fields); }
        return result;
    }
    public static Library Read(string filename)
    {
        try { return ReadNative(filename); }
        catch (COMException ex) when (ex.HResult is unchecked((int)0x800A0BE1) or unchecked((int)0x80040154))
        {
            // Jet 3/Access 97 files need 32-bit DAO 3.6, even when ACE is installed.
            var powershell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SysWOW64", "WindowsPowerShell", "v1.0", "powershell.exe");
            var script = Path.Combine(AppContext.BaseDirectory, "Export-LegacyMdb.ps1");
            if (!File.Exists(powershell) || !File.Exists(script)) throw;
            var xml = Path.Combine(Path.GetTempPath(), "Sandy-" + Guid.NewGuid().ToString("N") + ".xml");
            try
            {
                var start = new System.Diagnostics.ProcessStartInfo(powershell) { UseShellExecute = false, CreateNoWindow = true, WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden, RedirectStandardError = true };
                foreach (var arg in new[] { "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script, "-InputFile", Path.GetFullPath(filename), "-OutputFile", xml }) start.ArgumentList.Add(arg);
                using var process = System.Diagnostics.Process.Start(start)!;
                var error = process.StandardError.ReadToEnd(); process.WaitForExit();
                if (process.ExitCode != 0) throw new InvalidOperationException("Legacy MDB import failed: " + error, ex);
                return Library.Load(xml);
            }
            finally { if (File.Exists(xml)) File.Delete(xml); }
        }
    }
    private static Library ReadNative(string filename)
    {
        dynamic engine = CreateEngine(); dynamic? database = null, categories = null, templates = null;
        try
        {
            database = engine.OpenDatabase(Path.GetFullPath(filename), false, true);
            categories = database.OpenRecordset("SELECT * FROM [Category]", 4);
            templates = database.OpenRecordset("SELECT * FROM [Template]", 4);
            var result = new Library();
            while (!categories.EOF)
            {
                var category = new Category(); foreach (var pair in Fields(categories)) category.Fields.Add(pair.Key, pair.Value);
                result.Categories.Add(category); categories.MoveNext();
            }
            while (!templates.EOF)
            {
                var template = new Template(); foreach (var pair in Fields(templates)) template[pair.Key] = pair.Value;
                var category = result.Categories.SingleOrDefault(c => c.Fields.GetValueOrDefault("CategoryID") == template["CategoryID"])
                    ?? throw new InvalidDataException("Orphan template: " + template.Name);
                category.Templates.Add(template); templates.MoveNext();
            }
            dynamic tableDefs=database.TableDefs;
            try {
                for(var i=0;i<(int)tableDefs.Count;i++) {
                    dynamic table=tableDefs.Item(i);
                    try {
                        string name=table.Name;if(name.StartsWith("MSys",StringComparison.OrdinalIgnoreCase)||name is "Category" or "Template")continue;
                        dynamic rows=database.OpenRecordset("SELECT * FROM ["+name.Replace("]","]]")+"]",4);
                        try { while(!rows.EOF) { var record=new System.Xml.Linq.XElement(System.Xml.XmlConvert.EncodeLocalName(name));foreach(var pair in Fields(rows))record.Add(new System.Xml.Linq.XElement(System.Xml.XmlConvert.EncodeLocalName(pair.Key),pair.Value));result.AuxiliaryRecords.Add(record);rows.MoveNext(); } }
                        finally{rows.Close();Release(rows);}
                    }
                    finally{Release(table);}
                }
            }
            finally{Release(tableDefs);}
            result.Validate(); return result;
        }
        finally
        {
            if (templates != null) { templates.Close(); Release(templates); }
            if (categories != null) { categories.Close(); Release(categories); }
            if (database != null) { database.Close(); Release(database); }
            Release(engine);
        }
    }
    public static Dictionary<string, string[]> Schema(string filename)
    {
        try { return SchemaNative(filename); }
        catch (COMException ex) when (ex.HResult is unchecked((int)0x800A0BE1) or unchecked((int)0x80040154))
        {
            var output = Path.Combine(Path.GetTempPath(),"Sandy-schema-"+Guid.NewGuid().ToString("N")+".json");
            try
            {
                var start = new System.Diagnostics.ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),"SysWOW64","WindowsPowerShell","v1.0","powershell.exe")) { UseShellExecute=false, CreateNoWindow=true, WindowStyle=System.Diagnostics.ProcessWindowStyle.Hidden, RedirectStandardError=true };
                foreach(var arg in new[]{"-NoProfile","-NonInteractive","-ExecutionPolicy","Bypass","-File",Path.Combine(AppContext.BaseDirectory,"Export-LegacyMdb.ps1"),"-InputFile",Path.GetFullPath(filename),"-OutputFile",output,"-SchemaOnly"})start.ArgumentList.Add(arg);
                using var process=System.Diagnostics.Process.Start(start)!;var error=process.StandardError.ReadToEnd();process.WaitForExit();if(process.ExitCode!=0)throw new InvalidOperationException("Legacy schema read failed: "+error,ex);
                return System.Text.Json.JsonSerializer.Deserialize<Dictionary<string,string[]>>(File.ReadAllText(output))!;
            }
            finally { if(File.Exists(output))File.Delete(output); }
        }
    }
    private static Dictionary<string, string[]> SchemaNative(string filename)
    {
        dynamic engine = CreateEngine(); dynamic? database = null, tables = null;
        try
        {
            database = engine.OpenDatabase(Path.GetFullPath(filename), false, true); tables = database.TableDefs;
            var result = new Dictionary<string, string[]>();
            for (var i = 0; i < (int)tables.Count; i++)
            {
                dynamic table = tables.Item(i); dynamic? fields = null;
                try
                {
                    string name = table.Name; if (name.StartsWith("MSys")) continue;
                    fields = table.Fields; var names = new List<string>();
                    for (var j = 0; j < (int)fields.Count; j++) { dynamic field = fields.Item(j); try { names.Add($"{field.Name} : DAO {field.Type} ({field.Size})"); } finally { Release(field); } }
                    result.Add(name, names.ToArray());
                }
                finally { Release(fields); Release(table); }
            }
            return result;
        }
        finally { Release(tables); if (database != null) { database.Close(); Release(database); } Release(engine); }
    }
}
