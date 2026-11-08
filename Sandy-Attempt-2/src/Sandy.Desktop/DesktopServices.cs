using System.Diagnostics;
using Microsoft.Win32;
using Sandy.Blazor;
using Sandy.Core;
namespace Sandy.Desktop;
internal sealed class DesktopServices(string? profilePath=null) : IDesktopServices
{
    private readonly ProfileStore profile=new(profilePath??System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"SandyAttempt2","preferences.json"));
    public string? ReadPreference(string key)=>profile.Read(key);
    public void WritePreference(string key,string value)=>profile.Write(key,value);
    internal Action<ShellOptions>? ShellConfiguration { get; set; }
    public void ConfigureShell(ShellOptions options)
    {
        (ShellConfiguration??throw new NotSupportedException("The desktop shell is not initialized."))(options);
        profile.Write("shell",System.Text.Json.JsonSerializer.Serialize(options));
    }
    public string? OpenLibrary()
    {
        var dialog = new OpenFileDialog { Filter = "Sandy libraries|*.xml;*.mdb|All files|*.*" };
        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }
    public string? SaveCopy()
    {
        var dialog = new SaveFileDialog { Filter = "Sandy XML library|*.xml", FileName = "Sandy-working-copy.xml", OverwritePrompt = true };
        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }
    public Library ReadLibrary(string path) => System.IO.Path.GetExtension(path).Equals(".mdb", StringComparison.OrdinalIgnoreCase) ? AccessLibrary.Read(path) : Library.Load(path);
    public string? ReadSetting(string key) => Sandy.Windows.WindowsEffects.ReadSetting(key);
    public string ReadClipboard() => Sandy.Windows.WindowsEffects.ReadClipboard();
    public void ApplyPlan(LegacyEditor plan)
    {
        Sandy.Windows.WindowsEffects.Validate(plan.Effects);
        var files=new FileCommit(plan.WorkspaceRoot.Length>0?plan.WorkspaceRoot:System.IO.Path.GetDirectoryName(plan.Documents.Values.First().Path)!);
        foreach(var document in plan.Documents.Values.Where(d=>d.Changed)) files.AddDocument(document);
        foreach(var effect in plan.Effects.Where(Sandy.Windows.WindowsEffects.IsFile)) {
            if(plan.Documents.TryGetValue(System.IO.Path.GetFullPath(effect.Target),out var document)&&document.Changed)
                throw new InvalidOperationException("File output conflicts with an edited document: "+effect.Target);
            files.AddEffect(effect);
        }
        files.Commit();
        foreach(var effect in DeferredActions.Apply(plan.Effects.Where(e=>!Sandy.Windows.WindowsEffects.IsFile(e)),Sandy.Windows.WindowsEffects.Apply)) {
            if(effect.Kind=="Message")System.Windows.MessageBox.Show(effect.Value,"Sandy");
        }
    }
    public IReadOnlyDictionary<string, string[]> ReadSchema(string path) => AccessLibrary.Schema(path);
    public IReadOnlyDictionary<string,string[]> ReadOdbcSchema(string connectionString)
    {
        using var connection=new System.Data.Odbc.OdbcConnection(connectionString);connection.Open();
        var tables=connection.GetSchema("Tables");var result=new Dictionary<string,string[]>();
        foreach(System.Data.DataRow row in tables.Rows)
        {
            var name=Convert.ToString(row["TABLE_NAME"])!;var schema=Convert.ToString(row["TABLE_SCHEM"])??"";
            if(!string.Equals(Convert.ToString(row["TABLE_TYPE"]),"TABLE",StringComparison.OrdinalIgnoreCase)&&!string.Equals(Convert.ToString(row["TABLE_TYPE"]),"VIEW",StringComparison.OrdinalIgnoreCase))continue;
            var columns=connection.GetSchema("Columns",[null,schema,name,null]);var fields=new List<string>();
            foreach(System.Data.DataRow column in columns.Rows)
            {
                var sqlType=Convert.ToInt32(column["DATA_TYPE"]);var daoType=sqlType switch{-7=>1,-6=>2,5=>3,4=>4,2 or 3=>5,6 or 7=>6,8=>7,9 or 91 or 92 or 93=>8,-2 or -3 or -4=>9,1 or 12 or -1 or -8 or -9 or -10=>10,-11=>15,-5=>16,_=>0};
                var length=column["COLUMN_SIZE"] is DBNull?0:Convert.ToInt64(column["COLUMN_SIZE"]);
                fields.Add($"{column["COLUMN_NAME"]} : DAO {daoType} ({length})");
            }
            result[schema.Length==0?name:schema+"."+name]=fields.ToArray();
        }
        return result;
    }
    public string? PickFile(string filter) { var dialog=new OpenFileDialog { Filter=filter }; return dialog.ShowDialog()==true?dialog.FileName:null; }
    public string? PickFolder() { var dialog=new OpenFolderDialog { Title="Choose output folder" };return dialog.ShowDialog()==true?dialog.FolderName:null; }
    public string? SaveText(string suggestedName) { var extension=System.IO.Path.GetExtension(suggestedName);var dialog=new SaveFileDialog { FileName=suggestedName, Filter="Output file|*"+extension+"|All files|*.*", OverwritePrompt=true };return dialog.ShowDialog()==true?dialog.FileName:null; }
    public void OpenUrl(string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || uri.Scheme is not ("http" or "https")) throw new ArgumentException("Enter an HTTP or HTTPS address.");
        Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true });
    }
}
