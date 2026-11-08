using System.Text;
using System.Text.Json;

namespace Sandy.Core;

/// <summary>Preflight, journal and rollback for the file part of an evaluated plan.</summary>
public sealed class FileCommit
{
    private readonly string root;
    private readonly Dictionary<string,byte[]?> before=new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string,byte[]?> after=new(StringComparer.OrdinalIgnoreCase);
    private readonly List<string> createdDirectories=[];
    private readonly List<string> directoriesToDelete=[];
    public string? JournalPath { get; private set; }
    public FileCommit(string root) => this.root=Path.TrimEndingDirectorySeparator(Path.GetFullPath(root));
    private string Resolve(string path)
    {
        var full=Path.GetFullPath(Path.IsPathRooted(path)?path:Path.Combine(root,path));
        if(!full.StartsWith(root+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("Output must remain inside the selected workspace: "+full);
        for(var parent=Path.GetDirectoryName(full);parent!=null&&parent.Length>=root.Length;parent=Path.GetDirectoryName(parent))
            if(Directory.Exists(parent)&&new DirectoryInfo(parent).LinkTarget!=null)throw new InvalidOperationException("Output through a linked directory is not supported: "+parent);
        if(File.Exists(full)&&new FileInfo(full).LinkTarget!=null)throw new InvalidOperationException("Output through a linked file is not supported.");
        return full;
    }
    private byte[]? Read(string path)
    {
        if(!before.ContainsKey(path))before[path]=File.Exists(path)?File.ReadAllBytes(path):null;
        return after.TryGetValue(path,out var data)?data:before[path];
    }
    public void AddDocument(PlannedDocument document)
    {
        var path=Resolve(document.Path);var bytes=Read(path);
        if((File.Exists(path)?TextFiles.Read(path):"")!=document.Original)throw new InvalidOperationException("File changed since evaluation: "+path);
        after[path]=TextFiles.Encode(document.Text,bytes);
    }
    public void AddEffect(ExternalEffect effect)
    {
        var path=Resolve(effect.Target);
        switch(effect.Kind)
        {
            case "FileAppend":var old=Read(path)??[];after[path]=[..old,..Encoding.UTF8.GetBytes(effect.Value)];break;
            case "FileOverwrite":Read(path);after[path]=effect.Data??Encoding.UTF8.GetBytes(effect.Value);break;
            case "DeleteFile":Read(path);after[path]=null;break;
            case "Rename":var destination=Resolve(effect.Value);if(Read(destination)!=null)throw new IOException("Rename destination already exists.");after[destination]=Read(path)??throw new FileNotFoundException("Rename source does not exist.",path);after[path]=null;break;
            case "CreateDirectory":if(!Directory.Exists(path))createdDirectories.Add(path);break;
            case "DeleteDirectory":if(Directory.Exists(path)&&Directory.EnumerateFileSystemEntries(path).Any())throw new IOException("Directory must be empty: "+path);directoriesToDelete.Add(path);break;
            default:throw new NotSupportedException("Not a file effect: "+effect.Kind);
        }
    }
    public void Commit()
    {
        foreach(var pair in before)
        {
            var current=File.Exists(pair.Key)?File.ReadAllBytes(pair.Key):null;
            if(current==null?pair.Value!=null:pair.Value==null||!current.SequenceEqual(pair.Value))throw new IOException("Output changed before commit: "+pair.Key);
        }
        if(after.Count==0&&createdDirectories.Count==0&&directoriesToDelete.Count==0)return;
        JournalPath=Path.Combine(root,".sandy-history",DateTime.UtcNow.ToString("yyyyMMddTHHmmssfff")+"-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(JournalPath);
        var entries=new List<object>();var index=0;
        foreach(var pair in before){var backup=pair.Value==null?null:(index++).ToString()+".bin";if(backup!=null)File.WriteAllBytes(Path.Combine(JournalPath,backup),pair.Value!);entries.Add(new{Path=pair.Key,Backup=backup});}
        File.WriteAllText(Path.Combine(JournalPath,"journal.json"),JsonSerializer.Serialize(entries,new JsonSerializerOptions{WriteIndented=true}));
        try
        {
            foreach(var path in createdDirectories.ToArray())Directory.CreateDirectory(path);
            foreach(var pair in after)
            {
                if(pair.Value==null){if(File.Exists(pair.Key))File.Delete(pair.Key);continue;}
                var parent=Path.GetDirectoryName(pair.Key)!;if(!Directory.Exists(parent)){Directory.CreateDirectory(parent);createdDirectories.Add(parent);}
                var temporary=pair.Key+".sandy-"+Guid.NewGuid().ToString("N")+".tmp";
                try{File.WriteAllBytes(temporary,pair.Value);File.Move(temporary,pair.Key,true);}finally{if(File.Exists(temporary))File.Delete(temporary);}
            }
            foreach(var directory in directoriesToDelete)if(Directory.Exists(directory))Directory.Delete(directory,false);
        }
        catch{Rollback();throw;}
    }
    public void Rollback()
    {
        foreach(var pair in before){if(pair.Value==null){if(File.Exists(pair.Key))File.Delete(pair.Key);}else{Directory.CreateDirectory(Path.GetDirectoryName(pair.Key)!);File.WriteAllBytes(pair.Key,pair.Value);}}
        foreach(var directory in directoriesToDelete)Directory.CreateDirectory(directory);
        foreach(var directory in createdDirectories.OrderByDescending(x=>x.Length))if(Directory.Exists(directory)&&!Directory.EnumerateFileSystemEntries(directory).Any())Directory.Delete(directory,false);
    }
}
