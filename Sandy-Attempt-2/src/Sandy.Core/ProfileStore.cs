using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
namespace Sandy.Core;

public sealed class ProfileStore(string path)
{
    public string Path { get; }=System.IO.Path.GetFullPath(path);
    private T WithLock<T>(Func<T> action)
    {
        using var mutex=new Mutex(false,"SandyProfile-"+Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(Path.ToUpperInvariant())))[..24]);
        bool acquired;try{acquired=mutex.WaitOne(TimeSpan.FromSeconds(5));}catch(AbandonedMutexException){acquired=true;}
        if(!acquired)throw new IOException("Sandy settings are busy in another process.");try{return action();}finally{mutex.ReleaseMutex();}
    }
    private Dictionary<string,string> ReadAll()=>File.Exists(Path)?JsonSerializer.Deserialize<Dictionary<string,string>>(File.ReadAllText(Path))??new():new();
    public string? Read(string key)=>WithLock(()=>ReadAll().GetValueOrDefault(key));
    public void Write(string key,string value)=>WithLock(()=> {
        var values=ReadAll();values[key]=value;Directory.CreateDirectory(System.IO.Path.GetDirectoryName(Path)!);var temporary=Path+"."+Guid.NewGuid().ToString("N")+".tmp";
        try{File.WriteAllText(temporary,JsonSerializer.Serialize(values,new JsonSerializerOptions{WriteIndented=true}),new UTF8Encoding(false));File.Move(temporary,Path,true);}finally{if(File.Exists(temporary))File.Delete(temporary);}return true;
    });
}
