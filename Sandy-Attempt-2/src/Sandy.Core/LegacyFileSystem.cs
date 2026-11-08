namespace Sandy.Core;
public static class LegacyFileSystem
{
    public static IReadOnlyList<string> Enumerate(string pattern,bool directories)
    {
        var path=Path.GetDirectoryName(pattern);if(string.IsNullOrEmpty(path))path=Environment.CurrentDirectory;var mask=Path.GetFileName(pattern);if(mask.Length==0)mask="*";
        return (directories?Directory.EnumerateDirectories(path,mask):Directory.EnumerateFiles(path,mask)).Take(10001).ToArray() is {} results&&results.Length<=10000?results:throw new InvalidOperationException("Directory results exceed 10,000 entries.");
    }
}
