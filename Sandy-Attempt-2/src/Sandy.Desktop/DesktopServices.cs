using System.Diagnostics;
using Microsoft.Win32;
using Sandy.Blazor;
using Sandy.Core;
namespace Sandy.Desktop;
internal sealed class DesktopServices : IDesktopServices
{
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
    public IReadOnlyDictionary<string, string[]> ReadSchema(string path) => AccessLibrary.Schema(path);
    public void OpenUrl(string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || uri.Scheme is not ("http" or "https")) throw new ArgumentException("Enter an HTTP or HTTPS address.");
        Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true });
    }
}
