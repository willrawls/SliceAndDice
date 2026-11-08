using System.Windows;
using System.IO;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.AspNetCore.Components.WebView.Wpf;
using Sandy.Blazor;
using Sandy.Blazor.Components;
namespace Sandy.Desktop;
internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length == 3 && args[0] == "--export-mdb")
        {
            try { AccessLibrary.Read(args[1]).Save(args[2]); return 0; }
            catch (Exception ex) { File.WriteAllText(args[2] + ".error.txt", ex.ToString()); return 1; }
        }
        var smoke = args.Length == 2 && args[0] == "--smoke-ui" ? Path.GetFullPath(args[1]) : null;
        var app = new Application();
        var registrations = new ServiceCollection();
        registrations.AddWpfBlazorWebView();
        using var services = registrations.AddSingleton<Workspace>().AddSingleton<IDesktopServices, DesktopServices>().BuildServiceProvider();
        var state = services.GetRequiredService<Workspace>();
        var filename = args.FirstOrDefault(x => File.Exists(x)) ?? Path.Combine(AppContext.BaseDirectory, "Libraries", "SliceAndDice.xml");
        try { state.Load(services.GetRequiredService<IDesktopServices>().ReadLibrary(filename), filename); }
        catch (Exception ex) { state.Status = ex.Message; }
        var view = new BlazorWebView { HostPage = "wwwroot/index.html", Services = services };
        view.RootComponents.Add(new RootComponent { Selector = "#app", ComponentType = typeof(Workbench) });
        var window = new Window { Title = "Sandy — Blazor workspace", Width = 1360, Height = 900, MinWidth = 850, MinHeight = 600, Content = view };
        window.Closing += (_, e) =>
        {
            if (smoke == null && state.IsDirty && MessageBox.Show(window, "Discard unsaved library changes?", "Sandy", MessageBoxButton.YesNo) != MessageBoxResult.Yes) e.Cancel = true;
        };
        view.BlazorWebViewInitializing += (_, e) => e.UserDataFolder = smoke == null
            ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SandyAttempt2", "WebView2")
            : Path.Combine(smoke, "webview-profile");
        if (smoke != null)
        {
            Directory.CreateDirectory(smoke);
            window.Left = -20000; window.Top = -20000; window.ShowInTaskbar = false;
            view.BlazorWebViewInitialized += async (_, e) =>
            {
                try { await HybridSmoke.RunAsync(e.WebView.CoreWebView2, state, smoke); app.Shutdown(0); }
                catch (Exception ex) { File.WriteAllText(Path.Combine(smoke, "error.txt"), ex.ToString()); app.Shutdown(1); }
            };
            _ = TimeoutAsync();
            async Task TimeoutAsync() { await Task.Delay(TimeSpan.FromSeconds(55)); if (app.MainWindow != null) { File.WriteAllText(Path.Combine(smoke, "timeout.txt"), "Hybrid smoke test timed out."); app.Shutdown(2); } }
        }
        var exit = app.Run(window);
        view.DisposeAsync().AsTask().GetAwaiter().GetResult();
        return exit;
    }
}
