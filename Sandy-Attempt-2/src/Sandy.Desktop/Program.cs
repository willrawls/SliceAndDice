using Sandy.Desktop.Forms;
using System.Text.Json;

namespace Sandy.Desktop;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        ApplicationConfiguration.Initialize();
        if (args.Length == 3 && args[0] == "--export-mdb")
        {
            try { AccessLibrary.Read(args[1]).Save(args[2]); return 0; }
            catch (Exception ex) { File.WriteAllText(args[2] + ".error.txt", ex.ToString(), System.Text.Encoding.UTF8); return 1; }
        }
        if (args.Contains("--smoke-forms"))
        {
            var path = Path.GetFullPath(args.ElementAtOrDefault(Array.IndexOf(args, "--smoke-forms") + 1) ?? "form-previews");
            Directory.CreateDirectory(path);
            var results = new List<object>();
            foreach (var definition in LegacyForm.Catalog)
            {
                using var form = new LegacyForm(definition.ClassName);
                form.StartPosition = FormStartPosition.Manual; form.Location = new(-30000, -30000);
                form.Show(); Application.DoEvents();
                using var image = new Bitmap(form.Width, form.Height); form.DrawToBitmap(image, new Rectangle(Point.Empty, form.Size));
                image.Save(Path.Combine(path, definition.ClassName + ".png"));
                form.SmokeMainWorkflow(path);
                results.Add(new { definition.ClassName, definition.Source, form.Width, form.Height, form.DeviceDpi, Scale = form.AutoScaleDimensions.ToString(), CurrentScale = form.CurrentAutoScaleDimensions.ToString(), Font = form.Font.ToString() });
                form.Close();
            }
            File.WriteAllText(Path.Combine(path, "results.json"), JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }), System.Text.Encoding.UTF8);
            return 0;
        }
        var requested = args.SkipWhile(a => a != "--form").Skip(1).FirstOrDefault();
        Application.Run(requested == null ? new frmMain() : new LegacyForm(requested));
        return 0;
    }
}
