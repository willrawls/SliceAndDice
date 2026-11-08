using System.Text.Json;

namespace Sandy.Desktop.Forms;

public sealed class FormDefinition
{
    public string Type { get; set; } = "";
    public string Name { get; set; } = "";
    public string ClassName { get; set; } = "";
    public string Source { get; set; } = "";
    public Dictionary<string, JsonElement> Properties { get; set; } = [];
    public List<FormDefinition> Children { get; set; } = [];
    public string Get(string key, string fallback = "") => Properties.TryGetValue(key, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString()! : fallback;
    public int Number(string key, int fallback = 0) => int.TryParse(Get(key), out var value) ? value : fallback;
    public bool Flag(string key, bool fallback = false) => Number(key, fallback ? -1 : 0) != 0;
    public string Key => Name + (Properties.ContainsKey("Index") ? "_" + Get("Index") : "");
}

public partial class LegacyForm : Form
{
    public static IReadOnlyList<FormDefinition> Catalog { get; } = LoadCatalog();
    private readonly Dictionary<string, Control> controlsByName = [];
    private readonly Dictionary<string, ToolStripItem> menus = [];
    private readonly ToolTip toolTips = new();
    public string SourceClass { get; }
    public string? Choice { get; private set; }
    public string Action { get; private set; } = "";
    public Control? FindControl(string name) => controlsByName.GetValueOrDefault(name);
    public T Get<T>(string name) where T : Control => (T)(FindControl(name) ?? throw new InvalidOperationException("Missing control: " + name));
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    public TextBox? ActiveEditor { get; set; }

    static IReadOnlyList<FormDefinition> LoadCatalog()
    {
        using var stream = typeof(LegacyForm).Assembly.GetManifestResourceStream("Sandy.Desktop.Forms.catalog.json")!;
        return JsonSerializer.Deserialize<List<FormDefinition>>(stream, new JsonSerializerOptions { PropertyNameCaseInsensitive = true })!;
    }

    public LegacyForm(string sourceClass)
    {
        SourceClass = sourceClass;
        var definition = Catalog.Single(x => x.ClassName == sourceClass);
        SuspendLayout();
        Name = sourceClass;
        Text = definition.Get("Caption", sourceClass);
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("Microsoft Sans Serif", 8.25f);
        AutoScroll = true;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(Math.Max(200, definition.Number("ClientWidth", 6000) / 15), Math.Max(100, definition.Number("ClientHeight", 4000) / 15));
        if (definition.Get("IconBase64").Length > 0)
        {
            using var stream = new MemoryStream(Convert.FromBase64String(definition.Get("IconBase64")));
            Icon = new Icon(stream);
        }
        if (definition.Number("BorderStyle") is 1 or 3 or 4) FormBorderStyle = FormBorderStyle.FixedDialog;
        var strip = new MenuStrip();
        foreach (var node in definition.Children)
            if (node.Type == "VB.Menu") strip.Items.Add(BuildMenu(node)); else AddControl(this, node);
        if (strip.Items.Count != 0) { Controls.Add(strip); MainMenuStrip = strip; }
        else strip.Dispose();
        WireDialogs();
        if (sourceClass == "frmMain") ConfigureMain();
        AutoScaleDimensions = new SizeF(96, 96);
        ResumeLayout(true);
    }

    private ToolStripItem BuildMenu(FormDefinition node)
    {
        ToolStripItem item;
        if (node.Get("Caption") == "-") item = new ToolStripSeparator();
        else
        {
            var menu = new ToolStripMenuItem(node.Get("Caption", node.Name)) { Checked = node.Flag("Checked") };
            foreach (var child in node.Children) menu.DropDownItems.Add(BuildMenu(child));
            if (menu.DropDownItems.Count == 0) menu.Click += (_, _) => HandleAction(node.Key);
            item = menu;
        }
        item.Name = node.Key; item.Enabled = node.Flag("Enabled", true); item.Visible = node.Flag("Visible", true);
        menus.Add(node.Key, item);
        return item;
    }

    private void AddControl(Control parent, FormDefinition node)
    {
        // Nonvisual COM components have no visible bounds in VB6 either.
        if (node.Type.EndsWith(".Timer") || node.Type.EndsWith(".ImageList") || node.Type.EndsWith(".CommonDialog") || node.Type.EndsWith(".Inet") || node.Type.EndsWith(".cSysTray")) return;
        var kind = node.Type.Split('.')[^1];
        Control control = kind switch
        {
            "CommandButton" => new Button(), "TextBox" => new TextBox(), "Label" => new Label(),
            "CheckBox" => new CheckBox(), "OptionButton" => new RadioButton(), "ComboBox" => new ComboBox(),
            "ListBox" => new ListBox(), "Frame" => new GroupBox(), "PictureBox" => new Panel(),
            "ListView" => new ListView { View = View.Details, FullRowSelect = true, HideSelection = false },
            "TreeView" or "FSListBar" => new TreeView { HideSelection = false },
            "ProgressBar" => new ProgressBar(), "TabStrip" => new TabControl(),
            "DataView" => new DataGridView { ReadOnly = true, AllowUserToAddRows = false },
            "Toolbar" => new ToolStrip(), "Image" => new PictureBox { SizeMode = PictureBoxSizeMode.Zoom },
            "Shape" => new Panel { BorderStyle = BorderStyle.FixedSingle },
            // Old IE browser is replaced with an explicit system-browser handoff.
            "WebBrowser" => new Label { Text = "Enter an address above to open it in your browser.", BorderStyle = BorderStyle.FixedSingle },
            _ => throw new NotSupportedException("Unmapped control: " + node.Type)
        };
        control.Name = node.Key;
        if (kind != "WebBrowser") control.Text = node.Get("Caption", node.Get("Text"));
        control.SetBounds(node.Number("Left") / 15, node.Number("Top") / 15, Math.Max(1, node.Number("Width", 1500) / 15), Math.Max(1, node.Number("Height", 300) / 15));
        control.TabIndex = node.Number("TabIndex");
        control.Visible = node.Flag("Visible", true); control.Enabled = node.Flag("Enabled", true);
        control.Tag = node;
        if (node.Get("PictureBase64").Length > 0)
        {
            using var stream = new MemoryStream(Convert.FromBase64String(node.Get("PictureBase64")));
            using var image = Image.FromStream(stream);
            if (control is PictureBox picture) picture.Image = new Bitmap(image);
            else control.BackgroundImage = new Bitmap(image);
        }
        if (node.Get("ToolTipText").Length != 0) toolTips.SetToolTip(control, node.Get("ToolTipText"));
        if (control is TextBox text)
        {
            text.Multiline = node.Flag("MultiLine"); text.AcceptsReturn = text.Multiline; text.AcceptsTab = text.Multiline;
            text.ReadOnly = node.Flag("Locked"); text.MaxLength = node.Number("MaxLength", 0) == 0 ? int.MaxValue : node.Number("MaxLength");
            text.UseSystemPasswordChar = node.Get("PasswordChar").Length > 0;
            if (text.Multiline) { text.ScrollBars = ScrollBars.Both; text.WordWrap = false; }
            if (node.Get("Font.Name") is "Fixedsys" or "Courier New") text.Font = new Font("Consolas", 10);
        }
        if (control is CheckBox check) check.Checked = node.Flag("Value");
        if (control is Label label && node.Flag("AutoSize")) label.AutoSize = true;
        if (control is RadioButton radio) radio.Checked = node.Flag("Value");
        if (control is ListView view) { view.Columns.Add("Name", Math.Max(100, control.Width / 2)); view.Columns.Add("Value", Math.Max(100, control.Width / 2)); }
        if (node.Properties.TryGetValue("Items", out var items))
        {
            var values = items.EnumerateArray().Select(x => (object)(x.GetString() ?? "")).ToArray();
            if (control is ListBox list) { list.Items.AddRange(values); if (values.Length > 0) list.SelectedIndex = 0; }
            if (control is ComboBox combo) { combo.Items.AddRange(values); if (values.Length > 0) combo.SelectedIndex = 0; }
        }
        if (control is ComboBox dropdown && node.Number("Style") == 2) dropdown.DropDownStyle = ComboBoxStyle.DropDownList;
        if (control is Button button)
        {
            button.Click += (_, _) => HandleAction(node.Key);
            if (node.Flag("Default")) AcceptButton = button;
            if (node.Flag("Cancel")) CancelButton = button;
        }
        controlsByName.Add(node.Key, control);
        parent.Controls.Add(control);
        foreach (var child in node.Children) AddControl(control, child);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) { toolTips.Dispose(); evaluationGate.Dispose(); }
        base.Dispose(disposing);
    }
}
