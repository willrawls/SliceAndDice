using Sandy.Core;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Sandy.Desktop.Forms;

public partial class LegacyForm
{
    private Library library = new();
    private Template? selectedTemplate;
    private readonly Dictionary<Template, string[]> drafts = [];
    private readonly Dictionary<CodeArea, TextBox> codeEditors = [];
    private TextBox? search, preview, variablesEditor;
    private TreeView? libraryTree;
    private bool dirty, refreshing;
    private string? sourcePath;
    private readonly SemaphoreSlim evaluationGate = new(1, 1);

    private void ConfigureMain()
    {
        Text = "Slice and Dice — Sandy · Attempt 2";
        ClientSize = new(1180, 760); MinimumSize = new(840, 520);
        var toolbar = new ToolStrip { Dock = DockStyle.Top };
        void Button(string title, System.Action action) => toolbar.Items.Add(title, null, (_, _) => Safe(action));
        Button("Open library", OpenLibrary); Button("New library", NewLibrary); Button("Save copy", SaveLibrary);
        Button("New category", AddCategory); Button("New template", AddTemplate); Button("Update template", CommitDraft);
        toolbar.Items.Add("Preview", null, async (_, _) => await PreviewTemplate());
        Button("Find / replace", FindReplace);
        var formsMenu = new ToolStripDropDownButton("All forms");
        foreach (var definition in Catalog.Where(x => x.ClassName != "frmMain"))
        {
            var formName = definition.ClassName;
            formsMenu.DropDownItems.Add(formName + " — " + definition.Get("Caption"), null, (_, _) => Safe(() => { using var form = new LegacyForm(formName); form.ShowDialog(this); }));
        }
        toolbar.Items.Add(formsMenu);
        var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 280, FixedPanel = FixedPanel.Panel1 };
        search = new TextBox { Dock = DockStyle.Top, PlaceholderText = "Filter templates" };
        libraryTree = Get<TreeView>("lsbJumpTo"); libraryTree.Dock = DockStyle.Fill; libraryTree.Visible = true;
        split.Panel1.Controls.Add(libraryTree); split.Panel1.Controls.Add(search);
        var editorLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1 };
        editorLayout.RowStyles.Add(new(SizeType.Absolute, 30)); editorLayout.RowStyles.Add(new(SizeType.Percent, 100)); editorLayout.RowStyles.Add(new(SizeType.Absolute, 28));
        var titleRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 };
        titleRow.ColumnStyles.Add(new(SizeType.Percent, 60)); titleRow.ColumnStyles.Add(new(SizeType.Percent, 40));
        var name = Get<TextBox>("txtName"); var shortName = Get<TextBox>("txtShortName");
        name.Visible = shortName.Visible = true; name.Dock = shortName.Dock = DockStyle.Fill;
        titleRow.Controls.Add(name); titleRow.Controls.Add(shortName);
        var tabs = new TabControl { Dock = DockStyle.Fill };
        TextBox Editor(string key, CodeArea area, string title)
        {
            var editor = Get<TextBox>(key); editor.Visible = true; editor.Dock = DockStyle.Fill;
            editor.Multiline = true; editor.AcceptsTab = true; editor.Font = new Font("Consolas", 10); editor.WordWrap = false;
            editor.TextChanged += (_, _) => { if (!refreshing) dirty = true; };
            editor.Enter += (_, _) => ActiveEditor = editor;
            codeEditors.Add(area, editor);
            var page = new TabPage(title); page.Controls.Add(editor); tabs.TabPages.Add(page);
            return editor;
        }
        Editor("txtCode_0", CodeArea.Declarations, "Declarations");
        Editor("txtCode_1", CodeArea.Cursor, "At cursor");
        Editor("txtCode_2", CodeArea.Bottom, "At bottom");
        Editor("txtCodeToFile", CodeArea.File, "To file");
        var fileName = Get<TextBox>("txtFilename"); fileName.Dock = DockStyle.Top; fileName.Visible = true;
        tabs.TabPages[3].Controls.Add(fileName);
        variablesEditor = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 10), PlaceholderText = "SoftVariables: one Name=Value per line" };
        var variablesPage = new TabPage("SoftVariables"); variablesPage.Controls.Add(variablesEditor); tabs.TabPages.Add(variablesPage);
        preview = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 10), WordWrap = false };
        var previewPage = new TabPage("Preview"); previewPage.Controls.Add(preview); tabs.TabPages.Add(previewPage);
        // Preserve all original detail controls, grouped into their original containers.
        foreach (var key in new[] { "frmTemplateInfo", "frmOptions" })
        {
            var details = FindControl(key);
            if (details != null) { var page = new TabPage(key == "frmOptions" ? "Template options" : "Template information"); details.Dock = DockStyle.Fill; details.Visible = true; page.Controls.Add(details); tabs.TabPages.Add(page); }
        }
        var status = new Label { Dock = DockStyle.Fill, Text = "Open a Sandy MDB or Category.xml library. Save copy preserves the original.", AutoEllipsis = true };
        editorLayout.Controls.Add(titleRow, 0, 0); editorLayout.Controls.Add(tabs, 0, 1); editorLayout.Controls.Add(status, 0, 2);
        split.Panel2.Controls.Add(editorLayout);
        Shown += (_, _) => split.SplitterDistance = Math.Min(320, split.Width / 3);
        // Retain source menu and controls not promoted into the working layout.
        foreach (Control control in Controls.Cast<Control>().ToArray()) if (control != MainMenuStrip) control.Visible = false;
        Controls.Add(split); Controls.Add(toolbar); if (MainMenuStrip != null) MainMenuStrip.BringToFront(); toolbar.BringToFront();
        search.TextChanged += (_, _) => RefreshTree();
        libraryTree.AfterSelect += (_, args) => { if (args.Node?.Tag is Template t) SelectTemplate(t); };
        name.TextChanged += (_, _) => { if (!refreshing) dirty = true; }; shortName.TextChanged += (_, _) => { if (!refreshing) dirty = true; };
        fileName.TextChanged += (_, _) => { if (!refreshing) dirty = true; };
        foreach (var (key, field) in new[] { ("chkLocked", "Locked"), ("chkFavorite", "Favorite"), ("chkUndeletable", "Undeletable") })
            if (FindControl(key) is CheckBox flag) flag.CheckedChanged += (_, _) => { if (!refreshing && selectedTemplate != null) { selectedTemplate[field] = flag.Checked ? "-1" : "0"; dirty = true; } };
        FormClosing += (_, e) => { if (dirty && MessageBox.Show(this, "Discard unsaved library changes and drafts?", "Sandy", MessageBoxButtons.YesNo) != DialogResult.Yes) e.Cancel = true; };
        var arguments = Environment.GetCommandLineArgs();
        var file = arguments.Skip(1).FirstOrDefault(x => !x.StartsWith("--") && System.IO.File.Exists(x));
        file ??= System.IO.File.Exists(Path.Combine(AppContext.BaseDirectory, "Libraries", "SliceAndDice.xml")) ? Path.Combine(AppContext.BaseDirectory, "Libraries", "SliceAndDice.xml") : null;
        if (file != null) Shown += (_, _) => Safe(() => LoadLibrary(file));
    }

    private void Safe(System.Action action)
    {
        try { action(); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Sandy", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }
    private void StoreDraft()
    {
        if (selectedTemplate == null) return;
        drafts[selectedTemplate] = [Get<TextBox>("txtName").Text, Get<TextBox>("txtShortName").Text, Get<TextBox>("txtFilename").Text,
            .. Enum.GetValues<CodeArea>().Select(a => codeEditors[a].Text)];
    }
    private void SelectTemplate(Template template)
    {
        StoreDraft(); selectedTemplate = template; refreshing = true;
        try
        {
            var draft = drafts.GetValueOrDefault(template);
            Get<TextBox>("txtName").Text = draft?[0] ?? template.Name;
            Get<TextBox>("txtShortName").Text = draft?[1] ?? template["ShortTemplateName"];
            Get<TextBox>("txtFilename").Text = draft?[2] ?? template["Filename"];
            foreach (var area in Enum.GetValues<CodeArea>()) { codeEditors[area].Text = draft?[(int)area + 3] ?? template.Code(area); codeEditors[area].ReadOnly = template.Locked; }
            foreach (var (key, field) in new[] { ("chkLocked", "Locked"), ("chkFavorite", "Favorite"), ("chkUndeletable", "Undeletable") })
                if (FindControl(key) is CheckBox flag) flag.Checked = template.Flag(field);
        }
        finally { refreshing = false; }
    }
    private void RefreshTree()
    {
        if (libraryTree == null) return;
        libraryTree.BeginUpdate(); libraryTree.Nodes.Clear();
        foreach (var category in library.Categories)
        {
            var node = new TreeNode(category.Name) { Tag = category };
            foreach (var template in category.Templates.Where(t => t.Name.Contains(search?.Text ?? "", StringComparison.OrdinalIgnoreCase)))
                node.Nodes.Add(new TreeNode(template.Name) { Tag = template });
            if (node.Nodes.Count > 0 || string.IsNullOrEmpty(search?.Text)) libraryTree.Nodes.Add(node);
        }
        libraryTree.ExpandAll(); libraryTree.EndUpdate();
    }
    private bool MayDiscard() => !dirty || MessageBox.Show(this, "Discard unsaved library changes and drafts?", "Sandy", MessageBoxButtons.YesNo) == DialogResult.Yes;
    private void OpenLibrary()
    {
        using var dialog = new OpenFileDialog { Filter = "Sandy libraries|*.xml;*.mdb|All files|*.*" };
        if (dialog.ShowDialog(this) == DialogResult.OK && MayDiscard()) LoadLibrary(dialog.FileName);
    }
    public void LoadLibrary(string filename)
    {
        var loaded = Path.GetExtension(filename).Equals(".mdb", StringComparison.OrdinalIgnoreCase) ? AccessLibrary.Read(filename) : Library.Load(filename);
        library = loaded; sourcePath = Path.GetFullPath(filename); drafts.Clear(); selectedTemplate = null; dirty = false; RefreshTree();
        if (libraryTree?.Nodes.Count > 0) libraryTree.SelectedNode = libraryTree.Nodes[0].FirstNode;
        Text = "Slice and Dice — " + Path.GetFileName(filename) + " · Attempt 2";
    }
    private void NewLibrary() { if (!MayDiscard()) return; library = new(); selectedTemplate = null; drafts.Clear(); sourcePath = null; dirty = true; AddCategory(); }
    private void SaveLibrary()
    {
        StoreDraft();
        foreach (var template in drafts.Keys.ToArray())
        {
            var draft = drafts[template];
            var code = Enum.GetValues<CodeArea>().ToDictionary(a => a, a => draft[(int)a + 3]);
            var changed = code.Any(x => template.Code(x.Key) != x.Value) || template["Filename"] != draft[2] || template.Name != draft[0] || template["ShortTemplateName"] != draft[1];
            if (changed && template.Locked) throw new InvalidOperationException("Locked template has unsaved edits: " + template.Name);
        }
        using var dialog = new SaveFileDialog { Filter = "Sandy XML library|*.xml", FileName = "Sandy-working-copy.xml", OverwritePrompt = true };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        if (string.Equals(Path.GetFullPath(dialog.FileName), sourcePath, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Choose a new filename to preserve the imported source.");
        foreach (var template in drafts.Keys.ToArray()) ApplyDraft(template);
        library.Save(dialog.FileName); dirty = false;
    }
    private void ApplyDraft(Template template)
    {
        if (!drafts.TryGetValue(template, out var draft)) return;
        library.Update(template, Enum.GetValues<CodeArea>().ToDictionary(a => a, a => draft[(int)a + 3]), draft[2]);
        if (!template.Locked) { template["TemplateName"] = draft[0]; template["ShortTemplateName"] = draft[1]; }
        drafts.Remove(template);
    }
    private void CommitDraft() { StoreDraft(); if (selectedTemplate != null) ApplyDraft(selectedTemplate); RefreshTree(); }
    private void AddCategory()
    {
        var name = InputDialog.Ask(this, "Category name", "New category"); if (string.IsNullOrWhiteSpace(name)) return;
        var category = new Category(); category.Fields["CategoryID"] = NextId(library.Categories.Select(c => c.Fields["CategoryID"])); category.Fields["CategoryName"] = name;
        library.Categories.Add(category); dirty = true; RefreshTree();
    }
    private static string NextId(IEnumerable<string> values) => checked(values.Select(s => int.TryParse(s, out var n) ? n : 0).DefaultIfEmpty(0).Max() + 1).ToString();
    private void AddTemplate()
    {
        var category = libraryTree?.SelectedNode?.Tag as Category ?? library.Categories.FirstOrDefault(c => selectedTemplate != null && c.Templates.Contains(selectedTemplate)) ?? library.Categories.FirstOrDefault();
        if (category == null) { AddCategory(); category = library.Categories.LastOrDefault(); if (category == null) return; }
        var name = InputDialog.Ask(this, "Template name", "New template"); if (string.IsNullOrWhiteSpace(name)) return;
        var template = new Template(); template["TemplateID"] = NextId(library.Templates.Select(t => t["TemplateID"])); template["CategoryID"] = category.Fields["CategoryID"];
        template["TemplateName"] = category.Name + " - " + name; template["ShortTemplateName"] = name; template["RevisionCount"] = "0";
        category.Templates.Add(template); dirty = true; RefreshTree(); SelectTemplate(template);
    }
    private void DeleteTemplate()
    {
        if (selectedTemplate == null) return;
        if (selectedTemplate.Locked || selectedTemplate.Flag("Undeletable")) throw new InvalidOperationException("The template is locked or marked undeletable.");
        if (MessageBox.Show(this, "Delete " + selectedTemplate.Name + " from this working library?", "Sandy", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
        foreach (var category in library.Categories) category.Templates.Remove(selectedTemplate);
        drafts.Remove(selectedTemplate); selectedTemplate = null; dirty = true; RefreshTree();
    }
    private async Task PreviewTemplate()
    {
        if (selectedTemplate == null || !await evaluationGate.WaitAsync(0)) return;
        try
        {
            StoreDraft(); var working = new Template(); foreach (var pair in selectedTemplate.Fields) working[pair.Key] = pair.Value;
            var draft = drafts[selectedTemplate]; foreach (var area in Enum.GetValues<CodeArea>()) working[Template.CodeField(area)] = draft[(int)area + 3]; working["Filename"] = draft[2];
            var variables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in variablesEditor!.Lines.Where(x => x.Trim().Length > 0)) { var split = line.IndexOf('='); if (split < 1) throw new FormatException("SoftVariables require Name=Value."); variables[line[..split].Trim()] = line[(split + 1)..]; }
            var result = await new Engine().EvaluateAsync(library, working, variables, (p, _) => Task.FromResult(InputDialog.Ask(this, p.Name, p.Default)));
            preview!.Text = string.Join("\r\n\r\n", result.Outputs.Select(o => $"[{o.Area}] {o.Target}\r\n{o.Text}"));
            ((TabControl)preview.Parent!.Parent!).SelectedTab = (TabPage)preview.Parent;
        }
        catch (OperationCanceledException) { }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Sandy"); }
        finally { evaluationGate.Release(); }
    }
    private void FindReplace()
    {
        var editor = ActiveEditor ?? codeEditors.GetValueOrDefault(CodeArea.Cursor); if (editor == null) return;
        using var dialog = new frmFindReplace();
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var find = dialog.Get<TextBox>("txtFind").Text; if (find.Length == 0) return;
        var replacement = dialog.Get<TextBox>("txtReplace").Text;
        var comparison = dialog.Get<CheckBox>("chkMatchCase").Checked ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var index = editor.Text.IndexOf(find, editor.SelectionStart + editor.SelectionLength, comparison);
        if (index < 0) index = editor.Text.IndexOf(find, comparison);
        if (dialog.Action == "cmdReplaceAll" && !editor.ReadOnly) editor.Text = editor.Text.Replace(find, replacement, comparison);
        else if (index >= 0) { editor.Select(index, find.Length); if (dialog.Action == "cmdReplace" && !editor.ReadOnly) editor.SelectedText = replacement; editor.Focus(); }
    }

    internal void SmokeMainWorkflow(string outputDirectory)
    {
        if (SourceClass != "frmMain") return;
        var first = library.Templates.First(t => !t.Locked);
        var second = library.Templates.First(t => t != first);
        SelectTemplate(first);
        var original = first.Code(CodeArea.Cursor);
        var edited = original + "\r\n// smoke draft";
        codeEditors[CodeArea.Cursor].Text = edited;
        SelectTemplate(second); SelectTemplate(first);
        if (codeEditors[CodeArea.Cursor].Text != edited) throw new InvalidOperationException("Draft did not survive template switching.");
        if (first.Code(CodeArea.Cursor) != original) throw new InvalidOperationException("Draft changed the model before Update.");
        CommitDraft();
        if (first.Code(CodeArea.Cursor) != edited) throw new InvalidOperationException("Update did not apply the draft.");
        var path = Path.Combine(outputDirectory, "workflow-library.xml");
        library.Save(path);
        if (Library.Load(path).Find(first.Name).Code(CodeArea.Cursor) != edited) throw new InvalidOperationException("Saved draft did not reload.");
        dirty = false;
    }
}

internal static class InputDialog
{
    public static string? Ask(IWin32Window owner, string question, string value)
    {
        using var form = new Form { Text = question, ClientSize = new(480, 125), FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, MinimizeBox = false, MaximizeBox = false };
        var label = new Label { Text = question, Left = 12, Top = 10, Width = 450 };
        var input = new TextBox { Text = value, Left = 12, Top = 35, Width = 450 };
        var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 290, Top = 80 };
        var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 380, Top = 80 };
        form.Controls.AddRange([label, input, ok, cancel]); form.AcceptButton = ok; form.CancelButton = cancel;
        return form.ShowDialog(owner) == DialogResult.OK ? input.Text : null;
    }
}
