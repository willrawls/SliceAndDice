using System.Diagnostics;
using System.Text.RegularExpressions;
using Sandy.Core;

namespace Sandy.Desktop.Forms;

public partial class LegacyForm
{
    private void WireDialogs()
    {
        if (SourceClass is "frmFieldType" or "frmODBCFieldType")
        {
            var list = Get<ListBox>("lstType");
            list.SelectedIndexChanged += (_, _) => Get<TextBox>("txtLength").Enabled = list.Text == "Text";
            Get<TextBox>("txtLength").Enabled = list.Text == "Text";
        }
        if (SourceClass == "frmBrowser")
        {
            var address = Get<ComboBox>("cboAddress");
            address.Dock = DockStyle.Top;
            address.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Safe(() => OpenUrl(address.Text)); } };
        }
        if (SourceClass == "frmCommandHelp")
        {
            var file = Path.Combine(AppContext.BaseDirectory, "Libraries", "SoftCommands.xml");
            Get<TextBox>("txtSoftCommandName").Text = "SoftVariables and commands";
            Get<TextBox>("txtSyntax").Text = "%%Name::Default%%   ~~Set Name=Value";
            Get<TextBox>("txtLongDescription").Text = "Inline commands: " + string.Join(", ", Engine.InlineCommands) + "\r\nLine commands: SET, SETVALUE, STRCOPY, APPEND, INC, INCREMENT, DEC, DECREMENT, TRIM, LTRIM, RTRIM, SAVEVARS, RESTOREVARS, IGNOREBLANKS, INCLUDE.\r\nUnported commands stop evaluation with their template, area and line. See the port status document for limits.";
        }
        if (SourceClass is "frmSplash" or "frmProgress")
        {
            KeyPreview = true; KeyDown += (_, e) => { if (e.KeyCode == Keys.Escape) Close(); };
        }
    }

    private void HandleAction(string key) => Safe(() =>
    {
        Action = key;
        if (SourceClass == "frmMain")
        {
            switch (key)
            {
                case "mnuSpecialOpenDatabase": OpenLibrary(); return;
                case "mnuSpecialNewDatabase": NewLibrary(); return;
                case "mnuFileCopy": SaveLibrary(); return;
                case "mnuFileNew": AddTemplate(); return;
                case "mnuFileDelete": DeleteTemplate(); return;
                case "mnuFileExit": Close(); return;
                case "mnuEditFind": case "mnuEditReplace": FindReplace(); return;
                case "mnuEditCopy": ActiveEditor?.Copy(); return;
                case "mnuEditCut": if (ActiveEditor?.ReadOnly == false) ActiveEditor.Cut(); return;
                case "mnuEditPaste": if (ActiveEditor?.ReadOnly == false) ActiveEditor.Paste(); return;
                case "mnuDBClassGen": using (var form = new frmDBClassGen()) form.ShowDialog(this); return;
                case "mnuHelpAbout": using (var form = new frmSplash()) form.ShowDialog(this); return;
                case "mnuHelpSoftCommandReference": using (var form = new frmCommandHelp()) form.ShowDialog(this); return;
                case "cmdRecalc":
                    Get<ListBox>("lstSoftVariables").Items.Clear(); Get<ListBox>("lstSoftCommands").Items.Clear();
                    foreach (var text in codeEditors.Values.Select(x => x.Text))
                    {
                        foreach (Match match in Regex.Matches(text, @"%%(.*?)%%")) if (!Get<ListBox>("lstSoftVariables").Items.Contains(match.Groups[1].Value)) Get<ListBox>("lstSoftVariables").Items.Add(match.Groups[1].Value);
                        foreach (Match match in Regex.Matches(text, @"(?m)^\s*~~(\w+)")) if (!Get<ListBox>("lstSoftCommands").Items.Contains(match.Groups[1].Value)) Get<ListBox>("lstSoftCommands").Items.Add(match.Groups[1].Value);
                    }
                    return;
            }
        }
        if (key is "cmdCancel" or "cmdDone" or "mnuExit" or "mnuFileExit") { DialogResult = DialogResult.Cancel; Close(); return; }
        if (SourceClass == "frmFindReplace" && key is "cmdOK" or "cmdReplace" or "cmdReplaceAll") { DialogResult = DialogResult.OK; Close(); return; }
        if (SourceClass is "frmFieldType" or "frmODBCFieldType" && key == "cmdOkay")
        {
            if (string.IsNullOrWhiteSpace(Get<TextBox>("FieldName").Text)) throw new InvalidOperationException("Enter a field name.");
            if (Get<ListBox>("lstType").SelectedIndex < 0) throw new InvalidOperationException("Choose a field type.");
            if (Get<ListBox>("lstType").Text == "Text" && (!int.TryParse(Get<TextBox>("txtLength").Text, out var n) || n is < 1 or > 255)) throw new InvalidOperationException("Text length must be between 1 and 255.");
            Choice = Get<TextBox>("FieldName").Text.Trim(); DialogResult = DialogResult.OK; Close(); return;
        }
        if (SourceClass == "frmListSelect" && key == "cmdOkay")
        {
            var list = Get<ListView>("lstChoose"); if (list.SelectedItems.Count == 0) throw new InvalidOperationException("Choose an item first.");
            Choice = list.SelectedItems[0].Text; DialogResult = DialogResult.OK; Close(); return;
        }
        if (SourceClass == "frmFileList" && key == "cmdSearch")
        {
            var directory = Path.GetFullPath(Get<TextBox>("txtStartDir").Text);
            var pattern = Get<TextBox>("txtFilePattern").Text;
            var files = Directory.EnumerateFiles(directory, pattern.Length == 0 ? "*.*" : pattern, new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true, AttributesToSkip = FileAttributes.ReparsePoint }).Take(10000).ToArray();
            Choice = string.Join(Environment.NewLine, files);
            using var results = new frmListSelect();
            foreach (var file in files) results.Get<ListView>("lstChoose").Items.Add(new ListViewItem([Path.GetFileName(file), file]));
            results.ShowDialog(this);
            return;
        }
        if (SourceClass == "frmAutoWebURLEntry")
        {
            if (key == "cmdAdd") { var item = InputDialog.Ask(this, "POST field (Name=Value)", ""); if (!string.IsNullOrEmpty(item)) Get<ListBox>("lstDataToPost").Items.Add(item); return; }
            if (key == "cmdRemove") { var list = Get<ListBox>("lstDataToPost"); if (list.SelectedIndex >= 0) list.Items.RemoveAt(list.SelectedIndex); return; }
            if (key == "cmdOkay") { if (!Uri.TryCreate(Get<TextBox>("txtURL").Text, UriKind.Absolute, out var uri) || uri.Scheme is not ("https" or "http")) throw new InvalidOperationException("Enter an HTTP or HTTPS address."); Choice = uri.AbsoluteUri; DialogResult = DialogResult.OK; Close(); return; }
        }
        if (SourceClass == "frmLogin" && key == "cmdOK") { DialogResult = DialogResult.OK; Close(); return; }
        if (SourceClass == "ShellMain" && key == "mnuMainWindow") { using var form = new frmMain(); form.ShowDialog(this); return; }
        if (SourceClass == "frmOptions" && key == "cmdPickFile")
        {
            using var picker = new OpenFileDialog { Filter = "C/C++ headers|*.h;*.hpp|All files|*.*" };
            if (picker.ShowDialog(this) != DialogResult.OK) return;
            var text = System.IO.File.ReadAllText(picker.FileName);
            var tree = Get<TreeView>("tvwHierarchy"); tree.Nodes.Clear();
            foreach (Match match in Regex.Matches(text, @"(?m)^\s*(?:typedef\s+)?(?:struct|class|enum)\s+(\w+)")) tree.Nodes.Add(match.Groups[1].Value);
            return;
        }
        if (SourceClass is "frmDBClassGen" or "frmODBCClassGen" && key == "cmdAddCategory")
        {
            using var picker = new OpenFileDialog { Filter = "Access database|*.mdb;*.accdb" };
            if (picker.ShowDialog(this) != DialogResult.OK) return;
            var tree = Get<TreeView>("tvwTables"); tree.Nodes.Clear();
            foreach (var table in AccessLibrary.Schema(picker.FileName))
            {
                var root = tree.Nodes.Add(table.Key); foreach (var field in table.Value) root.Nodes.Add(field);
            }
            return;
        }
        if (SourceClass is "frmDBClassGen" or "frmODBCClassGen" && key == "cmdDeleteCategory") { Get<TreeView>("tvwTables").SelectedNode?.Remove(); return; }
        MessageBox.Show(this, $"The {SourceClass}.{key} workflow has not been ported yet. The original implementation is recorded in the source inventory.", "Sandy — port status");
    });

    private static void OpenUrl(string address)
    {
        if (!Uri.TryCreate(address, UriKind.Absolute, out var uri) || uri.Scheme is not ("http" or "https")) throw new InvalidOperationException("Enter an HTTP or HTTPS URL.");
        Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true });
    }
}
