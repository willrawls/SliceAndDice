using System.Xml.Linq;
namespace Sandy.Core;
public sealed partial class LegacyEditor
{
    public string? ProjectPath { get; private set; }
    public string WorkspaceRoot { get; private set; } = "";
    public void AttachProject(string path,string xml)
    {
        var project=XDocument.Parse(xml);if(project.Root?.Name.LocalName!="Project")throw new InvalidDataException("Not an MSBuild project.");
        ProjectPath=Path.GetFullPath(path);WorkspaceRoot=Path.GetDirectoryName(ProjectPath)!;ProjectName=Path.GetFileNameWithoutExtension(path);
        Documents.TryAdd(ProjectPath,new(ProjectPath,xml));
    }
    public static string? FindProject(string sourcePath)
    {
        for(var dir=Path.GetDirectoryName(Path.GetFullPath(sourcePath));dir!=null;dir=Path.GetDirectoryName(dir))
        {
            var projects=Directory.EnumerateFiles(dir,"*.*proj").Where(p=>Path.GetExtension(p) is ".csproj" or ".vbproj").Take(2).ToArray();
            if(projects.Length==1)return projects[0];
            if(projects.Length>1)return null;
            if(Directory.Exists(Path.Combine(dir,".git")))break;
        }
        return null;
    }
    private void AddProjectItem(string item,string kind)
    {
        if(ProjectPath==null)throw new InvalidOperationException("No project is attached. Open a source file inside a .csproj/.vbproj project.");
        var document=Documents[ProjectPath];var project=XDocument.Parse(document.Text,LoadOptions.PreserveWhitespace);var root=project.Root!;var ns=root.Name.Namespace;
        if(item.Contains("{",StringComparison.Ordinal)||item.StartsWith("*",StringComparison.Ordinal))throw new NotSupportedException("VB6 type-library references require an explicit .NET interop assembly path.");
        var include=item;XElement? hint=null;
        if(kind=="Reference" && item.EndsWith("proj",StringComparison.OrdinalIgnoreCase))kind="ProjectReference";
        else if(kind=="Reference" && item.EndsWith(".dll",StringComparison.OrdinalIgnoreCase)) { include=Path.GetFileNameWithoutExtension(item);hint=new(ns+"HintPath",Path.GetRelativePath(WorkspaceRoot,Path.GetFullPath(PathName(item)))); }
        if(kind is "Compile" or "ProjectReference")include=Path.GetRelativePath(WorkspaceRoot,Path.GetFullPath(PathName(item)));
        if(kind=="Compile" && root.Attribute("Sdk")!=null && !include.StartsWith("..",StringComparison.Ordinal) &&
            !root.Descendants(ns+"EnableDefaultCompileItems").Any(e=>e.Value.Equals("false",StringComparison.OrdinalIgnoreCase)))return;
        if(root.Descendants(ns+kind).Any(e=>((string?)e.Attribute("Include"))?.Equals(include,StringComparison.OrdinalIgnoreCase)==true))return;
        var group=root.Elements(ns+"ItemGroup").LastOrDefault();if(group==null){group=new(ns+"ItemGroup");root.Add(group);}
        var element=new XElement(ns+kind,new XAttribute("Include",include));if(hint!=null)element.Add(hint);group.Add(element);
        document.Text=project.ToString();
    }
}
