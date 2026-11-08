using System.Text.RegularExpressions;
namespace Sandy.Core;

/// <summary>Source-derived DBClassGen template dispatch, sharing SoftVariables and one virtual editor.</summary>
public static class LegacyDatabaseGenerator
{
    public static async Task<IReadOnlyDictionary<string,string>> GenerateAsync(DatabaseDesign design,Library library,string category,IEnumerable<DatabaseTable> selected,
        bool wrapper=false,string databaseName="Database",Func<Prompt,CancellationToken,Task<string?>>? prompt=null,CancellationToken cancellation=default)
    {
        design.Validate();var tables=selected.ToArray();var workingLibrary=library.Clone();var root=Path.Combine(Path.GetTempPath(),"Sandy-virtual-generation");
        var editor=new LegacyEditor(Path.Combine(root,"Generated.bas"),"");var variables=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase){["Database Name"]=databaseName,["Spaced Database Name"]=Spaced(databaseName),["DSN"]=databaseName};
        if(wrapper){await Run("Settings",true);await Run("Routines",true);await Run("Wrapper class");foreach(var table in tables.Where(t=>!design.Relations.Any(r=>r.ChildTable==t.Name))){TableVariables(table);variables["Property Name"]=Name(table.Name)+"s";await Run("Wrapper class - Add collection");}}
        foreach(var table in tables)
        {
            cancellation.ThrowIfCancellationRequested();TableVariables(table);
            var parent=design.Relations.FirstOrDefault(r=>r.ChildTable==table.Name);var children=design.Relations.Where(r=>r.ParentTable==table.Name).Select(r=>r.ChildTable).Distinct().ToArray();
            var suffix=(parent==null?", No Parent":"")+(children.Length==0?", No Child":"");
            var key=table.Fields.FirstOrDefault(f=>f.PrimaryKey)??table.Fields[0];var display=table.Fields.FirstOrDefault(f=>!f.PrimaryKey&&f.Name!=parent?.ChildField)??key;
            variables["AutoNumber Field Name"]=key.Name;variables["AutoNumber Property Name"]=Name(key.Name);variables["Key Field Name"]=display.Name;variables["Key Property Name"]=Name(display.Name);
            variables["Parent AutoNumber Field Name"]=parent?.ChildField??"";variables["Parent AutoNumber Property Name"]=Name(parent?.ChildField??"");
            variables["Class to collect"]=children.FirstOrDefault() is {} firstChild?Name(firstChild)+"s":"";
            variables["Child Table Name"]=children.FirstOrDefault()??"";variables["Singular Property Name"]=children.FirstOrDefault()??"";
            await Run("Table - "+table.Name,true);await Run("Collection"+suffix);await Run(children.Length==0?"Collection Member, Terminal":"Collection Member");
            foreach(var field in table.Fields.Where(f=>!f.Name.StartsWith("s_",StringComparison.OrdinalIgnoreCase)))
            {
                variables["Field Name"]=field.Name;variables["Spaced Field Name"]=Spaced(field.Name);variables["Property Name"]=Name(field.Name);
                variables["Property Type"]=PropertyType(field.DaoType);variables["Field Type"]=PropertyType(field.DaoType);variables["Property Size"]=variables["Property Length"]=variables["Field Length"]=field.Length.ToString();
                await Run("Property - "+PropertyType(field.DaoType));
                if(field.Name.EndsWith("ID",StringComparison.OrdinalIgnoreCase)&&field.Name!=key.Name)await Run("Property - 3D Link",true);
            }
            if(workingLibrary.Templates.Any(t=>workingLibrary.CategoryOf(t).Equals(category,StringComparison.OrdinalIgnoreCase)&&t["ShortTemplateName"].Equals("Table - "+table.Name+" - Finalize",StringComparison.OrdinalIgnoreCase)))await Run("Table - "+table.Name+" - Finalize");else await Run("Table - "+table.Name+", Finalize",true);
            foreach(var child in children.Skip(1)){variables["Singular Property Name"]=Name(child);variables["Property Name"]=Name(child)+"s";variables["Pure Child Table Name"]=child;variables["Child Table Name"]=Name(child);await Run("Collection Member - New Subcollection");}
        }
        if(wrapper)await Run("Finalize",true);
        var files=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        string Relative(string path){var relative=Path.GetRelativePath(root,Path.GetFullPath(path));if(Path.IsPathRooted(relative)||relative==".."||relative.StartsWith(".."+Path.DirectorySeparatorChar))throw new InvalidOperationException("A generator template targeted a path outside its output folder: "+path);return relative;}
        foreach(var doc in editor.Documents.Values.Where(d=>d.Changed))files.Add(Relative(doc.Path),doc.Text);
        foreach(var effect in editor.Effects)
        {
            if(effect.Kind is "EditorClose" or "EditorCloseAll")continue;
            if(effect.Kind is not ("FileAppend" or "FileOverwrite"))throw new NotSupportedException("Database generation produced a non-file action: "+effect.Kind);
            var name=Relative(effect.Target);files[name]=effect.Kind=="FileAppend"?files.GetValueOrDefault(name,"")+effect.Value:effect.Value;
        }
        return files;
        async Task Run(string shortName,bool optional=false)
        {
            var templates=workingLibrary.Categories.Where(c=>c.Name.Equals(category,StringComparison.OrdinalIgnoreCase)).SelectMany(c=>c.Templates).Where(t=>t["ShortTemplateName"].Equals(shortName,StringComparison.OrdinalIgnoreCase)).ToArray();
            if(templates.Length==0&&optional)return;if(templates.Length!=1)throw new InvalidOperationException("Missing or ambiguous generator template: "+category+" - "+shortName);
            var result=await new Engine().EvaluateAsync(workingLibrary,templates[0],variables,prompt,cancellation,editor);
            variables=new(result.Variables,StringComparer.OrdinalIgnoreCase);if(result.LibraryMutation is {} mutation)workingLibrary=mutation.Commit(workingLibrary);
        }
        void TableVariables(DatabaseTable t){variables["Pure Table Name"]=t.Name;variables["Table Name"]=variables["Object Name"]=Name(t.Name);variables["Spaced Table Name"]=variables["Spaced Object Name"]=variables["Label Name"]=Spaced(t.Name);}
    }
    private static string Name(string name)=>name.Replace(" ","_").Replace("*","_").Replace("-","_").Replace(".","__");
    private static string Spaced(string name)=>Regex.Replace(Name(name).Replace('_',' '),"([a-z0-9])([A-Z])","$1 $2");
    private static string PropertyType(int type)=>type switch{1=>"Boolean",2=>"Byte",3=>"Integer",4 or 16=>"Long",5=>"Currency",6=>"Single",7=>"Double",8=>"Date",10 or 15=>"String",12=>"Memo",_=>"Variant"};
}
