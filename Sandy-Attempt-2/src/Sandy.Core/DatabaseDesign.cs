using System.Text;
using System.Text.Json;
namespace Sandy.Core;

public enum SqlDialect { Access, SqlServer, Sqlite }
public sealed class DatabaseField
{
    public string Name { get; set; }="Field";
    public int DaoType { get; set; }=10;
    public int Length { get; set; }=255;
    public bool Nullable { get; set; }=true;
    public bool PrimaryKey { get; set; }
    public bool AutoNumber { get; set; }
    public SchemaField Schema=>new(Name,DaoType,Length,Nullable&&!PrimaryKey);
}
public sealed class DatabaseTable
{
    public string Name { get; set; }="Table";
    public bool Marked { get; set; }=true;
    public bool Favorite { get; set; }
    public List<DatabaseField> Fields { get; set; }=[];
}
public sealed record DatabaseRelation(string ParentTable,string ParentField,string ChildTable,string ChildField,bool CascadeDelete=false,bool CascadeUpdate=false);
public sealed class DatabaseDesign
{
    public List<DatabaseTable> Tables { get; set; }=[];
    public List<DatabaseRelation> Relations { get; set; }=[];
    public static DatabaseDesign FromSchema(IReadOnlyDictionary<string,string[]> schema)
    {
        var design=new DatabaseDesign();
        foreach(var table in schema)design.Tables.Add(new(){Name=table.Key,Fields=DatabaseGenerator.ParseDaoFields(table.Value).Select(f=>new DatabaseField{Name=f.Name,DaoType=f.DaoType,Length=f.Length,Nullable=f.Nullable}).ToList()});
        return design;
    }
    public DatabaseTable AddTable(string name)
    {
        if(string.IsNullOrWhiteSpace(name)||Tables.Any(t=>t.Name.Equals(name,StringComparison.OrdinalIgnoreCase)))throw new ArgumentException("Enter a unique table name.");
        var table=new DatabaseTable{Name=name.Trim()};Tables.Add(table);return table;
    }
    public void RemoveTable(DatabaseTable table){Tables.Remove(table);Relations.RemoveAll(r=>r.ParentTable==table.Name||r.ChildTable==table.Name);}
    public void RenameTable(DatabaseTable table,string name)
    {
        name=name.Trim();if(name.Length==0||Tables.Any(t=>t!=table&&t.Name.Equals(name,StringComparison.OrdinalIgnoreCase)))throw new ArgumentException("Enter a unique table name.");
        var old=table.Name;table.Name=name;Relations=Relations.Select(r=>r with{ParentTable=r.ParentTable==old?name:r.ParentTable,ChildTable=r.ChildTable==old?name:r.ChildTable}).ToList();
    }
    public void RenameField(DatabaseTable table,DatabaseField field,string name)
    {
        name=name.Trim();if(name.Length==0||table.Fields.Any(f=>f!=field&&f.Name.Equals(name,StringComparison.OrdinalIgnoreCase)))throw new ArgumentException("Enter a unique field name.");
        var old=field.Name;field.Name=name;Relations=Relations.Select(r=>r with{ParentField=r.ParentTable==table.Name&&r.ParentField==old?name:r.ParentField,ChildField=r.ChildTable==table.Name&&r.ChildField==old?name:r.ChildField}).ToList();
    }
    public void ApplyRules(bool addKey,bool autoNumber,bool dateCreated,bool dateModified)
    {
        foreach(var table in Tables.Where(t=>t.Marked))
        {
            if(addKey&&!table.Fields.Any(f=>f.PrimaryKey)) {
                var field=table.Fields.FirstOrDefault(f=>f.Name.Equals(table.Name+"ID",StringComparison.OrdinalIgnoreCase));
                if(field==null){field=new(){Name=table.Name+"ID",DaoType=4,Length=0};table.Fields.Insert(0,field);}field.PrimaryKey=true;field.Nullable=false;field.AutoNumber=autoNumber;
            }
            foreach(var name in new[]{dateCreated?"DateCreated":"",dateModified?"DateModified":""}.Where(n=>n.Length>0))
                if(!table.Fields.Any(f=>f.Name.Equals(name,StringComparison.OrdinalIgnoreCase)))table.Fields.Add(new(){Name=name,DaoType=8,Length=0,Nullable=false});
        }
    }
    public IReadOnlyList<DatabaseTable> Branch(string name)
    {
        var visited=new HashSet<string>(StringComparer.OrdinalIgnoreCase);var result=new List<DatabaseTable>();
        void Visit(string table){if(!visited.Add(table))return;result.Add(Tables.Single(t=>t.Name==table));foreach(var child in Relations.Where(r=>r.ParentTable==table))Visit(child.ChildTable);}
        Visit(name);return result;
    }
    public void Validate()
    {
        if(Tables.Select(t=>t.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count()!=Tables.Count)throw new InvalidDataException("Duplicate table names.");
        if(Tables.Select(t=>DatabaseGenerator.Identifier(t.Name)).Distinct(StringComparer.Ordinal).Count()!=Tables.Count)throw new InvalidDataException("Table names collide as C# identifiers.");
        foreach(var table in Tables)
        {
            if(string.IsNullOrWhiteSpace(table.Name)||table.Fields.Count==0)throw new InvalidDataException("Each table needs a name and at least one field.");
            var names=new HashSet<string>(StringComparer.OrdinalIgnoreCase);var identifiers=new HashSet<string>(StringComparer.Ordinal);
            foreach(var field in table.Fields)
            {
                if(string.IsNullOrWhiteSpace(field.Name)||!names.Add(field.Name)||!identifiers.Add(DatabaseGenerator.Identifier(field.Name)))throw new InvalidDataException("Duplicate/invalid field name in "+table.Name);
                if(DatabaseGenerator.Identifier(field.Name)==DatabaseGenerator.Identifier(table.Name))throw new InvalidDataException("A field cannot have the same C# name as its table class.");
                if(field.AutoNumber&&(!field.PrimaryKey||field.DaoType is not (4 or 16)))throw new InvalidDataException("AutoNumber requires an integer primary key.");
                if(field.PrimaryKey&&field.Nullable)throw new InvalidDataException("Primary keys cannot be nullable.");
                if(field.DaoType==10&&field.Length is <1 or >65535)throw new InvalidDataException("Text length must be 1 to 65535.");
            }
            if(table.Fields.Count(f=>f.PrimaryKey)>1)throw new NotSupportedException("Choose a single primary key for generated repositories.");
        }
        foreach(var relation in Relations)
        {
            var parent=Tables.SingleOrDefault(t=>t.Name==relation.ParentTable)?.Fields.SingleOrDefault(f=>f.Name==relation.ParentField);
            var child=Tables.SingleOrDefault(t=>t.Name==relation.ChildTable)?.Fields.SingleOrDefault(f=>f.Name==relation.ChildField);
            if(parent==null||child==null||!parent.PrimaryKey||parent.DaoType!=child.DaoType)throw new InvalidDataException("Relationships require an existing parent primary key and a child field of the same type.");
        }
    }
    public string ToJson(){Validate();return JsonSerializer.Serialize(this,new JsonSerializerOptions{WriteIndented=true});}
    public static DatabaseDesign FromJson(string text){var design=JsonSerializer.Deserialize<DatabaseDesign>(text)??throw new InvalidDataException("Missing database design.");design.Validate();return design;}
    public static string Quote(string name,SqlDialect dialect)=>dialect==SqlDialect.Sqlite?"\""+name.Replace("\"","\"\"")+"\"":"["+name.Replace("]","]]")+"]";
    public string GenerateSql(SqlDialect dialect,IEnumerable<DatabaseTable>? tables=null)
    {
        Validate();var selected=(tables??Tables.Where(t=>t.Marked)).ToArray();var builder=new StringBuilder();var constraints=new List<string>();
        foreach(var table in selected)
        {
            var lines=new List<string>();foreach(var field in table.Fields)
            {
                var type=field.DaoType switch{1=>dialect==SqlDialect.Access?"YESNO":"BIT",2=>"TINYINT",3=>"SMALLINT",4=>"INTEGER",5 or 20=>"DECIMAL(19,4)",6=>"REAL",7=>"DOUBLE",8=>"DATETIME",9 or 11=>dialect==SqlDialect.SqlServer?"VARBINARY(MAX)":"BLOB",10=>dialect==SqlDialect.Access?$"TEXT({Math.Min(255,field.Length)})":$"VARCHAR({field.Length})",12=>dialect==SqlDialect.SqlServer?"NVARCHAR(MAX)":"TEXT",15=>dialect==SqlDialect.SqlServer?"UNIQUEIDENTIFIER":"VARCHAR(36)",16=>"BIGINT",_=>throw new NotSupportedException("Unknown DAO type: "+field.DaoType)};
                if(dialect==SqlDialect.SqlServer&&type=="DOUBLE")type="FLOAT";
                if(field.AutoNumber)type=dialect switch{SqlDialect.Access=>"COUNTER",SqlDialect.SqlServer=>type+" IDENTITY(1,1)",_=>"INTEGER"};
                var suffix=field.AutoNumber&&dialect==SqlDialect.Sqlite?" PRIMARY KEY AUTOINCREMENT":field.PrimaryKey?" PRIMARY KEY":!field.Nullable?" NOT NULL":"";
                lines.Add("    "+Quote(field.Name,dialect)+" "+type+suffix);
            }
            foreach(var relation in Relations.Where(r=>r.ChildTable==table.Name))
            {
                var constraint="FOREIGN KEY ("+Quote(relation.ChildField,dialect)+") REFERENCES "+Quote(relation.ParentTable,dialect)+" ("+Quote(relation.ParentField,dialect)+")"+(relation.CascadeDelete?" ON DELETE CASCADE":"")+(relation.CascadeUpdate?" ON UPDATE CASCADE":"");
                if(dialect==SqlDialect.Sqlite)lines.Add("    "+constraint);else constraints.Add("ALTER TABLE "+Quote(table.Name,dialect)+" ADD "+constraint+";\r\n");
            }
            builder.Append("CREATE TABLE ").Append(Quote(table.Name,dialect)).Append(" (\r\n").Append(string.Join(",\r\n",lines)).Append("\r\n);\r\n\r\n");
        }
        return builder.Append(string.Join("",constraints)).ToString();
    }
}
