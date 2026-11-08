using System.Text;
using System.Text.Json;
namespace Sandy.Core;

public static class RepositoryGenerator
{
    internal static string Type(int type)=>type switch{1=>"bool",2=>"byte",3=>"short",4=>"int",5 or 20=>"decimal",6=>"float",7=>"double",8=>"DateTime",9 or 11=>"byte[]",10 or 12=>"string",15=>"Guid",16=>"long",_=>"object"};
    private static string DbType(int type)=>type switch{1=>"Boolean",2=>"Byte",3=>"Int16",4=>"Int32",5 or 20=>"Decimal",6=>"Single",7=>"Double",8=>"DateTime",9 or 11=>"Binary",10 or 12=>"String",15=>"Guid",16=>"Int64",_=>"Object"};
    private static string Literal(string value)=>JsonSerializer.Serialize(value);
    private static string Id(string value)=>DatabaseGenerator.Identifier(value);
    private static string ConvertValue(int type,string expression)=>type switch {
        9 or 11=>"(byte[])"+expression,15=>"Guid.Parse(Convert.ToString("+expression+", CultureInfo.InvariantCulture)!)",0=>expression,
        _=>"Convert.To"+(type switch{1=>"Boolean",2=>"Byte",3=>"Int16",4=>"Int32",5 or 20=>"Decimal",6=>"Single",7=>"Double",8=>"DateTime",10 or 12=>"String",16=>"Int64",_=>"String"})+"("+expression+", CultureInfo.InvariantCulture)"+(type is 10 or 12?"!":"")
    };
    public static string Generate(DatabaseTable table,string nameSpace,SqlDialect dialect)
    {
        new DatabaseDesign{Tables=[table]}.Validate();
        var key=table.Fields.SingleOrDefault(f=>f.PrimaryKey)??throw new InvalidOperationException("Choose a primary key for "+table.Name+" before generating a repository.");
        var type=Id(table.Name);var repository=Id(table.Name+"Repository");var b=new StringBuilder("#nullable enable\r\nusing System;\r\nusing System.Collections.Generic;\r\nusing System.Data;\r\nusing System.Data.Common;\r\nusing System.Globalization;\r\nusing System.Threading;\r\nusing System.Threading.Tasks;\r\n\r\nnamespace "+string.Join(".",nameSpace.Split('.').Select(Id))+";\r\n\r\npublic sealed class "+repository+"(Func<DbConnection> createConnection)\r\n{\r\n");
        string Q(string value)=>DatabaseDesign.Quote(value,dialect);
        string Marker(int n)=>dialect==SqlDialect.Access?"?":"@p"+n;
        var select="SELECT "+string.Join(", ",table.Fields.Select(f=>Q(f.Name)))+" FROM "+Q(table.Name);
        b.Append("    private async Task<DbConnection> OpenAsync(CancellationToken cancellation)\r\n    {\r\n        var connection = createConnection();\r\n        try { if (connection.State != ConnectionState.Open) await connection.OpenAsync(cancellation); return connection; }\r\n        catch { await connection.DisposeAsync(); throw; }\r\n    }\r\n");
        b.Append("    private static void Bind(DbCommand command, string name, DbType type, object? value, int size = 0)\r\n    {\r\n        var parameter = command.CreateParameter(); parameter.ParameterName = name; parameter.DbType = type; parameter.Value = value ?? DBNull.Value; if (size > 0) parameter.Size = size; command.Parameters.Add(parameter);\r\n    }\r\n");
        b.Append("    private static ").Append(type).Append(" Read(DbDataReader reader) => new()\r\n    {\r\n");
        for(var i=0;i<table.Fields.Count;i++) {
            var f=table.Fields[i];var value=ConvertValue(f.DaoType,"reader.GetValue("+i+")");
            b.Append("        ").Append(Id(f.Name)).Append(" = reader.IsDBNull(").Append(i).Append(") ? ").Append(f.Nullable&&!f.PrimaryKey?"null":"throw new DataException("+Literal("Unexpected NULL in "+f.Name)+")").Append(" : ").Append(value).Append(",\r\n");
        }
        b.Append("    };\r\n");
        b.Append("    public async Task<IReadOnlyList<").Append(type).Append(">> GetAllAsync(CancellationToken cancellation = default)\r\n    {\r\n        await using var connection = await OpenAsync(cancellation); await using var command = connection.CreateCommand();\r\n        command.CommandText = ").Append(Literal(select+" ORDER BY "+Q(key.Name))).Append(";\r\n        await using var reader = await command.ExecuteReaderAsync(cancellation); var rows = new List<").Append(type).Append(">();\r\n        while (await reader.ReadAsync(cancellation)) rows.Add(Read(reader)); return rows;\r\n    }\r\n");
        b.Append("    public async Task<").Append(type).Append("?> FindAsync(").Append(Type(key.DaoType)).Append(" key, CancellationToken cancellation = default)\r\n    {\r\n        await using var connection = await OpenAsync(cancellation); await using var command = connection.CreateCommand();\r\n        command.CommandText = ").Append(Literal(select+" WHERE "+Q(key.Name)+" = "+Marker(0))).Append(";\r\n");
        Bind(key,"key",0);b.Append("        await using var reader = await command.ExecuteReaderAsync(cancellation); return await reader.ReadAsync(cancellation) ? Read(reader) : null;\r\n    }\r\n");
        var insert=table.Fields.Where(f=>!f.AutoNumber).ToArray();
        if(insert.Length==0&&dialect==SqlDialect.Access)throw new NotSupportedException("An Access repository needs a non-AutoNumber field to insert.");
        var insertSql=insert.Length==0?"INSERT INTO "+Q(table.Name)+" DEFAULT VALUES":"INSERT INTO "+Q(table.Name)+" ("+string.Join(", ",insert.Select(f=>Q(f.Name)))+") VALUES ("+string.Join(", ",insert.Select((_,i)=>Marker(i)))+")";
        var scalarInBatch=key.AutoNumber&&dialect==SqlDialect.SqlServer;
        b.Append("    public async Task<").Append(type).Append("> InsertAsync(").Append(type).Append(" value, CancellationToken cancellation = default)\r\n    {\r\n");
        Dates(false);b.Append("        await using var connection = await OpenAsync(cancellation); await using var command = connection.CreateCommand();\r\n        command.CommandText = ").Append(Literal(insertSql+(scalarInBatch?"; SELECT CAST(SCOPE_IDENTITY() AS bigint);":""))).Append(";\r\n");
        for(var i=0;i<insert.Length;i++)Bind(insert[i],"value."+Id(insert[i].Name),i);
        if(scalarInBatch)b.Append("        var identity = await command.ExecuteScalarAsync(cancellation) ?? throw new DataException(\"No generated identity.\");\r\n");
        else b.Append("        await command.ExecuteNonQueryAsync(cancellation);\r\n");
        if(key.AutoNumber) {
            if(!scalarInBatch)b.Append("        command.Parameters.Clear(); command.CommandText = ").Append(Literal(dialect==SqlDialect.Sqlite?"SELECT last_insert_rowid();":"SELECT @@IDENTITY;")).Append(";\r\n        var identity = await command.ExecuteScalarAsync(cancellation) ?? throw new DataException(\"No generated identity.\");\r\n");
            b.Append("        value.").Append(Id(key.Name)).Append(" = ").Append(ConvertValue(key.DaoType,"identity")).Append(";\r\n");
        }
        b.Append("        return value;\r\n    }\r\n");
        var update=table.Fields.Where(f=>!f.PrimaryKey&&!f.AutoNumber&&!f.Name.Equals("DateCreated",StringComparison.OrdinalIgnoreCase)).ToArray();
        b.Append("    public async Task<int> UpdateAsync(").Append(type).Append(" value, CancellationToken cancellation = default)\r\n    {\r\n");
        if(update.Length==0)b.Append("        await Task.CompletedTask; throw new InvalidOperationException(\"No updatable fields.\");\r\n");
        else {
            Dates(true);b.Append("        await using var connection = await OpenAsync(cancellation); await using var command = connection.CreateCommand();\r\n        command.CommandText = ").Append(Literal("UPDATE "+Q(table.Name)+" SET "+string.Join(", ",update.Select((f,i)=>Q(f.Name)+" = "+Marker(i)))+" WHERE "+Q(key.Name)+" = "+Marker(update.Length))).Append(";\r\n");
            for(var i=0;i<update.Length;i++)Bind(update[i],"value."+Id(update[i].Name),i);Bind(key,"value."+Id(key.Name),update.Length);b.Append("        return await command.ExecuteNonQueryAsync(cancellation);\r\n");
        }
        b.Append("    }\r\n    public async Task<int> DeleteAsync(").Append(Type(key.DaoType)).Append(" key, CancellationToken cancellation = default)\r\n    {\r\n        await using var connection = await OpenAsync(cancellation); await using var command = connection.CreateCommand();\r\n        command.CommandText = ").Append(Literal("DELETE FROM "+Q(table.Name)+" WHERE "+Q(key.Name)+" = "+Marker(0))).Append(";\r\n");
        Bind(key,"key",0);b.Append("        return await command.ExecuteNonQueryAsync(cancellation);\r\n    }\r\n}\r\n");return b.ToString();

        void Bind(DatabaseField field,string value,int index)=>b.Append("        Bind(command, \"p").Append(index).Append("\", DbType.").Append(DbType(field.DaoType)).Append(", ").Append(value).Append(", ").Append(field.DaoType==10?field.Length:0).Append(");\r\n");
        void Dates(bool updating) { foreach(var field in table.Fields.Where(f=>f.DaoType==8)) { if(field.Name.Equals("DateModified",StringComparison.OrdinalIgnoreCase))b.Append("        value.").Append(Id(field.Name)).Append(" = DateTime.UtcNow;\r\n");else if(!updating&&field.Name.Equals("DateCreated",StringComparison.OrdinalIgnoreCase))b.Append("        if (value.").Append(Id(field.Name)).Append(" == default) value.").Append(Id(field.Name)).Append(" = DateTime.UtcNow;\r\n"); } }
    }
    public static IReadOnlyDictionary<string,string> GenerateSet(DatabaseDesign design,string nameSpace,SqlDialect dialect,IEnumerable<DatabaseTable>? selected=null,bool repositories=true)
    {
        design.Validate();var tables=(selected??design.Tables.Where(t=>t.Marked)).ToArray();var files=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        var generatedNames=tables.SelectMany(t=>repositories?new[]{Id(t.Name),Id(t.Name+"Collection"),Id(t.Name+"Repository")}:new[]{Id(t.Name),Id(t.Name+"Collection")}).ToArray();
        if(generatedNames.Distinct(StringComparer.OrdinalIgnoreCase).Count()!=generatedNames.Length)throw new InvalidDataException("Generated class/file names collide. Rename the affected tables.");
        foreach(var table in tables)
        {
            var model=DatabaseGenerator.Generate(table.Name,table.Fields.Select(f=>f.Schema),nameSpace);
            var navigation=new StringBuilder();
            foreach(var child in design.Relations.Where(r=>r.ParentTable==table.Name).Select(r=>r.ChildTable).Distinct())if(tables.Any(t=>t.Name==child)) {
                if(table.Fields.Any(f=>Id(f.Name)==Id(child+"Children")))throw new InvalidDataException("Relationship property conflicts with field "+child+"Children");
                navigation.Append("    public System.Collections.Generic.List<").Append(Id(child)).Append("> ").Append(Id(child+"Children")).Append(" { get; } = new();\r\n");
            }
            var close=model.LastIndexOf('}');model=model.Insert(close,navigation.ToString());files[Id(table.Name).TrimStart('@')+".cs"]=model;
            files[Id(table.Name+"Collection").TrimStart('@')+".cs"]="#nullable enable\r\nnamespace "+string.Join(".",nameSpace.Split('.').Select(Id))+";\r\npublic sealed class "+Id(table.Name+"Collection")+" : System.Collections.ObjectModel.Collection<"+Id(table.Name)+"> { }\r\n";
            if(repositories)files[Id(table.Name+"Repository").TrimStart('@')+".cs"]=Generate(table,nameSpace,dialect);
        }
        files["schema.sql"]=design.GenerateSql(dialect,tables);files["database.sandy-schema.json"]=design.ToJson();return files;
    }
}
