using Sandy.Core;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
namespace Sandy.Tests;
public sealed class DatabaseDesignTests
{
    internal static DatabaseDesign Fixture()
    {
        var design=new DatabaseDesign();var parent=design.AddTable("Customers");parent.Fields.Add(new(){Name="Name",DaoType=10,Length=100,Nullable=false});
        var child=design.AddTable("Orders");child.Fields.Add(new(){Name="CustomersID",DaoType=4,Length=0});child.Fields.Add(new(){Name="Amount",DaoType=5});
        design.ApplyRules(true,true,true,true);design.Relations.Add(new("Customers","CustomersID","Orders","CustomersID",true,true));return design;
    }
    [Fact]public void RulesAndRelationshipsRoundTripAndGenerateCascadeSql()
    {
        var design=DatabaseDesign.FromJson(Fixture().ToJson());Assert.Equal(2,design.Branch("Customers").Count);
        var sql=design.GenerateSql(SqlDialect.Sqlite);Assert.Contains("INTEGER PRIMARY KEY AUTOINCREMENT",sql);Assert.Contains("ON DELETE CASCADE ON UPDATE CASCADE",sql);Assert.Contains("DateModified",sql);
        design.RemoveTable(design.Tables[0]);Assert.Empty(design.Relations);
    }
    [Fact]public void ReapplyingRulesDoesNotDuplicateFields()
    {
        var design=Fixture();var count=design.Tables.Sum(t=>t.Fields.Count);design.ApplyRules(true,true,true,true);Assert.Equal(count,design.Tables.Sum(t=>t.Fields.Count));
    }
    [Theory][InlineData(SqlDialect.Access)][InlineData(SqlDialect.SqlServer)][InlineData(SqlDialect.Sqlite)]
    public void GeneratedModelsCollectionsAndRepositoriesCompileTogether(SqlDialect dialect)
    {
        var files=RepositoryGenerator.GenerateSet(Fixture(),"Generated.Models",dialect);
        var refs=((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")!).Split(Path.PathSeparator).Select(path=>MetadataReference.CreateFromFile(path));
        var compilation=CSharpCompilation.Create("Generated",files.Where(f=>f.Key.EndsWith(".cs")).Select(f=>CSharpSyntaxTree.ParseText(f.Value,path:f.Key)),refs,new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        var errors=compilation.GetDiagnostics().Where(d=>d.Severity==DiagnosticSeverity.Error).ToArray();Assert.True(errors.Length==0,string.Join(Environment.NewLine,errors.Select(e=>e.ToString())));
        Assert.Contains("CreateParameter()",files["CustomersRepository.cs"]);Assert.Contains("DBNull.Value",files["CustomersRepository.cs"]);
        Assert.Contains(dialect==SqlDialect.Access?" = ?":" = @p0",files["CustomersRepository.cs"]);
    }
    [Fact]public void IdentifierAndRelationshipConflictsFailBeforeGeneration()
    {
        var design=Fixture();design.Tables[0].Fields.Add(new(){Name="Na-me"});design.Tables[0].Fields.Add(new(){Name="Na me"});Assert.Throws<InvalidDataException>(design.Validate);
        design=Fixture();design.Tables[1].Fields[0].DaoType=10;Assert.Throws<InvalidDataException>(design.Validate);
    }
    [Fact]public async Task GeneratedRepositoryPerformsRealSqliteCrudAndCascadeDelete()
    {
        var design=Fixture();var files=RepositoryGenerator.GenerateSet(design,"Generated.Models",SqlDialect.Sqlite);
        var refs=((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")!).Split(Path.PathSeparator).Select(path=>MetadataReference.CreateFromFile(path));
        var compilation=CSharpCompilation.Create("Generated"+Guid.NewGuid().ToString("N"),files.Where(f=>f.Key.EndsWith(".cs")).Select(f=>CSharpSyntaxTree.ParseText(f.Value)),refs,new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        using var stream=new MemoryStream();var emitted=compilation.Emit(stream);Assert.True(emitted.Success,string.Join("\n",emitted.Diagnostics));
        var assembly=System.Reflection.Assembly.Load(stream.ToArray());var model=assembly.GetType("Generated.Models.Customers")!;var repository=assembly.GetType("Generated.Models.CustomersRepository")!;
        var path=Path.Combine(Path.GetTempPath(),"Sandy-crud-"+Guid.NewGuid()+".db");
        try
        {
            string connectionString="Data Source="+path+";Foreign Keys=True;Pooling=False";
            await using(var connection=new Microsoft.Data.Sqlite.SqliteConnection(connectionString)){await connection.OpenAsync();using var command=connection.CreateCommand();command.CommandText=files["schema.sql"];await command.ExecuteNonQueryAsync();}
            Func<System.Data.Common.DbConnection> factory=()=>new Microsoft.Data.Sqlite.SqliteConnection(connectionString);
            var repo=Activator.CreateInstance(repository,factory)!;var row=Activator.CreateInstance(model)!;model.GetProperty("Name")!.SetValue(row,"O'Brien; DROP TABLE Customers; --");
            await Call("InsertAsync",row,CancellationToken.None);var key=(int)model.GetProperty("CustomersID")!.GetValue(row)!;Assert.True(key>0);
            var found=await Call("FindAsync",key,CancellationToken.None);Assert.Equal("O'Brien; DROP TABLE Customers; --",model.GetProperty("Name")!.GetValue(found));
            model.GetProperty("Name")!.SetValue(row,"Updated");Assert.Equal(1,await Call("UpdateAsync",row,CancellationToken.None));
            await using(var connection=new Microsoft.Data.Sqlite.SqliteConnection(connectionString)){await connection.OpenAsync();using var command=connection.CreateCommand();command.CommandText="INSERT INTO Orders (CustomersID,Amount,DateCreated,DateModified) VALUES (@id,12,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)";command.Parameters.AddWithValue("@id",key);await command.ExecuteNonQueryAsync();}
            Assert.Equal(1,await Call("DeleteAsync",key,CancellationToken.None));Assert.Null(await Call("FindAsync",key,CancellationToken.None));
            await using(var connection=new Microsoft.Data.Sqlite.SqliteConnection(connectionString)){await connection.OpenAsync();using var command=connection.CreateCommand();command.CommandText="SELECT COUNT(*) FROM Orders";Assert.Equal(0L,await command.ExecuteScalarAsync());}
            async Task<object?> Call(string method,params object[] parameters){var task=(Task)repository.GetMethod(method)!.Invoke(repo,parameters)!;await task;return task.GetType().GetProperty("Result")?.GetValue(task);}
        }
        finally{File.Delete(path);}
    }
}
