using Sandy.Core;
using Microsoft.CodeAnalysis.CSharp;
namespace Sandy.Tests;
public sealed class GeneratorTests
{
    [Fact] public void DatabaseModelsMapTypesAndEscapeNames()
    {
        var code=DatabaseGenerator.Generate("123 Order",[new("class",4),new("Total",5),new("Title",10,50,false)]);
        Assert.Contains("class _123_Order",code);Assert.Contains("int? @class",code);Assert.Contains("decimal? Total",code);Assert.Contains("string Title { get; set; } = string.Empty;",code);
        Assert.DoesNotContain(CSharpSyntaxTree.ParseText(code).GetDiagnostics(),d=>d.Severity==Microsoft.CodeAnalysis.DiagnosticSeverity.Error);
    }
    [Fact] public void HeaderConversionProducesStructsUnionsEnumsAndArrays()
    {
        var result=HeaderGenerator.Generate("#define LIMIT 12\nstruct Point { int x; int y; }; union Bits { int number; float fraction; }; enum Kind { One=1, Two=2 }; struct Buffer { BYTE data[8]; };");
        Assert.Empty(result.Diagnostics);Assert.Contains("FieldOffset(0)",result.Code);Assert.Contains("ByValArray, SizeConst = 8",result.Code);Assert.Contains("public enum Kind",result.Code);
        Assert.DoesNotContain(CSharpSyntaxTree.ParseText(result.Code).GetDiagnostics(),d=>d.Severity==Microsoft.CodeAnalysis.DiagnosticSeverity.Error);
    }
    [Fact] public void UnknownHeaderFieldsAreReported()
    {
        var result=HeaderGenerator.Generate("struct Test { UnknownThing value; };");Assert.Contains(result.Diagnostics,x=>x.Contains("UnknownThing"));
    }
    [Theory]
    [InlineData("")][InlineData("ABC")][InlineData("Sample legacy text with accents: café")]
    public void OriginalObfuscationRoundTrips(string text)=>Assert.Equal(text,LegacyObfuscation.Decrypt(LegacyObfuscation.Encrypt(text)));
    [Fact] public void ObfuscationMatchesCharacterScrambleOrder()=>Assert.Equal("EN* DBEBBC",LegacyObfuscation.Encrypt("ABC"));
    [Fact] public void CommandDocumentationKeepsAliasesSyntaxAndExamples()
    {
        var entries=CommandReference.Parse("** Command Set\r\n** Alias SetValue, Let\r\n** Syntax Name=Value\r\n** Short Assign a value\r\n** Example\r\n~~Set Name=William\r\n** '\r\n** Inline Len\r\n** Short Get length\r\n","fixture");
        Assert.Equal(2,entries.Count);Assert.Equal("SetValue, Let",entries[0].Aliases);Assert.Equal("Name=Value",entries[0].Syntax);Assert.Contains("~~Set Name=William",entries[0].Examples);
    }
}
