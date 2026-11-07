using System.Reflection;
using System.Runtime.Versioning;
using Sandy.Core;

namespace Sandy.Core.Tests;

public class MetXIntegrationTests
{
    [Fact]
    public void CoreReferencesAndLoadsTheMatchingMetXAssembly()
    {
        var reference = typeof(TemplateEngine).Assembly.GetReferencedAssemblies().Single(a => a.Name == "MetX.Standard.Strings");
        var assembly = Assembly.Load(reference);
        var framework = assembly.GetCustomAttribute<TargetFrameworkAttribute>()!.FrameworkName;
#if STANDARD21_CORE
        Assert.Equal(".NETStandard,Version=v2.1", framework);
#elif STANDARD_CORE
        Assert.Equal(".NETStandard,Version=v2.0", framework);
#else
        Assert.Equal(".NETCoreApp,Version=v10.0", framework);
#endif
    }

    [Theory]
    [InlineData("Left::abc,2", "ab")]
    [InlineData("Left::abc,99", "abc")]
    [InlineData("Left::abc,-1", "")]
    [InlineData("Right::abc,2", "bc")]
    [InlineData("Right::abc,0", "")]
    [InlineData("Mid::abc,1,1", "a")]
    [InlineData("Mid::abc,0,2", "ab")]
    [InlineData("Mid::abc,2", "bc")]
    [InlineData("Mid::abc,99,3", "")]
    [InlineData("Mid::abc,2,99", "bc")]
    [InlineData("Mid::abc,2,-1", "")]
    [InlineData("Token::a||b,2,|", "")]
    [InlineData("Token::a||b,3,|", "b")]
    [InlineData("Token::a|,2,|", "")]
    [InlineData("Token::abc,99,|", "")]
    [InlineData("Token::abc,0,|", "abc")]
    [InlineData("Token::aXXbxxc,2,xx", "c")]
    [InlineData("Before::aXXbxxc,2,xx", "aXXb")]
    [InlineData("After::a||b|c,2,|", "b|c")]
    [InlineData("After::a|b||c,2,|", "|c")]
    [InlineData("Before::a||b,3,|", "a|")]
    [InlineData("Except::a||b,2,|", "a|b")]
    [InlineData("TokenCount::a||b,|", "3")]
    [InlineData("TokenCount::aXXbxxc,xx", "2")]
    [InlineData("TokenCount::Empty,|", "0")]
    [InlineData("FindToken::a|A,A,|", "2")]
    public async Task MetXBackedOperationsPreserveSandyBoundaries(string expression, string expected)
    {
        var (library, template) = EngineTests.Library("%%" + expression + "%%");
        var options = new ExecutionOptions(); options.Variables["Empty"] = "";
        var result = await new TemplateEngine().ExecuteAsync(library, template, options);
        Assert.Equal(expected + "\r\n", result.Outputs.Single().Text);
    }

    [Fact]
    public async Task ExtractUpdatesItsSourceAndEmptyDelimitersStillFailClearly()
    {
        var (library, template) = EngineTests.Library("~~Set Parts=a||b\n%%Extract::Parts,2,|%%\n%%Parts%%");
        var result = await new TemplateEngine().ExecuteAsync(library, template);
        Assert.Equal("a|b", result.Variables["Parts"]);
        Assert.Equal("\r\na|b\r\n", result.Outputs.Single().Text);
        template.Fields["memoCodeAtCursor"] = "%%TokenCount::a,%%";
        var error = await Assert.ThrowsAsync<SandyExecutionException>(() => new TemplateEngine().ExecuteAsync(library, template));
        Assert.Contains("delimiter cannot be empty", error.Message);
    }
}
