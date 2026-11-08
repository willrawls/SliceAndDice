using Sandy.Core;
namespace Sandy.Tests;
public sealed class LegacyExecutionTests
{
    private static (Library,Template) Fixture(string source)
    {
        var library=new Library();var c=new Category();c.Fields["CategoryID"]="1";c.Fields["CategoryName"]="Test";library.Categories.Add(c);
        var t=new Template();t["TemplateID"]="1";t["CategoryID"]="1";t["TemplateName"]="Test - Run";t["memoCodeAtCursor"]=source;c.Templates.Add(t);return(library,t);
    }
    [Fact] public async Task NestedConditionalsDoNotPromptForSkippedVariables()
    {
        var(l,t)=Fixture("~~Set n=2\r\n~~If n > 1\r\n~~If n = 2\r\nyes\r\n~~Else\r\n%%NeverAsk%%\r\n~~EndIf\r\n~~Else\r\nno\r\n~~EndIf");
        var r=await new Engine().EvaluateAsync(l,t,prompt:(_,_)=>throw new Exception("Unexpected prompt"));Assert.Equal("yes\r\n",r.Outputs.Single().Text);
    }
    [Fact] public async Task CodaArithmeticAndIncrementAmountsExecute()
    {
        var(l,t)=Fixture("~~Set n=0\r\n~~StartCoda\r\n~~Inc n 2\r\n~~Eval more=n < 6\r\n%%n%%\r\n~~Coda more");
        var r=await new Engine().EvaluateAsync(l,t);Assert.Equal("2\r\n4\r\n6\r\n",r.Outputs.Single().Text);
    }
    [Fact] public async Task EditorNavigationUsesPendingGeneratedText()
    {
        var(l,t)=Fixture("~~AbsLine 2\r\n~~ReadLine Current\r\n~~ReplaceLine %%UCase::Current%%\r\n~~GotoEnd~~\r\nlast");
        var editor=new LegacyEditor("sample.bas","first\r\nsecond\r\n");await new Engine().EvaluateAsync(l,t,host:editor);
        Assert.Equal("first\r\nSECOND\r\nlast",editor.Active.Text);Assert.Equal("first\r\nsecond\r\n",editor.Active.Original);
    }
    [Fact] public async Task FilesArePlannedAndNeverWrittenByEvaluation()
    {
        var(l,t)=Fixture("~~CreateFile NeverWrite.txt,Hello\r\n~~GotoFile Other.txt\r\nworld");
        var editor=new LegacyEditor("sample.cs","");await new Engine().EvaluateAsync(l,t,host:editor);
        Assert.Equal(2,editor.Effects.Count);Assert.Equal("FileOverwrite",editor.Effects[0].Kind);Assert.Equal("FileAppend",editor.Effects[1].Kind);
    }
    [Fact] public async Task MultiAreaOutputUsesTypeBoundaries()
    {
        var(l,t)=Fixture("public int P { get; set; }\r\n");t["memoCodeAtTop"]="private int field;\r\n";t["memoCodeAtBottom"]="public void Last() {}\r\n";
        var editor=new LegacyEditor("sample.cs","class Sample\r\n{\r\n}\r\n",17);await new Engine().EvaluateAsync(l,t,host:editor);
        Assert.Contains("private int field;",editor.Active.Text);Assert.Contains("public void Last() {}\r\n}",editor.Active.Text);
    }
    [Fact] public async Task DesignerDefaultsAreAppliedToGeneratedForm()
    {
        var(l,t)=Fixture("~~SetDefault VB.Label Left=120\r\n~~GotoForm Example\r\n~~AddControl VB.Label lblTitle\r\n~~SetProperty Caption=\"Title\"\r\n~~SetFormProperty Width=6000");
        var editor=new LegacyEditor("sample.bas","");await new Engine().EvaluateAsync(l,t,host:editor);
        Assert.Contains("Begin VB.Label lblTitle",editor.Active.Text);Assert.Contains("Left = 120",editor.Active.Text);Assert.Contains("Caption = \"Title\"",editor.Active.Text);
    }
    [Fact] public async Task FileStreamsReadPlannedWritesWithoutWritingDisk()
    {
        var(l,t)=Fixture("~~CreateFile data.txt,First\r\n~~AppendToFile data.txt,Second\r\n~~OpenForInput data.txt\r\n~~NextLine Item\r\n%%Item%%\r\n%%EOF%%");
        var editor=new LegacyEditor("sample.cs","");var r=await new Engine().EvaluateAsync(l,t,host:editor);
        Assert.Contains("FirstSecond",editor.Active.Text);Assert.EndsWith("True",editor.Active.Text);
    }
    [Fact] public async Task WordWrapAndSubstringCommandsProduceOutput()
    {
        var(l,t)=Fixture("~~Left 3 Prefix=abcdef\r\n~~Right 2 Suffix=abcdef\r\n~~Mid 2 3 Center=abcdef\r\n%%Prefix%%:%%Suffix%%:%%Center%%\r\n~~WordWrap 9 \"// \" one two three four");
        var r=await new Engine().EvaluateAsync(l,t);Assert.Equal("abc:ef:bcd\r\n// one two\r\n// three\r\n// four\r\n",r.Outputs.Single().Text);
    }
    [Fact] public async Task MessageIsDeliveredToHostWithoutChangingEditor()
    {
        var(l,t)=Fixture("~~Set Greeting=Hello\r\n~~MsgBox Greeting");var editor=new LegacyEditor("sample.cs","original");
        await new Engine().EvaluateAsync(l,t,host:editor);Assert.False(editor.Active.Changed);Assert.Equal(new ExternalEffect("Message","","Hello"),editor.Effects.Single());
    }
    [Fact] public async Task ExistingDesignerControlIsEditableAndFrxReferenceSurvives()
    {
        var(l,t)=Fixture("~~SelectControl lblTitle\r\n~~SetProperty Caption=Updated");
        var source="VERSION 5.00\r\nBegin VB.Form Sample\r\n   Picture = \"Sample.frx\":0000\r\n   Begin VB.Label lblTitle\r\n      Caption = \"Old\"\r\n   End\r\nEnd\r\nAttribute VB_Name = \"Sample\"\r\n";
        var editor=new LegacyEditor("Sample.frm",source);await new Engine().EvaluateAsync(l,t,host:editor);
        Assert.Contains("Caption = \"Updated\"",editor.Active.Text);Assert.Contains("Picture = \"Sample.frx\":0000",editor.Active.Text);
    }
    [Fact] public async Task NewModuleHasEmptyOriginalForFileCommit()
    {
        var(l,t)=Fixture("~~GotoClass NewClass\r\n~~GotoDeclarations\r\nprivate int value;");var editor=new LegacyEditor("Sample.cs","");
        await new Engine().EvaluateAsync(l,t,host:editor);Assert.Equal("",editor.Active.Original);Assert.Contains("class NewClass",editor.Active.Text);
    }
    [Fact] public void DeepExpressionsFailWithControlledError()
    {
        Assert.Throws<FormatException>(()=>SandyExpression.Evaluate(new string('(',2000)+"1"+new string(')',2000),new Dictionary<string,string>()));
    }
    [Fact] public async Task ProjectReferencesModifyAPlanAndPreserveProjectSettings()
    {
        var(l,t)=Fixture("~~AddReference System.Xml\r\n~~AddReference System.Xml");var root=Path.Combine(Path.GetTempPath(),"SandyProject");var editor=new LegacyEditor(Path.Combine(root,"Code.cs"),"");
        var xml="<Project Sdk=\"Microsoft.NET.Sdk\"><PropertyGroup><TargetFramework>net10.0</TargetFramework><Nullable>enable</Nullable></PropertyGroup></Project>";
        var path=Path.Combine(root,"Sample.csproj");editor.AttachProject(path,xml);await new Engine().EvaluateAsync(l,t,host:editor);
        var result=System.Xml.Linq.XDocument.Parse(editor.Documents[path].Text);Assert.Single(result.Descendants("Reference"));Assert.Equal("enable",result.Descendants("Nullable").Single().Value);Assert.Equal(xml,editor.Documents[path].Original);
    }
    [Fact] public async Task ForeachControlExpandsPropertiesOfExistingControls()
    {
        var(l,t)=Fixture("~~ForEachControl\r\n**Label\r\n**Control.Name**=**Caption**\r\n~~NoInsert");
        var editor=new LegacyEditor("Sample.frm","VERSION 5.00\r\nBegin VB.Form Sample\r\n   Begin VB.Label lblTitle\r\n      Caption = \"Hello\"\r\n   End\r\nEnd\r\n",0);
        await new Engine().EvaluateAsync(l,t,host:editor);Assert.StartsWith("lblTitle=Hello\r\n",editor.Active.Text);
    }
    [Fact] public async Task InlineIncludeExecutesIncludedSoftCommands()
    {
        var(l,t)=Fixture("%%Include::2,Part%%");var part=new Template();part["TemplateID"]="2";part["CategoryID"]="1";part["TemplateName"]="Part";part["ShortTemplateName"]="Part";part["memoCodeAtCursor"]="~~Set Name=Included\r\n%%Name%%";l.Categories[0].Templates.Add(part);
        var r=await new Engine().EvaluateAsync(l,t);Assert.Equal("Included",r.Outputs.Single().Text);
    }
    [Fact] public async Task MangleAndFlipFlopFollowLegacyBehavior()
    {
        var(l,t)=Fixture("~~Set Input=left=right\r\n~~ReverseTokens Input Output \"=\"\r\n%%Mangle::CUSTOMER_NAME%%:%%Output%%");var editor=new LegacyEditor("Sample.cs","");
        await new Engine().EvaluateAsync(l,t,host:editor);Assert.Equal("CustomerName:right=left",editor.Active.Text);
    }
    [Fact] public async Task PlannedSettingsAreReadableBeforeCommit()
    {
        var(l,t)=Fixture("~~SaveSetting Test,Key,Value\r\n~~GetSetting Result,Test,Key,Fallback\r\n%%Result%%:%%GetSetting::Test,Key,Fallback%%");var editor=new LegacyEditor("Sample.cs","");
        await new Engine().EvaluateAsync(l,t,host:editor);Assert.Equal("Value:Value",editor.Active.Text);Assert.Equal("Setting",editor.Effects.Single().Kind);
    }
    [Fact] public async Task DynamicCommandsAndNestedSoftCodeShareVariables()
    {
        var(l,t)=Fixture("~~Set Command=Set\r\n~~%%Command%% Answer=42\r\n~~RunSoftCode Snippet\r\n%%Answer%%");
        var r=await new Engine().EvaluateAsync(l,t,new Dictionary<string,string>{{"Snippet","~~Eval Answer=Answer + 1"}});Assert.Equal("43",r.Outputs.Single().Text);
    }
    [Fact] public async Task NamedTokenListsKeepIndependentPositions()
    {
        var(l,t)=Fixture("~~Set ItemDelimiter=,\r\n~~Parse A=one,two\r\n~~Parse B=three,four\r\n%%NextItem::A%%:%%NextItem::B%%:%%NextItem::A%%:%%Next Item in A%%");var editor=new LegacyEditor("Sample.cs","");
        await new Engine().EvaluateAsync(l,t,host:editor);Assert.Equal("one:three:two:0",editor.Active.Text);
    }
    [Fact] public async Task CommentedParametersUsesContainingCSharpSignature()
    {
        var(l,t)=Fixture("~~CommentedParameters Description\r\n%%Description%%:%%Function Returns%%");var source="class C { public int Add(int a, int b = 1) { return a+b; } }";
        var editor=new LegacyEditor("Sample.cs",source,source.IndexOf("return",StringComparison.Ordinal));
        var result=await new Engine().EvaluateAsync(l,t,host:editor);Assert.Contains("a||(I) int$$b||(I) int Defaults to 1:int",editor.Active.Text);
    }
}
