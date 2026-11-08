using Sandy.Core;
using System.Net;
namespace Sandy.Tests;
public sealed class RemainingBehaviorTests
{
    [Fact]public async Task DesignerRenameAndIndexKeepSelectionAndUpdateFormAttribute()
    {
        var source="VERSION 5.00\r\nBegin VB.Form F\r\n   Begin VB.Label lbl\r\n      Index = 0\r\n   End\r\nEnd\r\nAttribute VB_Name = \"F\"\r\n";
        var(l,t)=LibraryMutationTests.Fixture("~~SelectControl lbl(0)\r\n~~SetProperty Name=Title\r\n~~SetProperty Index=2\r\n~~SetProperty Caption=Changed\r\n~~SetFormProperty Name=NewForm");var editor=new LegacyEditor("F.frm",source);
        await new Engine().EvaluateAsync(l,t,host:editor);Assert.Contains("Begin VB.Label Title",editor.Active.Text);Assert.Contains("Index = 2",editor.Active.Text);Assert.Contains("Caption = \"Changed\"",editor.Active.Text);Assert.Contains("Attribute VB_Name = \"NewForm\"",editor.Active.Text);Assert.Equal("Title(2)",editor.SelectedControl);
    }
    [Fact]public void AddinListPreservesUnrelatedEntriesAndComments()
    {
        const string ini="[Add-Ins32]\r\n; Sandy.Connect=1\r\nOther.Connect=1\r\n\r\n";
        var added=LegacyAddinList.Update(ini,"Sandy.Connect",false);Assert.Equal(ini+"Sandy.Connect=0\r\n",added);Assert.Equal(added,LegacyAddinList.Update(added,"sandy.connect",false));Assert.Equal(ini,LegacyAddinList.Update(added,"Sandy.Connect",true));
    }
    [Theory][InlineData("On Error=Resume Next","1")][InlineData("textbox.Value == 1","1")][InlineData("&HFF=255","1")][InlineData("1 XOR 1","0")][InlineData("\"a\" & 1 + 2","a3")][InlineData(".5*2","1")]
    public void LegacyExpressionOperandsAndOperators(string expression,string expected)=>Assert.Equal(expected,SandyExpression.Evaluate(expression,new Dictionary<string,string>{{"On Error","Resume Next"},{"textbox.Value","1"}}));
    [Fact]public async Task ControlArraysAndNestedFontsKeepUnchangedBlocksExact()
    {
        var source="VERSION 5.00\r\nBegin VB.Form F\r\n   Begin VB.Label lbl\r\n      Index = 0\r\n      Caption = \"First\"\r\n      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}\r\n         Name = \"Arial\"\r\n         Weight = 400\r\n      EndProperty\r\n   End\r\n   Begin VB.Label lbl\r\n      Index = 1\r\n      Caption = \"Second\"\r\n   End\r\nEnd\r\nAttribute VB_Name = \"F\"\r\n";
        var(l,t)=LibraryMutationTests.Fixture("~~SelectControl lbl(0)\r\n~~SetProperty Font.Weight=700\r\n~~SelectControl lbl(1)\r\n~~SetProperty Caption=Changed");var editor=new LegacyEditor("F.frm",source);
        await new Engine().EvaluateAsync(l,t,host:editor);Assert.Equal(source.Replace("Weight = 400","Weight = 700").Replace("Caption = \"Second\"","Caption = \"Changed\""),editor.Active.Text);Assert.Equal(2,editor.Controls.Count);
    }
    [Fact]public async Task EmptyFileAndTrailingBlankLinesHaveCorrectEof()
    {
        var(l,t)=LibraryMutationTests.Fixture("~~OpenFile input.txt\r\n~~NextLine First\r\n~~NextLine Empty\r\n%%EOF%%:%%First%%");var editor=new LegacyEditor("Example.cs",""){FileReader=_=>"one\r\n\r\n"};
        await new Engine().EvaluateAsync(l,t,host:editor);Assert.Equal("True:one",editor.Active.Text);
        t["memoCodeAtCursor"]="~~OpenFile input.txt\r\n%%EOF%%";editor=new("Example.cs",""){FileReader=_=>""};await new Engine().EvaluateAsync(l,t,host:editor);Assert.Equal("True",editor.Active.Text);
    }
    [Fact]public async Task DirectoryMatchingHasIndependentNextResults()
    {
        var(l,t)=LibraryMutationTests.Fixture("%%MatchFile::*.txt%%:%%NextFile%%:%%NextFile%%");var editor=new LegacyEditor("Example.cs",""){PathReader=(_,_)=>new[]{"b.txt","a.txt"}};
        await new Engine().EvaluateAsync(l,t,host:editor);Assert.Equal("a.txt:b.txt:",editor.Active.Text);
    }
    [Fact]public void ProcessIdsResolveOnlyWhenActionsExecute()
    {
        var seen=new List<ExternalEffect>();var applied=DeferredActions.Apply(new[]{new ExternalEffect("Process","fixture.exe",ResultToken:"<pid>"),new ExternalEffect("Activate","<pid>")},e=>{seen.Add(e);return e.Kind=="Process"?"1234":null;});Assert.Equal("1234",seen[1].Target);Assert.Equal("1234",applied[1].Target);
    }
    [Fact]public void ProfileUpdatesFromSeparateInstancesPreserveOtherSettings()
    {
        var root=Path.Combine(Path.GetTempPath(),"Sandy-profile-"+Guid.NewGuid());var path=Path.Combine(root,"settings.json");try{new ProfileStore(path).Write("one","first");new ProfileStore(path).Write("two","second");Assert.Equal("first",new ProfileStore(path).Read("one"));Assert.Equal("second",new ProfileStore(path).Read("two"));}finally{Directory.Delete(root,true);}
    }
    private sealed class PostHandler:HttpMessageHandler
    {
        public string? Body;public HttpMethod? Method;
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request,CancellationToken cancellationToken){Method=request.Method;Body=await request.Content!.ReadAsStringAsync(cancellationToken);return new(HttpStatusCode.BadRequest){Content=new StringContent("response fixture")};}
    }
    [Fact]public async Task WebPostEncodesFieldsAndDisplaysErrorResponse()
    {
        var handler=new PostHandler();using var client=new HttpClient(handler);var response=await WebAutomation.RequestAsync(client,"https://fixture.test",true,WebAutomation.ParseFields("name=A&B\r\nname=C=D"),1024);
        Assert.Equal(HttpMethod.Post,handler.Method);Assert.Equal("name=A%26B&name=C%3DD",handler.Body);Assert.Equal(400,response.Status);Assert.Equal("response fixture",response.Text);
    }
    [Fact]public void HeaderInspectionKeepsQuotedUrlsAndClassMembers()
    {
        const string header="#define URL \"https://example.test/a//b\"\nclass Example { public: int Get(int value); private: int count; };\n#pragma pack(push,1)\nstruct Small { bool enabled; BYTE bytes[2]; };\n#pragma pack(pop)\n#define BIG (1UL << 32)";
        var inspected=HeaderInspection.Inspect(header);Assert.Contains(inspected.Entries,e=>e.Name=="Example.Get"&&e.Kind=="Method");Assert.Contains(inspected.Entries,e=>e.Name=="Example.count"&&e.Declaration.StartsWith("private"));
        var generated=HeaderGenerator.Generate(header);Assert.Contains("https://example.test/a//b",generated.Code);Assert.Contains("Pack = 1",generated.Code);Assert.Contains("const ulong BIG",generated.Code);
    }
}
