using Sandy.Core;
namespace Sandy.Tests;
public sealed class WebContentTests
{
    private sealed class Handler(byte[] bytes):HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request,CancellationToken cancellationToken)=>Task.FromResult(new HttpResponseMessage(System.Net.HttpStatusCode.OK){Content=new ByteArrayContent(bytes)});
    }
    [Fact] public async Task DownloadPreservesBinaryBytesAndRejectsOversizeResponse()
    {
        var bytes=new byte[]{0,255,128,13,10};using var client=new HttpClient(new Handler(bytes));
        Assert.Equal(bytes,await WebContent.FetchAsync(client,new Uri("https://fixture.test/file"),10));
        await Assert.ThrowsAsync<InvalidDataException>(()=>WebContent.FetchAsync(client,new Uri("https://fixture.test/file"),2));
        await Assert.ThrowsAsync<ArgumentException>(()=>WebContent.FetchAsync(client,new Uri("file:///C:/test"),10));
    }
}
