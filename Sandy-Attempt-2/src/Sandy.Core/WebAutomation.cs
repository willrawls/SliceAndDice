using System.Text;
namespace Sandy.Core;
public sealed record WebFormField(string Name,string Value);
public sealed record WebResponse(string Url,int Status,string ContentType,byte[] Body)
{
    public string Text=>TextFiles.Decode(Body);
}
public static class WebAutomation
{
    private static readonly HttpClient Client=new(){Timeout=TimeSpan.FromSeconds(30)};
    public static IReadOnlyList<WebFormField> ParseFields(string text)=>text.Replace("\r\n","\n").Split('\n').Where(line=>line.Length>0).Select(line=>{
        var equals=line.IndexOf('=');if(equals<1)throw new FormatException("POST fields require Name=Value, one per line.");return new WebFormField(line[..equals].Trim(),line[(equals+1)..]);
    }).ToArray();
    public static Task<WebResponse> RequestAsync(string url,bool post,IEnumerable<WebFormField> fields,CancellationToken cancellation=default)=>RequestAsync(Client,url,post,fields,20*1024*1024,cancellation);
    public static async Task<WebResponse> RequestAsync(HttpClient client,string url,bool post,IEnumerable<WebFormField> fields,int limit,CancellationToken cancellation=default)
    {
        if(!Uri.TryCreate(url,UriKind.Absolute,out var uri)||uri.Scheme is not ("http" or "https")||uri.UserInfo.Length>0)throw new ArgumentException("Enter an HTTP/HTTPS URL without embedded credentials.");
        using var request=new HttpRequestMessage(post?HttpMethod.Post:HttpMethod.Get,uri);
        if(post)request.Content=new FormUrlEncodedContent(fields.Select(f=>new KeyValuePair<string,string>(f.Name,f.Value)));
        using var response=await client.SendAsync(request,HttpCompletionOption.ResponseHeadersRead,cancellation).ConfigureAwait(false);
        if(response.Content.Headers.ContentLength>limit)throw new InvalidDataException("Response exceeds size limit.");
        await using var input=await response.Content.ReadAsStreamAsync(cancellation).ConfigureAwait(false);using var output=new MemoryStream();var buffer=new byte[81920];
        while(true){var count=await input.ReadAsync(buffer,cancellation).ConfigureAwait(false);if(count==0)break;if(output.Length+count>limit)throw new InvalidDataException("Response exceeds size limit.");await output.WriteAsync(buffer.AsMemory(0,count),cancellation).ConfigureAwait(false);}
        return new(response.RequestMessage?.RequestUri?.ToString()??url,(int)response.StatusCode,response.Content.Headers.ContentType?.ToString()??"application/octet-stream",output.ToArray());
    }
}
