namespace Sandy.Core;
public static class WebContent
{
    private static readonly HttpClient Client=new(){Timeout=TimeSpan.FromSeconds(30)};
    public static Task<byte[]> FetchAsync(Uri uri,CancellationToken cancellation)=>FetchAsync(Client,uri,20*1024*1024,cancellation);
    public static async Task<byte[]> FetchAsync(HttpClient client,Uri uri,int limit,CancellationToken cancellation=default)
    {
        if(uri.Scheme is not ("http" or "https"))throw new ArgumentException("Only HTTP and HTTPS URLs are supported.");
        using var response=await client.GetAsync(uri,HttpCompletionOption.ResponseHeadersRead,cancellation).ConfigureAwait(false);response.EnsureSuccessStatusCode();
        if(response.Content.Headers.ContentLength>limit)throw new InvalidDataException("Response exceeds download size limit.");
        using var output=new MemoryStream();await using var input=await response.Content.ReadAsStreamAsync(cancellation).ConfigureAwait(false);var buffer=new byte[81920];
        while(true){var count=await input.ReadAsync(buffer,cancellation).ConfigureAwait(false);if(count==0)break;if(output.Length+count>limit)throw new InvalidDataException("Response exceeds download size limit.");await output.WriteAsync(buffer.AsMemory(0,count),cancellation).ConfigureAwait(false);}
        return output.ToArray();
    }
}
