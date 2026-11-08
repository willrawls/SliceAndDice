using System.Text;
namespace Sandy.Core;

/// <summary>Preserves BOMs and the Windows-1252 fallback used by legacy VB source files.</summary>
public static class TextFiles
{
    static TextFiles() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    private static Encoding Detect(byte[] bytes)
    {
        if(bytes is [0xFF,0xFE,0,0,..])return new UTF32Encoding(false,true,true);
        if(bytes is [0,0,0xFE,0xFF,..])return new UTF32Encoding(true,true,true);
        if(bytes is [0xFF,0xFE,..])return new UnicodeEncoding(false,true,true);
        if(bytes is [0xFE,0xFF,..])return new UnicodeEncoding(true,true,true);
        if(bytes is [0xEF,0xBB,0xBF,..])return new UTF8Encoding(true,true);
        try{new UTF8Encoding(false,true).GetString(bytes);return new UTF8Encoding(false,true);}
        catch(DecoderFallbackException){return Encoding.GetEncoding(1252,EncoderFallback.ExceptionFallback,DecoderFallback.ExceptionFallback);}
    }
    public static string Read(string path) => Decode(File.ReadAllBytes(path));
    public static string Decode(byte[] bytes)
    {
        var encoding=Detect(bytes);var prefix=encoding.GetPreamble();var skip=prefix.Length>0&&bytes.AsSpan().StartsWith(prefix)?prefix.Length:0;
        return encoding.GetString(bytes,skip,bytes.Length-skip);
    }
    public static byte[] Encode(string text,byte[]? original=null)
    {
        var encoding=original==null?new UTF8Encoding(false,true):Detect(original);
        return [..encoding.GetPreamble(),..encoding.GetBytes(text)];
    }
}
