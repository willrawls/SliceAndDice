using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

namespace Sandy.Core;
internal static class SourceNavigation
{
    public static (int Top,int Bottom)? TypeBoundaries(string text,int position)
    {
        var cs=CSharpSyntaxTree.ParseText(text).GetRoot().DescendantNodes().OfType<TypeDeclarationSyntax>().Where(t=>!t.OpenBraceToken.IsMissing&&!t.CloseBraceToken.IsMissing).ToArray();
        var type=cs.Where(t=>t.SpanStart<=position&&t.Span.End>=position).OrderBy(t=>t.Span.Length).FirstOrDefault()??cs.FirstOrDefault();
        if(type!=null)return(type.OpenBraceToken.Span.End,type.CloseBraceToken.SpanStart);
        var vb=VisualBasicSyntaxTree.ParseText(text).GetRoot().DescendantNodes().OfType<TypeBlockSyntax>().ToArray();
        var block=vb.Where(t=>t.SpanStart<=position&&t.Span.End>=position).OrderBy(t=>t.Span.Length).FirstOrDefault()??vb.FirstOrDefault();
        return block==null?null:(block.BlockStatement.Span.End,block.EndBlockStatement.SpanStart);
    }
    public static (int Start,int End,string Name)? Member(string text,int position,string name)
    {
        var methods=CSharpSyntaxTree.ParseText(text).GetRoot().DescendantNodes().OfType<MethodDeclarationSyntax>();
        var method=methods.FirstOrDefault(m=>name.Length==0?m.FullSpan.Contains(position):m.Identifier.ValueText.Equals(name,StringComparison.OrdinalIgnoreCase));
        if(method!=null)return(method.SpanStart,method.Span.End,method.Identifier.ValueText);
        return null;
    }
}
