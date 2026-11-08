namespace Sandy.Core;
public static class DeferredActions
{
    public static IReadOnlyList<ExternalEffect> Apply(IEnumerable<ExternalEffect> effects,Func<ExternalEffect,string?> execute)
    {
        var bindings=new Dictionary<string,string>(StringComparer.Ordinal);var applied=new List<ExternalEffect>();
        string Resolve(string text){foreach(var binding in bindings)text=text.Replace(binding.Key,binding.Value,StringComparison.Ordinal);return text;}
        foreach(var effect in effects)
        {
            var resolved=effect with{Target=Resolve(effect.Target),Value=Resolve(effect.Value)};var result=execute(resolved);applied.Add(resolved);
            if(effect.ResultToken!=null)bindings[effect.ResultToken]=result??throw new InvalidOperationException("Action did not return a process ID.");
        }
        return applied;
    }
}
