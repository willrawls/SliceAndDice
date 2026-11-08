using System.IO;
using System.Text.Json;
using Microsoft.Web.WebView2.Core;
using Sandy.Blazor;
namespace Sandy.Desktop;
internal static class HybridSmoke
{
    public static async Task RunAsync(CoreWebView2 webview, Workspace state, string directory,Action? verifyNativeShell=null)
    {
        await webview.CallDevToolsProtocolMethodAsync("Runtime.enable", "{}");
        webview.GetDevToolsProtocolEventReceiver("Runtime.exceptionThrown").DevToolsProtocolEventReceived += (_, e) => File.AppendAllText(Path.Combine(directory, "javascript-errors.txt"), e.ParameterObjectAsJson + Environment.NewLine);
        async Task WaitFor(string expression)
        {
            for (var i = 0; i < 150; i++)
            {
                if (await webview.ExecuteScriptAsync(expression) == "true") return;
                await Task.Delay(100);
            }
            File.WriteAllText(Path.Combine(directory, "failed-dom.json"), await webview.ExecuteScriptAsync("document.documentElement.outerHTML"));
            throw new InvalidOperationException("Timed out: " + expression + "; URL=" + webview.Source);
        }
        await WaitFor("document.querySelector('[data-ready=true]') !== null");
        await WaitFor("document.querySelectorAll('.template').length === 419");
        await using (var stream = File.Create(Path.Combine(directory, "workspace.png"))) await webview.CapturePreviewAsync(CoreWebView2CapturePreviewImageFormat.Png, stream);
        await webview.ExecuteScriptAsync("var editor=document.getElementById('code-editor'); editor.value='Blazor smoke draft'; editor.dispatchEvent(new Event('input',{bubbles:true}));");
        await Task.Delay(200);
        if (state.Current?.Code[Sandy.Core.CodeArea.Cursor] != "Blazor smoke draft") throw new InvalidOperationException("DOM input did not update the Blazor draft.");
        await webview.ExecuteScriptAsync("document.querySelectorAll('.template')[1].click()"); await Task.Delay(150);
        await webview.ExecuteScriptAsync("document.querySelectorAll('.template')[0].click()");
        await WaitFor("document.getElementById('code-editor').value === 'Blazor smoke draft'");
        // Exercise two consecutive asynchronous prompts through the actual Razor dispatcher.
        state.Current!.Code[Sandy.Core.CodeArea.Cursor]="%%First::one%%|%%Second::two%%";
        foreach(var area in new[]{Sandy.Core.CodeArea.Declarations,Sandy.Core.CodeArea.Bottom,Sandy.Core.CodeArea.File})state.Current.Code[area]="";
        await webview.ExecuteScriptAsync("document.getElementById('preview-button').click()");
        await WaitFor("document.querySelector('.prompt-dialog h2')?.textContent === 'First'");
        await webview.ExecuteScriptAsync("document.getElementById('prompt-accept').click()");
        await WaitFor("document.querySelector('.prompt-dialog h2')?.textContent === 'Second'");
        await webview.ExecuteScriptAsync("document.getElementById('prompt-accept').click()");
        await WaitFor("document.getElementById('output-preview')?.value.includes('one|two') === true");
        state.Current.Code[Sandy.Core.CodeArea.Cursor]="%%Choose::One;Two%%";
        await webview.ExecuteScriptAsync("document.getElementById('preview-button').click()");
        await WaitFor("document.querySelectorAll('select#prompt-input option').length === 2");
        await webview.ExecuteScriptAsync("var choice=document.getElementById('prompt-input');choice.value='Two';choice.dispatchEvent(new Event('change',{bubbles:true}));document.getElementById('prompt-accept').click()");
        await WaitFor("document.getElementById('output-preview')?.value.includes('Two') === true");
        state.Current.Code[Sandy.Core.CodeArea.Cursor]="~~GotoSandyTemplate Smoke - Created\r\n~~SetCodeArea Cursor=Created by macro";
        await webview.ExecuteScriptAsync("document.getElementById('preview-button').click()");
        await WaitFor("document.getElementById('output-preview')?.value.includes('[Library template] Smoke - Created') === true");
        await webview.ExecuteScriptAsync("[...document.querySelectorAll('button')].find(b=>b.textContent==='Apply this plan').click()");
        await WaitFor("document.querySelector('.prompt-dialog') !== null");
        await webview.ExecuteScriptAsync("document.getElementById('prompt-accept').click()");
        await WaitFor("document.getElementById('template-name').value === 'Smoke - Created'");
        var originalPath=state.SourcePath!;state.Load(Sandy.Core.Library.Load(originalPath),originalPath);
        await WaitFor("document.querySelectorAll('.template').length === 419");
        state.Current.Code[Sandy.Core.CodeArea.Cursor]="Blazor smoke draft";
        var table=state.DatabaseDesign.AddTable("Products");table.Fields.Add(new(){Name="Name",DaoType=10,Length=100,Nullable=false});state.DatabaseDesign.ApplyRules(true,true,true,true);
        var rendered = new List<string>();
        foreach (var form in FormCatalog.All.Where(f => f.ClassName != "frmMain"))
        {
            await webview.ExecuteScriptAsync($"var select=document.getElementById('form-selector');select.value={JsonSerializer.Serialize(form.ClassName)};select.dispatchEvent(new Event('change',{{bubbles:true}}));");
            await WaitFor("document.querySelector('.tool-dialog') !== null");
            if(form.ClassName=="TestCrypt")
            {
                await webview.ExecuteScriptAsync("var input=document.getElementById('txtIn');input.value='ABC';input.dispatchEvent(new Event('input',{bubbles:true}));[...document.querySelectorAll('.tool-dialog button')].find(b=>b.textContent.includes('Encrypt')).click()");
                await WaitFor("document.getElementById('txtOut').value === 'EN* DBEBBC'");
            }
            if(form.ClassName=="frmCommandHelp")await WaitFor("document.querySelectorAll('.command-help option').length > 100");
            if(form.ClassName=="frmDBClassGen")
            {
                await webview.ExecuteScriptAsync("[...document.querySelectorAll('.tool-dialog button')].find(b=>b.textContent==='Generate model, collection, repository and schema').click()");
                await WaitFor("document.querySelector('.tool-dialog textarea.generated-code')?.value.includes('class Products') === true");
                await webview.ExecuteScriptAsync("var generated=[...document.querySelectorAll('.tool-dialog select')].find(s=>[...s.options].some(o=>o.value==='ProductsRepository.cs'));generated.value='ProductsRepository.cs';generated.dispatchEvent(new Event('change',{bubbles:true}))");
                await WaitFor("document.querySelector('.tool-dialog textarea.generated-code')?.value.includes('InsertAsync') === true");
            }
            await using (var stream = File.Create(Path.Combine(directory, form.ClassName + ".png"))) await webview.CapturePreviewAsync(CoreWebView2CapturePreviewImageFormat.Png, stream);
            rendered.Add(form.ClassName);
            await webview.ExecuteScriptAsync("document.querySelector('[aria-label=\"Close dialog\"]').click()");
            await WaitFor("document.querySelector('.tool-dialog') === null");
        }
        state.Save(Path.Combine(directory, "smoke-working-copy.xml"));
        verifyNativeShell?.Invoke();
        File.WriteAllText(Path.Combine(directory, "results.json"), JsonSerializer.Serialize(new { Workbench = true, Forms = rendered, DraftSwitching = true, SaveReload = true, RepeatedPrompts = true, ChoicePrompt=true,LibraryMutation=true,DatabaseGeneration=true,NativeShellRegistration=verifyNativeShell!=null,LegacyObfuscation = true, CommandReference = true }, new JsonSerializerOptions { WriteIndented = true }));
    }
}
