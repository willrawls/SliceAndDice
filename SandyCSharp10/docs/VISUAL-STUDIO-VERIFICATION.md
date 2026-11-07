# Live Visual Studio verification

Recorded on 2026-11-07. Implementation tested: commit `ccf538e` on `codex-attempt-1`.

## Environment and deployment

- Visual Studio Community 2026, version 18.10.12224.181, instance `8d3778d5`.
- Dedicated experimental profile: `SandyCSharp10`.
- .NET 10 out-of-process extension with its matching MetX.Standard.Strings assembly.
- Release build: zero warnings and errors. The automated suite passed 61 tests for each of net10.0, netstandard2.0 and netstandard2.1; packaged MetX hashes matched the supplied inputs.
- Live actions used the actual Visual Studio menus, Sandy Remote UI, editor, prompts, Save and Undo.

The original VSIXInstaller bootstrapper failed with exit 1015. Its delayed diagnostic output was found in the default `%TEMP%/dd_VSIXInstaller_*.log`, rather than the requested log location. It reported `RequiresInstallerException` and: "Multiple targets are currently not supported for this operation; retry with the instance you wish to target on the command line." Retrying with an explicit instance ID also failed. This was an installer-routing failure, not evidence of a missing .NET 10 host.

The supported SDK development deployment succeeded when **rebuild and deployment ran together** with the installed Visual Studio's native `MSBuild.exe`:

```powershell
& '.\SandyCSharp10\Deploy-Experimental.ps1' -Launch
```

The script discovers the instance and invokes Rebuild with `DeployExtension=true`, `DeployTargetInstanceId`, and `VSSDKTargetPlatformRegRootSuffix=SandyCSharp10`. Running only DeployVsixExtensionFiles is insufficient because that target needs the VSIX generated earlier in the same build. `dotnet build` can create the package but cannot perform this SDK deployment.

This verifies the experimental deployment route. It does not claim that double-click installation through the original VSIXInstaller bootstrapper is repaired or tested for end-user distribution.

## Results

| Check | Actual result |
| --- | --- |
| Registration and loading | Extensions > Slice and Dice — Sandy opened the real tool window. The bundled library loaded three templates. |
| Template editing | Added `// Sandy live edit` to CSharp - Property's cursor code through Sandy's editor and invoked Update selected template. |
| Persistence | Save a library copy created a separate XML library. The saved Property template had revision 2 and retained the added comment. Reloading and selecting it again showed the edit. The saved copy was also loaded in the newly launched IDE for the insertion test. |
| Insertion prompts | Insert Sandy Template displayed Property Type with default `string`, then Property Name with default `Name`. Defaults were accepted through the actual prompt buttons. |
| Preview | Final confirmation displayed 54 generated characters and the edited comment/property output. |
| Selected-range replacement | Selected the entire disposable test document and confirmed insertion. The original selection was replaced by exactly the two generated lines below. |
| Saved output | Saving the editor and reading the file confirmed the generated content. |
| Single undo | One Ctrl+Z restored the complete original document. After Save, its SHA-256 was identical to the before-test snapshot. |

Generated output, with CRLF line endings and a final CRLF:

```csharp
// Sandy live edit
public string Name { get; set; }
```

Original document restored by one Undo:

```csharp
// Sandy insertion/undo smoke test
public class SandySmokeTest
{
    // Replace this marker with the selected Sandy template.
}
```

Before insertion and after Undo+Save, the file hash was:

```text
SHA-256 6F3D8370ED7A46601399348D6FE8E325C8F7AD624D16AFE111CADAD53B774038
```

The original fixture contained mixed line endings. Visual Studio's load-time normalization offer was declined, so the matching hash also checks that Undo retained those exact bytes. Fixture/output snapshots are local ignored files under `artifacts/smoke`; the test used no application source documents.

## Scope

No insertion or undo defect was found in this scenario. No engine or extension behavior change was necessary. The follow-up changes make the working deployment route and measured results explicit.

This test does not establish live cancellation handling, concurrent-document mutation rejection, caret-only insertion, multi-document undo, language-aware declaration placement, designer commands, or complete historical-template compatibility. Those remain separate checks or implementation work described in COMPATIBILITY.md.
