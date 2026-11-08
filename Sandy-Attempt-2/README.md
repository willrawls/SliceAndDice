# Sandy — Attempt 2

Fresh C# / .NET 10 recreation based on **`../Sandy`**, with a Windows Forms application and a Visual Studio extension. Attempt 1 is preserved in `../SandyCSharp10`; this solution does not reference it.

**This is a working partial port, not full legacy behavior parity.** All 21 form layouts are represented, including DataView's `Form1`, the standalone shell and the add-in tool. Unported actions report their status; unsupported SoftCommands stop evaluation with context. See [port status](docs/PORT-STATUS.md).

## Run

```powershell
dotnet run --project '.\Sandy-Attempt-2\src\Sandy.Desktop' -c Release
```

The desktop loads the bundled library exported directly from `Sandy/SliceAndDice.mdb`: **25 categories, 419 templates**. The MDB remains unchanged. `Libraries/Examples.xml` contains two separately authored C# examples.

The toolbar opens/imports libraries, creates categories and templates, updates templates, saves an XML copy, previews output, searches/replaces text, and opens every form. Drafts survive template switching. Preview uses those drafts; Save copy applies all drafts. The original library path cannot be overwritten. Locked templates reject code edits; the template options tab can unlock the working copy.

All four code areas are retained. Preview never writes generated files or applies registry, clipboard or IDE effects. MDB reads use DAO; Access 97 databases fall back to 32-bit DAO 3.6 in an isolated Windows PowerShell process. This requires DAO 3.6 to be installed; no COM components are registered by the app. The fallback invokes only the bundled read-only export script, with execution policy set for that child process only.

## Build and check

```powershell
& '.\Sandy-Attempt-2\Verify.ps1'
```

Open `Sandy.slnx` in Visual Studio, or build with `dotnet build .\Sandy-Attempt-2\Sandy.slnx -c Release -m:1`. `Verify.ps1` builds, runs regression tests, renders all forms, and checks that the VSIX contains the desktop payload and original library export. Form snapshots go to the ignored `artifacts/form-previews` folder. Snapshot creation is not full interaction or visual verification.

## Visual Studio

The VSIX is at `src/Sandy.VisualStudio/bin/Release/net10.0-windows8.0/Sandy.VisualStudio.vsix`. It has a distinct extension identity, `WilliamRawls.Sandy.Attempt2`, and targets a .NET 10 extensibility host. Experimental deployment uses the `SandyAttempt2` profile:

```powershell
& '.\Sandy-Attempt-2\Deploy-Experimental.ps1' -Launch
```

Deployment and live IDE behavior have **not** been verified for this attempt. Build/package verification does not establish host compatibility.

Choose **Extensions > Sandy — Attempt 2**, load a library, select a template and preview. **Open desktop forms** launches the bundled desktop app with the current library. After saving edits to a working copy, enter that XML path and reload in the extension. Select the target editor and choose **Extensions > Insert Sandy Template — Attempt 2**. Missing values prompt with defaults; one version-checked edit replaces the selection after confirmation. Insertion currently accepts cursor-only output; templates with other code areas are rejected as a whole.

## Source fidelity

`tools/inventory.py` reads only `Sandy`, captures every form/control/property and source handler name, decodes FRX strings/lists/pictures, and emits SHA-256 records for every source file. Running it regenerates `Forms.g.cs`, `catalog.json` and `docs/source-inventory.json`.

The source's project files contain 24 unresolved references. No missing behavior has been silently substituted from another legacy snapshot. The `Sandy` folder is the behavioral reference for this attempt.
