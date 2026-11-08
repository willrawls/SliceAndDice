# Sandy — Attempt 2 · Blazor Hybrid

The desktop GUI is now **Blazor Hybrid on .NET 10**: Razor components and CSS inside a minimal WPF/BlazorWebView shell. WPF only hosts the window and native file dialogs. No WinForms code is compiled into the application.

## Run

```powershell
dotnet run --project '.\Sandy-Attempt-2\src\Sandy.Desktop' -c Release
```

Requires Windows 10 version 1809 or later, .NET 10 desktop and ASP.NET Core runtimes (included with the development SDK), and Microsoft Edge WebView2 Runtime. The first launch creates a WebView profile under `%LOCALAPPDATA%/SandyAttempt2/WebView2`. The app runs locally; no web server or cloud account is required.

## Structure

| Project | Responsibility |
| --- | --- |
| Sandy.Core | Library model, XML persistence, template interpreter |
| Sandy.Blazor | Reusable Razor workspace/dialogs and editing state; native operations behind IDesktopServices |
| Sandy.Desktop | WPF window hosting BlazorWebView, native file dialogs, read-only MDB import |
| Sandy.VisualStudio | Native Remote UI tool window and checked editor insertion; launches the bundled Blazor workspace |
| Sandy.Tests | Core, draft/persistence, and Razor rendering regressions |

The workspace has category/template browsing and filtering, four code areas, variables, prompted preview, scoped find/replace, lock/favorite/undeletable options, category rename/delete and template duplication. Per-template drafts survive navigation. Update commits the current draft in memory; Save copy validates and saves every draft atomically. Source files cannot be overwritten by Save copy. Ctrl+S opens Save copy; Tab inserts indentation in the code editor. Editor context lets you load a source buffer and selection for the same virtual editor used by the extension. Preview displays document and external-action plans; Apply confirms them and journals file changes.

All **21 original forms** are represented: the main form becomes the workspace, and the other 20 are available under Tools and dialogs. Several auxiliary dialogs still have unfinished legacy actions, explicitly reported when invoked. This GUI migration does not establish full interpreter or legacy-handler parity. See [port status](docs/PORT-STATUS.md).

The bundled library was exported directly from `Sandy/SliceAndDice.mdb`: **25 categories, 419 templates**. Original fields and line endings are retained. Read-only MDB import supports Access 97 through an isolated 32-bit DAO 3.6 PowerShell helper when DAO is installed. Examples.xml contains separately authored C# examples.

## Verify

```powershell
& '.\Sandy-Attempt-2\Verify.ps1'
```

This builds the solution and VSIX, runs the regression suite, exercises the actual WebView2 UI (DOM binding, draft switching, save/reload and all 20 dialogs), saves screenshots under the ignored `artifacts/blazor-verification` folder, and checks the packaged desktop payload including web assets and WebView2Loader. Rendering a dialog does not prove all its legacy actions are implemented.

## Visual Studio

The VSIX is at `src/Sandy.VisualStudio/bin/Release/net10.0-windows10.0.17763.0/Sandy.VisualStudio.vsix`. Extension identity remains `WilliamRawls.Sandy.Attempt2`.

```powershell
& '.\Sandy-Attempt-2\Deploy-Experimental.ps1' -Launch
```

The extension requires Visual Studio with a .NET 10 extensibility host. Experimental profile: SandyAttempt2. For a normal-profile installation, close Visual Studio and run `Install-VisualStudio.ps1`; `-Launch` opens the installed IDE afterward. Choose Extensions > Sandy — Attempt 2, then Open Blazor workspace for editing. Save an XML copy and reload that path in the extension before insertion. The VSIX launches the separate desktop workspace; it does not embed Blazor inside the IDE tool window.

The install script uses native Visual Studio MSBuild deployment because this SDK's multi-architecture Extensibility manifest is rejected by the native VSIXInstaller on this machine. If Sandy is initially listed under Disabled by user, enable it in Extensions > Manage Extensions > Installed, then restart Visual Studio. The regular profile was enabled during installation here.

Insertion supports cursor, declarations, bottom-of-type, generated files and virtual editor commands. Open source buffers use checked editor snapshots; file changes have preflight and byte backups under `.sandy-history`. Editor changes are one edit batch. Undo does not reverse generated files or OS actions. Live IDE checks passed for cursor/multi-area insertion, generated files, cancellation, and exact single-Undo restoration in the experimental profile. Read [port status](docs/PORT-STATUS.md) for remaining compatibility and live-test boundaries.

## Sources

Primary behavioral source: `../Sandy`. The previous WinForms implementation remains in `src/Sandy.Desktop/Forms` as an excluded reference; only its generated catalog is embedded by Sandy.Blazor. `tools/inventory.py` regenerates this catalog and the source SHA-256 inventory without modifying the original sources.

Matching archived classes may supplement missing Sandy files with the user's authorization. [Archive sources](docs/ARCHIVE-SOURCES.md) records the evidence and version boundaries; `docs/archive-matches.json` lists candidates. The deterministic interpreter/expression parser adapts algorithms from `../SandyCSharp10`, replacing its MetX dependencies with local helpers and extending the host contract. That previous attempt remains unchanged and is not a project dependency.

`tools/Sandy.Audit` probes the bundled templates against a synthetic virtual editor without committing effects. Run it with a library XML path and a report JSON path. Its reports identify compatibility/fixture gaps; they do not establish output parity.
