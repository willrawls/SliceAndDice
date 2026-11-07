# Slice and Dice — Sandy C# 10

A new C# implementation alongside the preserved VB6 sources. Behavioral reference: `Archive/sad2`, version 1.8.198. This is an initial working port, with a buildable VSIX and tested interpreter; it is not full Sandy parity yet.

## Targets

| Project | Target | Purpose |
| --- | --- | --- |
| `Sandy.Core` | `netstandard2.0;netstandard2.1;net10.0` | Library model, XML/CSV import, template editing, interpreter, managed Sandal contract |
| `Sandy.Cli` | `net10.0` | Library listing, command usage audit, conversion, execution preview |
| `Sandy.VisualStudio` | `net10.0-windows8.0` | Out-of-process VisualStudio.Extensibility VSIX and Remote UI tool window |
| `Sandy.Core.Tests` | `net10.0` | Regression suite runnable against either Core assembly |

The .NET 10 extension requires a Visual Studio installation with a .NET 10 extensibility host. The generated SDK manifest has a 17.14 API minimum; this does **not** mean every VS2022 installation can host the extension. There is no .NET Framework dependency in the new projects.

Core uses William's **MetX.Standard.Strings** assemblies from the sibling xlg checkout at `../xlg/MetX/MetX.Standard.Strings/bin/Debug` (relative to the Slice And Dice repository). Each Core target references the matching framework's DLL; the CLI and VSIX receive the .NET 10 dependency. See [MetX integration](docs/METX-INTEGRATION.md) for configuration and compatibility details.

## Build and verify

From the repository root, in PowerShell:

```powershell
& '.\SandyCSharp10\Verify.ps1'
```

Or open `SandyCSharp10.slnx` in Visual Studio. The VSIX is produced at:

```text
src/Sandy.VisualStudio/bin/Release/net10.0-windows8.0/Sandy.VisualStudio.vsix
```

`Verify.ps1` builds all three Core targets, the CLI and VSIX, then runs the engine tests against each Core target using the .NET 10 test runner. It also compares copied MetX assemblies with the supplied files and checks the .NET 10 DLL inside the VSIX. It does not claim to automate Visual Studio UI validation.

## Use the extension

1. Install the VSIX in a compatible experimental Visual Studio instance. Prefer an isolated profile during development.
2. Choose **Extensions > Slice and Dice — Sandy**.
3. Load the included C# library, or enter the full path to a historical `MdbGallery/BrokenDown/<library>/Category.xml` file.
4. Search and select a template. Inspect or edit its cursor, declarations, bottom and file areas. **Update selected template** commits edits in memory and increments the revision when content changes. Locked templates reject updates.
5. Enter SoftVariables as one `Name=Value` per line. **Preview with defaults** evaluates without applying any output. The preview JSON describes destinations and resulting variables.
6. To insert, focus a code editor and choose **Extensions > Insert Sandy Template**. Missing values prompt interactively with their defaults; a final preview confirms insertion. VS receives a single version-checked editor mutation.
7. **Save a library copy** writes committed library contents under `%LOCALAPPDATA%/SandyCSharp10/Libraries` with a unique filename. Source libraries are not overwritten. Editor drafts survive template switching within a loaded library; loading a library resets those drafts. Use Update and Save Copy before loading another library.

The first extension host applies **cursor-only editor output**. Other areas and destinations can be evaluated and inspected, but insertion rejects them as a whole. This avoids silently placing declarations, generated files or clipboard output at the cursor. Language-aware placement, external effects and project-wide editing are next milestones.

## Use the CLI

Commands below run from the repository root:

```powershell
dotnet run --project '.\SandyCSharp10\src\Sandy.Cli' -- list '.\MdbGallery\BrokenDown\SliceAndDice 2000\Category.xml' 'Public Function'

dotnet run --project '.\SandyCSharp10\src\Sandy.Cli' -- preview '.\SandyCSharp10\samples\CSharp.xml' 'CSharp - Property' 'Property Name=Title' 'Property Type=string'

dotnet run --project '.\SandyCSharp10\src\Sandy.Cli' -- audit '.\MdbGallery\BrokenDown\SliceAndDice 2000\Category.xml'

dotnet run --project '.\SandyCSharp10\src\Sandy.Cli' -- import-csv '.\MdbGallery\BrokenDown\SliceAndDice 2000' '.\SandyCSharp10\imported.xml' windows-1252
```

CLI preview emits JSON; it never applies generated output. Import/roundtrip commands require a new output filename. XML is the preferred historical interchange format. Headerless 18-column CSV imports use an explicit schema; 22-column exports are rejected pending a verified schema. CSV encoding is explicit; some old exports require `windows-1252`.

See [compatibility and remaining work](docs/COMPATIBILITY.md) and [the source investigation](docs/SOURCE-REVIEW.md).
