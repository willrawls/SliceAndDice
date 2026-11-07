# Compatibility and remaining implementation

This version is a usable starting port, not a claim that every historical template executes successfully. Import preservation and interpreter compatibility are separate measurements.

## Implemented

- `netstandard2.0`, `netstandard2.1` and `net10.0` engine binaries, with no Visual Studio/COM dependency. Each uses the corresponding local MetX.Standard.Strings build for token and substring operations.
- Original category/template fields, all four code areas, unknown metadata retention, locked-template edit protection, revision updates and copy-based XML persistence.
- Access XML libraries and explicitly profiled 18-column CSV with quoted multiline fields. Duplicate identities and mismatched relationships fail import.
- Case-insensitive SoftVariables, prompts/defaults, indirect `@Variable` defaults, built-in EOL values, per-run saved variables, blank-value prompting and IgnoreBlanks/WatchBlanks.
- Preprocessor includes with explicit sections, recursive-include detection, script/mixed/comment regions and legacy command reference labels.
- Set/Let, Eval, Append, Increment/Decrement with optional amounts, StrCopy, Trim variants, Replace with ToFind/ReplaceWith, String, SaveVars/RestoreVars, Block, nested If/Else/EndIf and Coda/LoopWhile/LoopUntil blocks.
- Output plans for editor, append/overwrite files, clipboard, messages and SoftVariables. Destination changes flush the previous buffer in order. No external effects occur during evaluation.
- Inline case/length/trim/substring/padding/replacement/character/token operations and Now without a format; common aliases are implemented in the interpreter.
- Deterministic expression parsing for numeric arithmetic, comparison, booleans and string concatenation. It does not evaluate arbitrary code.
- Managed Sandal provider interface for commands and inline functions.
- Bounded execution, expansion, include depth and output; cancellation propagates without returning partial output plans.
- VSIX library browser/search, four-area code drafts, explicit updates, preview, copy export, interactive insertion prompts and cursor-only editor insertion.
- Serialization metadata explicitly declares the net10.0 host. MessagePack is pinned to a patched 2.x version; the SDK supplies shared host assemblies at runtime.

## Known limitations

| Area | Current boundary / next work |
| --- | --- |
| IDE application | Cursor-only output is submitted in one version-checked edit. Declarations, bottom, file and other destinations cause the whole insert operation to be rejected before editor mutation. |
| Placement | Add Roslyn C#/VB.NET adapters for declarations, enclosing member/type and member insertion. Text-only insertion remains language-independent. |
| Navigation/transformation | GotoProject/Module/Proc, line manipulation, selection commands, project-wide operations and automatic IDE variables remain to be implemented. |
| Designer | AddControl, SetProperty, control iteration and event generation require designer-specific adapters. VB6 form/control types have no general direct equivalence. |
| Attributes | VB6 ProcAttr values require deliberate mappings to .NET attributes/interop metadata; no silent conversion. |
| Expressions | A deterministic subset of VB-like expressions, not a complete VB6 expression evaluator. VB formatting, coercion, locale, overflow and edge cases require characterization. |
| Inline parsing | Comma-separated legacy arguments; full quoted/nested inline grammar and delimiter-override directives are not yet implemented. Unknown inline names retain the original prompt-variable interpretation. |
| Saved variables | SaveVars/RestoreVars currently share values within one execution. Persistence across separate runs is not implemented. |
| Files/clipboard/processes | Destinations are previewed only. Add transactional file commit/backups, explicit host capabilities, clipboard and process adapters. Keyboard automation is deferred to a companion host. |
| Generated SoftCode | GotoSoftCode/evaluation of newly generated commands is not implemented. Unsupported commands fail with template/area/expanded-line diagnostics. |
| CSV/MDB | 22-column layouts require schema verification. Direct MDB import and newer CSV profiles are pending. Source XML/CSV files are never rewritten. |
| UI library management | Code editing and copy export exist. Creating/deleting/reorganizing categories/templates, favorite management and command help completion remain pending. Drafts are reset when another library loads. |
| Plug-ins | Managed providers are programmatically registered. No arbitrary DLL discovery UI or VB6 COM binary compatibility. |
| Database generation | SQL Server/ODBC schema providers and schema-driven templates remain to be ported. |
| Timers/events/global hotkeys | Not implemented. A future companion host can share Core. |
| Error locations | Errors identify template, area and line in expanded source. Include-origin source maps remain pending. |

## Next milestones

1. Keep experimental-instance validation as a release gate. Browser, editing, save/reload, prompt defaults, insertion and undo have passed live checks; cancellation and concurrent-edit mutation rejection still need live coverage. See [the verification record](VISUAL-STUDIO-VERIFICATION.md).
2. Roslyn host adapter and simulated workspace for multi-area and multi-document edit plans; include selection context and common navigation commands.
3. Run a representative historical-template corpus with fixed inputs/host fixtures and expected results; generate a per-command and per-template coverage report. Existing tests verify imports and representative semantics, not complete corpus execution.
4. Complete template/category management, command metadata/help and favorites. Add library revisions and import conflict handling before supporting merges.
5. Transactional file output, clipboard host support, managed Sandal discovery and database generation.
6. Designer-specific feasibility spikes and explicit mappings, followed by advanced automation/standalone companion support.

## Manual experimental-instance test

Use a Visual Studio installation that provides the .NET 10 extension host. Keep tests in a dedicated experimental profile.

1. Close the test IDE, then run `Deploy-Experimental.ps1 -Launch` to rebuild, deploy and launch the `SandyCSharp10` profile with the native SDK tooling.
2. Open Sandy through Extensions. Load the included C# library and confirm three templates appear.
3. Preview Numbered comments: expect three numbered comment lines. Provide overrides and verify that they affect output.
4. Edit a template, update it, save a library copy, reload the copy and verify code/revision retention. Confirm locked historical templates reject updates.
5. Select Property, return to a disposable C# document and invoke Insert Sandy Template. Verify prompt defaults, preview, exact selected-range replacement and one Undo restoring the document.
6. Cancel a prompt and final confirmation; verify no editor change. Change the document while the confirmation is open and verify mutation rejection is reported.
7. Load the historical 550-template library. Confirm browse/import succeeds. A designer template should report its unsupported command, rather than inserting command text.
8. Preview a multi-area template and verify insertion rejects it without partial changes.

Do not equate a successful build or VSIX installation with these UI checks passing.

## Validation from the initial implementation session

- Release solution build: successful, zero warnings/errors, both Core targets and VSIX produced.
- Regression suite: run against both Core targets. The suite verifies the loaded assembly's TargetFrameworkAttribute so the Standard run cannot silently use the net10.0 assembly.
- Historical imports: seven XML libraries roundtrip every category/template field; one original VB6 Public Function template is executed with fixed inputs.
- Package contents: inspected the VSIX manifest, net10.0 runtime requirement, Core assembly and bundled sample library.
- Initial Visual Studio UI attempt: not verified at that point. The quiet installer exited with code **1015**. Later investigation found the diagnostic log in `%TEMP%`: `RequiresInstallerException` with a multiple-target error. The verified development solution is native SDK deployment through `Deploy-Experimental.ps1`; see the updated [live verification record](VISUAL-STUDIO-VERIFICATION.md).

MetX integration extends verification to all three Core targets. Regression checks assert the actual framework metadata of both Core and MetX, cover token/substring edge cases, compare copied DLL hashes with the local inputs and inspect the packaged .NET 10 dependency. See [integration details](METX-INTEGRATION.md).

## Live Visual Studio verification, 2026-11-07

The extension was deployed into the isolated VS 2026 profile and exercised through its actual UI. An edited Property template survived copy export and reload. Its default prompts and insertion preview ran inside the IDE, the generated text replaced the selected document range, and one Ctrl+Z restored the complete original file. SHA-256 comparison after saving confirmed byte-for-byte equality, including the original mixed line endings. No insertion/undo defect was found in this scenario.

This is a focused smoke test, not proof of every manual case above. [Full steps, expected output and evidence](VISUAL-STUDIO-VERIFICATION.md).
