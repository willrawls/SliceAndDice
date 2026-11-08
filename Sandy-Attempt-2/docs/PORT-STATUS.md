# Port status - 0.3.0

Sandy Attempt 2 is a .NET 10 Blazor Hybrid application with a Visual Studio extension. It is a substantial working port, not yet a complete replacement for every VB6 handler or historical add-in. The source folder remains unchanged.

## Implemented and exercised

- Category/template XML models, unknown fields, auxiliary root records, revision tracking, locks, favorites, undeletable flags and atomic save-copy. Per-template drafts survive navigation. Category rename/delete, duplication and scoped find/replace are available.
- Read-only Access import and schema inspection, including Jet 3/Access 97 through the bundled 32-bit DAO helper. The primary library contains 25 categories and 419 templates. The original database's SystemInfo table currently has no records; auxiliary records are preserved when present.
- Blazor workspace plus all 20 auxiliary form representations. Working tools include scoped library search, library data browsing, field validation, file search, Access/ODBC schema-based C# model generation, basic C header declarations, source-derived command help and the original EN* obfuscation algorithm.
- Interpreter conditionals, bounded loops, arithmetic, nested includes, dynamic command names, nested SoftCode, prompts, named token lists and counters, string operations, parameter comments, multiple output destinations and command-provider contracts. Errors identify the template, area and line.
- Virtual document navigation and editing, procedure operations, class/module/project plans, project references, VB6 form/control definitions, control iteration, property assignment and procedure attributes. C# and VB.NET type placement use Roslyn. VB6 designer definitions remain .frm source operations; they do not create Blazor components.
- File streams, reads of planned writes, file operations, settings, clipboard, process/activation/keyboard plans and bounded HTTP reads/downloads. Evaluation creates plans. Apply requires confirmation before file or OS mutations. Historical Pause/WaitForFile commands were no-ops in sadFile.cls and retain that behavior.
- File commits perform preflight, concurrent-change checks, workspace path checks, persistent byte backups and rollback on file/editor failure. Existing UTF-8 BOMs, UTF-16 and Windows-1252 source files retain their encoding.
- VSIX load/filter/select and prompts; active and other open source buffers are snapshotted and edited through a checked editor batch. Declarations/bottom use type boundaries. Generated files and project changes use the file journal. Desktop launch includes the Blazor payload and assets.

## Verification

The regression suite has 61 passing tests. Verify.ps1 builds the solution and VSIX, runs the suite, drives the real WebView2 workspace and all auxiliary dialogs, checks consecutive prompts, the legacy encryption example and command-reference loading, and checks the VSIX payload. Results and screenshots are under ignored artifacts/.

Live Visual Studio Community 2026 checks in profile SandyAttempt2 verified the tool window, cursor replacement, multi-area placement, generated-file output, the backup journal, cancellation without writes and exact byte restoration with one Undo for editor edits. The 0.3 build repeated cursor insertion and exact Undo after adding open-document snapshots. File commits have their own journal; Visual Studio Undo does not roll back generated files or OS actions.

The final 0.3.0 build was also installed and enabled in the regular Visual Studio profile. Its extension/core DLL hashes match the verified build. Command registration, tool-window load, cursor insertion, saved output and exact single-Undo restoration passed there. The native VSIXInstaller rejected the multi-architecture manifest; Install-VisualStudio.ps1 uses the working native MSBuild SDK deployment path. Machine-local results are summarized in verification-0.3.0.json.

The corpus audit is a synthetic fixture probe, not a parity test. It deliberately supplies dummy values and no real project/files, and commits no effects. See tools/Sandy.Audit for a repeatable report. A successful evaluation is not proof of correct legacy output. Fragment templates, missing dependencies, source defects and unported behavior are reported separately from successful probes.

## Still incomplete

- Full database-generator behavior: table/field mutation, favorites and relation trees, marked branches, cascade/date/key rules, collection/wrapper templates and database-specific persistence generation. Current output is a typed C# model, not the original complete generated object model. ODBC code builds but has not been exercised against a configured DSN.
- Full C/C++ preprocessing, includes, packing, nested declarations, typedef resolution and API declarations. Unsupported declarations produce diagnostics.
- Self-modifying library commands (GotoSandyTemplate and SetCodeAreaText), complete legacy expression syntax, every command-set-specific behavior and exact output parity for all 419 templates. The expression evaluator is deterministic and does not execute arbitrary VBScript.
- Native IDE close/hide/event behavior, global hotkeys, system tray, context menus, drag/drop, event macros, delta patches, history navigation and persistent auxiliary-dialog state. Several generic dialog actions still report their migration status.
- VB6 control arrays and nested designer property objects need more fidelity testing. COM type-library references require a .NET interop assembly; the port does not automatically recreate obsolete COM references.
- Registration/purchasing, old website-specific POST/history automation and VB5/6 COM add-in installation are not working modern integrations. Their layouts are retained; obsolete external services are not represented as functioning.
- Process IDs are not available while planning, and some action-dependent templates require further host work. Clipboard, SendKeys, process and registry runners compile but were not exercised against the user's applications or settings in live testing. HTTP behavior is tested with a deterministic HTTP fixture, not historical sites.

The source-inventory.json file records 163 original files and 24 unresolved VB project references. Authorized archive candidates remain separately documented in ARCHIVE-SOURCES.md. A matching filename alone does not establish matching behavior.
