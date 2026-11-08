# Port status — 0.4.0

Sandy Attempt 2 is a .NET 10 Blazor Hybrid recreation with a Visual Studio extension. All 21 original forms have modern representations. This release fills the major outstanding feature areas; exact output parity for every legacy template, COM integration and historical service remains unproven. The original source remains unchanged.

## Implemented

- Library XML preserves unknown fields, auxiliary records and code line endings. Per-template drafts, revisions, locks, favorites, undeletable flags, category/template operations, scoped search, import/export, history, drag/drop and template context menus work. Save copy preserves the source library. Versioned JSON delta patches validate baseline/result hashes and apply atomically; this is a modern format rather than the original binary patch format.
- GotoSandyTemplate and SetCodeArea/Text edits use an isolated working library. Preview/cancel leaves the original untouched. Apply checks concurrent library edits and conflicting drafts. The extension writes a new XML working copy through the file journal before accepting the revised selection.
- Read-only Access schema/library import includes the isolated 32-bit DAO helper for Access 97. The bundled library contains 25 categories and 419 templates. ODBC schema import is available.
- Database design supports table/field edits, relationships, favorites, marked/all/branch generation, key/AutoNumber/date rules, cascades and design JSON. Output includes typed models, collections, parameterized async CRUD repositories and schema SQL for Access, SQL Server or SQLite. Source-derived collection/member/property/wrapper template dispatch is also available with shared variables and virtual output files. This edits a working design; it does not migrate the source database.
- Header tools inspect declarations, constants, class members and local includes. They generate supported enums, structs/unions, arrays and constants with packing and Boolean marshalling. The original Generate handler was a placeholder; this is working replacement functionality.
- Interpreter supports bounded control flow, expressions, nested includes/SoftCode, prompts and choices, string/token operations, provider contracts and four code destinations. Virtual editor operations cover source/procedure navigation, modules/projects/references, VB6 controls and arrays, nested font properties and declaration-preserving edits. C# and VB.NET placement uses Roslyn.
- File streams preserve blank lines and EOF; planned writes are readable; directory matching advances independently. File/settings/clipboard/process/activation/keyboard actions are planned for explicit Apply. Process IDs bind between deferred actions when executed. File commits use workspace/stale-state checks, byte backups and rollback on editor failure; supported source encodings are retained.
- Web tools provide explicit GET/POST, encoded duplicate POST fields, alternate URLs, cancellable bounded downloads, response inspection and persistent URL history. Responses display as text. Auxiliary preferences omit password fields.
- Native desktop tray, favorites, configurable global shortcut, close-to-tray and Exit are implemented. The VSIX adds shortcuts and editor context insertion. VB5 INI editing previews changes and applies through the file journal; it does not register COM DLLs.
- Other tools provide library data browsing/search, field selection, bounded file search, choice selection, keyboard diagnostics, command help and original EN* obfuscation. About/registration describes the modern app without claiming the old activation server works.

## Verification

Verify.ps1 passes a Release solution/VSIX build with zero warnings/errors and 95 regression tests. Tests include generated-code compilation for all three SQL dialects, real SQLite CRUD/cascade deletion, self-edit isolation, stale-state rejection, patch integrity/protections, designer arrays/font edits/renames, POST encoding, preferences, EOF, process-action binding and legacy generator dispatch.

The real WebView2 run drives the workspace and all 20 auxiliary forms, repeated prompts, choices, library mutation Apply, database model/repository generation, drafts, save/reload, obfuscation and help. It registers/unregisters a native tray icon and hotkey. Screenshots/results are under ignored artifacts/blazor-verification. Registration does not prove every tray interaction or hotkey keystroke.

Version 0.3 live tests in experimental and regular Visual Studio profiles verified cursor insertion and exact saved-byte restoration with one Undo. Experimental tests also verified multi-area placement, generated files, journal and cancellation. Those receipts remain in verification-0.3.0.json; they are not automatically attributed to this release. Version 0.4 results are recorded separately. Install-VisualStudio.ps1 uses native MSBuild SDK deployment because this machine's VSIXInstaller rejects the SDK's multi-architecture manifest.

Version 0.4 is installed in the regular Visual Studio Community 2026 profile. Its opening shortcut, editor context-menu command, cursor insertion and exact saved-byte single-Undo restoration passed live. Library self-editing created a new XML working copy, changed the selected template and recorded a file journal while leaving the source buffer unchanged. Installed extension/core/desktop/Blazor hashes match the verified build. See verification-0.4.0.json.

## Compatibility and live-test boundaries

- Exact results for all 419 templates remain unproven. Some require missing files, particular VB6 projects, historical providers or contain source defects. The synthetic corpus probe commits no effects and is not a parity test. The deterministic expression evaluator does not execute arbitrary VBScript.
- ODBC against a configured DSN, repositories against Access/SQL Server, and their DDL execution remain untested. SQLite CRUD is tested. Repositories require one primary key; database migration and optimistic concurrency are not supplied.
- Header handling is not a full C/C++ compiler/preprocessor: conditional compilation, complete typedef/API resolution and C++ ABI generation remain limited. Review diagnostics.
- VB6 .frm operations edit source definitions; they do not create native VB6 designers or Blazor components. COM references require suitable .NET interop assemblies. Obsolete COM add-ins are not loaded by the managed provider contract.
- IDE close/hide maps to native document close with save prompting. The standalone virtual editor has no IDE views to close. The original IDE-event dispatcher in frmMain.HandleIDEEvents is entirely commented out; no new automatic event-macro engine is invented here. Original Pause/WaitForFile handlers were no-ops and retain that behavior.
- Deferred process IDs can feed later actions but cannot generate source during preview. Clipboard, SendKeys, process and registry runners build but have not been exercised against the user's applications/settings. OS actions and generated files are not undone by Visual Studio editor Undo.
- Purchasing/registration servers and historical site-specific workflows are not restored. General POST/history and VB5 INI editing are implemented replacements with the above boundaries.

The source inventory records 163 original files and 24 unresolved project references. Archive candidates/version boundaries are documented in ARCHIVE-SOURCES.md; matching names alone do not prove matching behavior.
