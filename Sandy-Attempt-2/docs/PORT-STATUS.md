# Port status

## Implemented

- Fresh .NET 10 library model and Category.xml reader/writer preserving all imported field names and values, including unrecognized fields and CRLF content.
- Read-only MDB import, including the original Jet 3/Access 97 `Sandy/SliceAndDice.mdb` through a 32-bit DAO 3.6 helper. Source-derived XML is bundled for machines without DAO.
- Desktop category/template browsing and filtering; create category/template; delete template with lock/undeletable checks; four code editors; per-template drafts; explicit update; revision increments for changed code; save-copy; current-editor find/replace; preview.
- `%%Name::Default%%`, variable-backed defaults, prompt cancellation, case-insensitive variables, inline string operations; SET/SETVALUE, STRCOPY, APPEND, INC/DEC, TRIM/LTRIM/RTRIM, SAVEVARS/RESTOREVARS, IGNOREBLANKS, INCLUDE and comments. Output and recursion limits; cancellation; separate output areas; no implicit external effects.
- All 21 form definitions, named C# classes, original menus, nested control trees, arrays, bounds, captions, initial lists and FRX picture references. Native controls replace third-party OCXs.
- Field dialogs, choice selection, file search, URL-entry data, login acceptance, database schema inspection, basic C/C++ header type discovery, command syntax help and standalone-shell launch actions.
- VSIX tool window, load/filter/select, variable entry, preview, prompted cursor insertion as one checked editor mutation; bundled desktop launch.

## Remaining behavior

- Most legacy IDE mutation commands: project/module/form navigation, designer/control creation and property assignment, procedure/member manipulation, references, event macros, hotkeys, tray integration, drag/drop and context menus.
- Interpreter conditionals, arithmetic expressions, Coda loops, command providers, registry/file/clipboard/SendKeys operations and the rest of `NewCommands.cls`, `sadFile.cls`, `sadRegister.cls` and the command-set classes. Unsupported commands fail explicitly.
- Language-aware declarations/bottom placement and generated file application in the VSIX. Current insertion is cursor-only.
- Full database class generation, ODBC connections and data grids, C/C++ header conversion, web POST automation/history, command-help navigation, persistence of dialog state and the cryptography test behavior.
- Registration/purchasing, legacy VB5/6 COM add-in installation and obsolete external website integrations. These forms are represented; their external workflows are not ported.
- `SystemInfo` and other MDB tables outside Category/Template are not imported into the working library model. MDB originals remain authoritative.

Unimplemented button/menu actions show a port-status message. Several auxiliary forms therefore represent preserved layouts rather than complete working tools. The generated catalog lists every source event handler to support the remaining migration; an existing form is not a claim that all its handlers were translated.

The 24 missing project references are listed in `source-inventory.json`. The source includes historical broken names (for example `CAssoc.cls` versus `CAssocItem.cls`) and absent model/helper classes. ShellMain and frmMain share a VB_Name; file identity keeps them separate in this port. DataView/Form1 is named DataViewForm1.

## Verification boundaries

Core regressions cover variables/defaults/cancellation, unsupported commands, bounded recursion, separate output areas, locks, revision overflow, unchanged revisions, XML fidelity, duplicate IDs and DTD rejection. The form smoke runner constructs, shows and renders every form. It does not establish full UI interaction parity. The VSIX is compiled and its ZIP contents checked; live deployment, editor insertion and Undo are unverified for attempt 2.

The main MDB exported successfully: 25 categories, 419 templates. An attempted export of SandyCF-Home.mdb was rejected for orphan template 97. No orphan was dropped or reassigned, and that auxiliary library was not bundled.
