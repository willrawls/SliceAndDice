# Source investigation and migration decision

The preserved sources were reviewed before implementation. No legacy source or database was changed.

| Source | Evidence and decision |
| --- | --- |
| `Sandy` and `SandyRadD` | Main VBP version 1.8.160; twelve referenced source files absent in each directory. Their frmMain files match SandyRadE byte-for-byte. |
| `SandyRadE` | Main VBP version 1.8.160; all main-project source references present. Secondary comparison reference. |
| `Archive/sad2` | Main VBP version 1.8.198; all 25 main-project source references present. Chosen behavior baseline. This is source completeness, not a verified VB6 build. |
| `Archive/Sandy2` | Useful split into instance/support/templates/IDE interfaces. SandyForVS targets the old Development Environment 6.0 and includes incomplete source declarations. Use the abstraction idea, not its adapter implementation. |
| `SandyB`, `SandyC` | Incomplete C# translations. SandyC is a Framework 4.8 WinForms executable and retains unresolved VB6/ActiveX types and empty members. Not used as the new foundation. |

## Important behavioral anchors

- `Archive/sad2/frmMain.frm`: `ppInclude`, `Preprocess`, `DoInsertion`, `FillTemplateWithUserInput`, `InternalInsertTemplate`, `ChangeFocusOfInsertion`.
- `Archive/sad2/CTemplate.cls`: template fields, flags, revision counter and persistence.
- `Archive/sad2/CAssocArray.cls`: case-insensitive keys and built-in EOL values.
- `Archive/sad2/CSadAddin.cls`, `CSadCommands.cls`, `NewCommands.cls`: Sandal lifecycle, command metadata, aliases and primitive operations.
- `Archive/sad2/sadFile.*`: external file, process, clipboard and keyboard automation.
- `Archive/sad2/frmDBClassGen.frm`, `sadSqlSp/`: database-driven generation and SQL Server metadata.
- `Archive/Sandy2/ISandyIDE.cls`, `ICodeModule.cls`: precedent for keeping host APIs outside the engine.

Preserve cursor → declarations → bottom → file execution order with one run's variables shared across all areas. Includes default to the declarations section; explicit area 2 selects cursor, 3 selects bottom, and other nonzero selectors select file. A declaration boundary is language-specific, not the first byte of a file.

The old `GotoSoftVar` dispatch has a misspelled sentinel compared with its flush implementation. The new interpreter implements the documented variable-output intent rather than reproducing that typo. Patch import/export is also present as menu/API surface but commented out in 1.8.198; it is not counted as functioning legacy behavior.

## Corpus inspected and tested

| XML library | Categories | Templates |
| --- | ---: | ---: |
| SliceAndDice 2000 | 27 | 550 |
| SandyClipboard | 4 | 36 |
| SandyColdFusion | 16 | 171 |
| SandySqlSp | 24 | 328 |
| OldSliceAndDice | 23 | 313 |
| GTESliceAndDice | 21 | 380 |
| Copy of SandyBeneMO | 26 | 419 |

These libraries overlap; counts must not be described as unique templates. Tests import and roundtrip every category/template field. XML normalizes source line endings on read; subsequent export preserves the loaded values and any CRLFs introduced by editing. The original files remain untouched.

The newer csSliceAndDice CSV contains 706 records and uses 22 fields, with a different ordering of code fields than the 18-column layout. It is deliberately not guessed. A verified schema/import adapter is still required.

A text scan of the 550-template library found 132 AddControl and 403 SetProperty occurrences. Designer behavior therefore needs an explicit migration track. Preserving template syntax does not translate emitted VB6 source into C#.

## SDK decision

The initial recommendation considered in-process VSSDK for the widest IDE access. The requested .NET 10 runtime changes the hosting choice: this implementation uses out-of-process VisualStudio.Extensibility and a dual-target Core. The separate engine permits another host later without rewriting the language.

References checked during implementation:

- [Microsoft extensibility model comparison](https://learn.microsoft.com/en-us/visualstudio/extensibility/visualstudio-extensibility/how-it-works/models/extensibility-models?view=visualstudio)
- [Official extension samples](https://github.com/microsoft/VSExtensibility/tree/main/New_Extensibility_Model/Samples)
- [Version-checked editor mutations](https://learn.microsoft.com/en-us/visualstudio/extensibility/visualstudio.extensibility/editor/walkthroughs/editing-text?view=visualstudio)
- [Runtime metadata](https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.extensibility.extensionmetadata.dotnettargetversions?view=vs-extensibility)

The code is newly written against those APIs. It does not include copied sample implementations or legacy COM binaries.
