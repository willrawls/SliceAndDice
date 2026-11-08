# Archived source supplements

William authorized matching archived files to fill missing Sandy source references. `Sandy` remains the primary behavioral source; no original source files are modified.

`archive-matches.json` records candidates and SHA-256 hashes. The main model candidates are in `Archive/sad2`, whose project identifies version 1.8.198, later than Sandy's 1.8.160. They are supplemental evidence rather than proof that behavior was identical.

Reviewed for the Blazor model migration:

- `CTemplate.cls`: four memoCode areas, filename, short name, lock/undeletable/favorite flags, revision and timestamps. The C# model retains these fields and the Razor editing state exposes them without reducing the original field dictionary.
- `CCategory.cls`: template ownership plus category metadata. Category identity/ownership and unknown fields remain in the model.
- `CInsertionInfo.cls`: explicit separate insertion/output state. The existing core continues to return output blocks without performing implicit host mutations.
- `CSadCommands.cls`: command-provider collections and command help. This remains reference material for unfinished provider/help behavior; no claim of complete migration.
- `CAssoc.cls`: declares `CAssocItem`, confirming the project filename/class-name mismatch. The contemporary Sandy folder already has CAssocItem.cls.

`Form1.frm` is ambiguous: archive copies include a SQL stored-procedure utility. A matching filename alone does not establish that it is Sandy/Project1.vbp's missing form. It is not incorporated into the 21-form inventory.

The remaining candidates need behavior review during their respective handler migrations. Discovering a candidate does not mean that its code is already ported.
