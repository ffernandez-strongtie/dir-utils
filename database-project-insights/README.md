# CSDirector database vs project file insights

`csdirector-database-project-insights.ps1` compares the trusses stored in the CSDirector database with the TrussStudio files inside each `.csdproj` archive and writes a browsable summary. It is read-only.

## Run

```powershell
$env:CSD_SQL_PASSWORD = '<password>'
.\csdirector-database-project-insights.ps1 -SqlUser insights_ro -Server 192.168.71.183 -ProjectsRoot C:\temp\csdproj-samples
```

Without `-SqlUser` it uses Windows authentication. The password is only read from `CSD_SQL_PASSWORD`.

| Parameter | Default | Purpose |
|---|---|---|
| `-ProjectsRoot` | `C:\SST\Server\Projects` | Folder with one subfolder per project holding `<ProjectNumber>_<timestamp>.CSDProj` |
| `-Server` / `-Database` | `.\SST` / `CSDatabaseServer` | SQL instance and database |
| `-SqlUser` | none | SQL login (password from `CSD_SQL_PASSWORD`) |
| `-OutputDirectory` | `.\output-insights` | Timestamp is appended |
| `-ProjectId` | all | Limit to specific project numbers |
| `-Limit` | 0 (all) | Stop after N projects |
| `-LegacyCutoffVersion` | `2024.7.2.7` | Truss versions strictly lower are before 2024.r7 (DB was the source of truth) |

The latest archive per project folder is used. Projects with no archive are not listed.

## Output

- `index.html`: one row per project with state, verdict, top causes and truss count; tabs and issue filters.
- `projects/<number>/project.html`: one row per truss with a symbol matrix; click a row for the field-by-field table.
- `projects/<number>/summary.json`, `findings.json`, `trusses/<truss>.json`: machine-readable detail.
- `rules.json`: rule id to plain name and description.

## States

- **Not in DB**: archive exists, no Project row.
- **Inconsistent**: at least one real mismatch (value, version, checksum, or truss missing on one side) while the DB is current.
- **Stale DB**: differences exist, but the DB check-in id is older than the archive timestamp (re-saved in Studio since).
- **Before 2024.r7**: every truss is DB-only at a version below the cutoff; nothing to compare.
- **Consistent**: no findings.

Symbols in the matrix: ✓ match, ≠ mismatch, ↻ stale DB, ⏸ before 2024.r7 DB-only, ∅ not comparable.

## Checks

Truss name sets (JSON in `Trusses/generated` vs `ComponentHeader`), `.tdlTruss` MD5 vs `ComponentHeader.TrussFileCheckSum`, JSON `Version` vs `ComponentTruss.AnalysisVersion`, and mapped quantity, ply, thickness, height, overhang, cantilever, heel, pitch, flag and spacing fields (numeric tolerance 0.001).
