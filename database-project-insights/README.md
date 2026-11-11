# CSDirector / TrussStudio Insights

`csdirector-database-project-insights.ps1` compares the trusses stored in the CSDirector database with the TrussStudio files inside each `.csdproj` archive and writes a browsable summary. It is read-only.

## Run

```powershell
$env:CSD_SQL_PASSWORD = '<password>'
.\csdirector-database-project-insights.ps1 -SqlUser '<sql-user>' -Server '<server-or-instance>' -ProjectsRoot 'C:\path\to\projects'
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
| `-Client` | off | Inspect the Client database instead of the Server one. Defaults become `C:\SST\Client\Projects` and `CSDatabaseClient`. A project folder with an expanded `Trusses` folder is read directly; otherwise its `.CSDProj` is used. The report and every project page are labelled CLIENT or SERVER |
| `-IncludeDbOnly` | off | Also write pages for DB projects with no archive folder (DB values only, "DB only" tab). Omitted by default |
| `-IncludeTrussJson` | off | Also write `projects/<number>/trusses/<truss>.json`. HTML detail and project-level JSON are always included |
| `-TimingLog` | off | Write `timing.csv` with phase durations, row counts, memory, CPU and garbage collection counters; project ordinals only |
| `-LegacyCutoffVersion` | `2024.7.2.7` | Truss versions strictly lower are before 2024.r7 (DB was the source of truth) |

The latest archive per project folder is used. Projects with no archive are not listed.

## Output

- `index.html`: one row per project with state, verdict, top causes and truss count; tabs and issue filters; aggregate coverage totals and a link to the contract manifest.
- `projects/<number>/project.html`: one row per truss with a symbol matrix; click a row for the field-by-field table, per-truss coverage (compared / different / not compared).
- `projects/<number>/summary.json`, `findings.json`: always included machine-readable project detail.
- `projects/<number>/trusses/<truss>.json`: optional, written only with `-IncludeTrussJson`; carries the per-truss coverage counts and the explicit comparison status of every mapped field (each with its comparison policy and basis; candidate rows are never flagged).
- `contract.json`: a manifest generated from the same map metadata as the comparison — every mapped pair with group, comparison policy and evidence basis, plus candidates, pending nested/structured encodings and pending lookup identities.
- `rules.json`: rule id to plain name and description.

## States

- **Not in DB**: archive exists, no Project row.
- **Inconsistent**: at least one real mismatch (value, version, checksum, or truss missing on one side) while the DB is current.
- **Stale DB**: differences exist, but the DB check-in id is older than the archive timestamp (re-saved in Studio since).
- **Before 2024.r7**: every truss is DB-only at a version below the cutoff; nothing to compare.
- **Consistent**: no differences in compared values; see coverage for values not compared.

Symbols in the matrix: ✓ match, ≠ mismatch, ↻ stale DB, ⏸ before 2024.r7 DB-only, ∅ not comparable.

Rows retain the compact report layout; check-in explanations are in the bottom legend. Causes use **Check-in times differ**, or **DB check-in is unknown** when the DB identifier is missing. Check-in identity is separate from truss data mismatches; Client mode remains informational only. A compact bottom legend explains symbols and comparison limits, with expandable cause definitions. Classification and mismatch counts are unchanged.

## Checks

Truss membership is the union of `Trusses/<name>.tdlTruss` in the archive and `ComponentHeader.Name` in the DB; JSON in `Trusses/generated` never defines membership, and a JSON file with no matching `.tdlTruss` is ignored. `.tdlTruss` MD5 vs `ComponentHeader.TrussFileCheckSum` is compared whenever a DB row and the `.tdlTruss` both exist, independent of JSON. JSON `Version` vs `ComponentTruss.AnalysisVersion`, and all mapped pairs in the contract table below, are compared only when the JSON exists; a missing JSON leaves those values not compared and is never itself a finding. Every mapped row in the per-truss detail has an explicit status: `match`, `differ`, `unresolved` (candidate mapping not confirmed), `file-only / not compared`, `DB-only / not compared`, `missing field / not compared`, `missing DB column / not compared`, `file NULL / not compared`, `DB NULL / not compared`, `missing side / not compared` or `both null / not compared`. Missing values are never matches; a both-null pair (both sides present, both NULL) is unknown, not equal and distinct from both sides being absent.

Coverage counts per truss and the project aggregate (`compared`, `different`, `notCompared`) cover mapped fields, version and checksum only; values that were not compared are not claims of equality. Unmapped field inventories are no longer built or included in HTML or JSON, so `notCompared` totals exclude those former inventory entries. Nested/structured mappings and FK lookup identities remain pending in `contract.json`; FK columns are never compared to human names.

Database reads use batches of up to 50 selected projects and a target maximum of 10,000 truss rows. A single project above that target is processed on its own. Count results are cached while projects wait for a later batch, so overlapping count windows do not repeat the same SQL work. Each batch then uses one check-in query and one truss query. Cached DB rows are released after each project, including analysis failures, and the cache is replaced between batches. Completed projects retain only compact index summaries.

DB-only project discovery streams project rows from SQL and stops reading once `-Limit` DB-only projects are found. DB-only project HTML stores mapped field descriptions once and a compact list of values per truss. Clicking a truss reconstructs the same field, value, status and coverage detail in the browser. Project HTML data is written in chunks of 128 trusses, and `-IncludeTrussJson` still writes the full per-truss JSON files. This bounds report-generation memory to a small chunk plus the current SQL batch.

A one-time metadata query checks which `ComponentTruss` columns exist. Absent mapped columns are omitted from SQL and shown as not compared, never as matches (candidate mappings remain unresolved); unavailable metadata stops the run with an explicit error. Queries select only identity, checksum, version and available mapped columns (including unresolved candidates); unused large columns such as `JigData` and `CarriedTrusses` are not fetched. `-ProjectId` and `-Limit` are applied before truss batches are loaded, including the DB-only path.

## Timing log

Add `-TimingLog` to write `timing.csv` in the timestamped output folder. The CSV records UTC time, elapsed time, phase duration, row and truss counts, report serialization and write time, HTML size, process memory and CPU time, and garbage collection counts. It uses project ordinals and the kind `archive` or `db_only`; it does not contain project names, paths, SQL text or field values. Start rows are written before potentially long SQL, analysis and report operations, so the last row also identifies an operation that has not finished.

The main phases are `sql_count`, `sql_projects`, `sql_trusses` (SQL fetch), `db_materialize` (turn fetched rows into PowerShell objects), `analysis`, `report`, `progress`, and `index`. Compare `duration_ms` for successive `report` rows with their `trusses` and `html_bytes` values. Rising `working_mb` or `private_mb` and `gc2` counts can indicate memory pressure. A slow `sql_trusses` row points to the database or transfer; a slow `db_materialize` row points to PowerShell object creation. `serialize_ms` and `write_ms` split out the chunked HTML payload work within `report`; `report` also includes other page and optional JSON work. Progress is updated at most about once every ten seconds, and the `progress` phase measures each update.

## Contract

This is a coverage expansion, not a complete cross-table mapping. The full contract also involves DB tables beyond `Project`, `ComponentHeader`, `Component` and `ComponentTruss`, which are not queried yet.

### Comparison policies

| Policy | Rule |
|---|---|
| `numeric` | absolute difference <= 0.001, applied only when both sides parse as numbers; otherwise ordinal case-sensitive text |
| `integer` | exact integer equality for quantity/ply/count values; ordinal text fallback |
| `boolean` | compared as bool or numeric 0/1 only; NULL is never coerced to false |
| `text` | ordinal case-sensitive exact string equality |
| Units and nulls | `*Inches` geometry values are in inches; pitch values are rise per 12; missing values are never matches — a missing field, an absent DB column, a database NULL, a file NULL and a missing side are distinguished; a both-null pair (both present, both NULL) is unknown, not true, and distinct from both sides being absent |

### Mapped pairs

Basis `documented persistence mapping` identifies pairs supported by persistence mapping documentation. These six fields are compared, not unresolved candidates. Weight and IsSymmetric are selected from `ComponentHeader`, Quantity from `Component`, and the remaining mapped values from `ComponentTruss`. `contract.json` records the source table per field. Live schema verification is still pending.

Basis `observed mapping` means the pair was checked against file and DB evidence; the mapping is observed, not derived from authoritative producer code. Basis `candidate` is unresolved: names/meaning are not confirmed, both side values are retained and the pair is never flagged as a mismatch. Additional flags not in the map are not fetched or included in output, and are never auto-confirmed from the same name.

| JSON field | DB column | Group | Policy | Basis |
|---|---|---|---|---|
| Quantity | ComponentQuantity | Identity | integer | observed mapping |
| NumPlies | TrussPlyCount | Geometry | integer | observed mapping |
| Thickness | TrussThicknessInches | Geometry | numeric | observed mapping |
| OverallTrussHeight | OverallTrussHeightInches | Geometry | numeric | observed mapping |
| LeftOverhang | LeftOverhangInches | Geometry | numeric | observed mapping |
| RightOverhang | RightOverhangInches | Geometry | numeric | observed mapping |
| LeftCantilever | LeftCantileverInches | Geometry | numeric | observed mapping |
| RightCantilever | RightCantileverInches | Geometry | numeric | observed mapping |
| LeftHeelHeight | LeftHeelHeightInches | Geometry | numeric | observed mapping |
| RightHeelHeight | RightHeelHeightInches | Geometry | numeric | observed mapping |
| PitchLeftTopOver12 | LeftTopPitchOver12 | Pitch | numeric | observed mapping |
| PitchRightTopOver12 | RightTopPitchOver12 | Pitch | numeric | observed mapping |
| PitchLeftBottomOver12 | LeftBottomPitchOver12 | Pitch | numeric | observed mapping |
| PitchRightBottomOver12 | RightBottomPitchOver12 | Pitch | numeric | observed mapping |
| IsAttic | IsAttic | Flags | boolean | observed mapping |
| IsGable | IsGable | Flags | boolean | observed mapping |
| IsGirder | IsGirder | Flags | boolean | observed mapping |
| IsFlipped | IsFlipped | Flags | boolean | observed mapping |
| OCSpacing | OCSpacing | Geometry | numeric | observed mapping |
| Height | TrussHeightInches | Geometry | numeric | observed mapping |
| LeftHeelToPeak | LeftHeelToPeakInches | Geometry | numeric | observed mapping |
| RightHeelToPeak | RightHeelToPeakInches | Geometry | numeric | observed mapping |
| LeftTopChordHeelToPeak | LeftTopChordHeelToPeakInches | Geometry | numeric | observed mapping |
| RightTopChordHeelToPeak | RightTopChordHeelToPeakInches | Geometry | numeric | observed mapping |
| LeftStub | LeftStubInches | Geometry | numeric | observed mapping |
| RightStub | RightStubInches | Geometry | numeric | observed mapping |
| TCPitchCount | TopChordPitchCount | Pitch | integer | observed mapping |
| BCPitchCount | BottomChordPitchCount | Pitch | integer | observed mapping |
| Camber | Camber | Geometry | numeric | observed mapping |
| CamberDistance | CamberDistanceInches | Geometry | numeric | observed mapping |
| CamberNote | CamberNote | Geometry | text | observed mapping |
| BearingPadNotes | BearingPadNotes | Geometry | text | observed mapping |
| LeftOverhangRake | LeftOverhangRake | Geometry | numeric | observed mapping |
| RightOverhangRake | RightOverhangRake | Geometry | numeric | observed mapping |
| SinglePlyOverhang | SinglePlyOverhang | Flags | boolean | observed mapping |
| SinglePlyFill | SinglePlyFill | Flags | boolean | observed mapping |
| FillPly | FillPly | Geometry | integer | observed mapping |
| OverhangPly | OverhangPly | Geometry | integer | observed mapping |
| TCRakeFirst | TCRakeFirst | Geometry | numeric | observed mapping |
| TCRakeLast | TCRakeLast | Geometry | numeric | observed mapping |
| CompatibilityVersion | CompatibilityVersion | Version | text | observed mapping |
| TrussMatchCode | TrussMatchCode | Geometry | text | observed mapping |
| Span | ComponentTruss.TrussSpanInches | Geometry | numeric | documented persistence mapping |
| ProdSpan | ComponentTruss.FullSpanInches | Geometry | numeric | documented persistence mapping |
| IsStructural | ComponentTruss.IsStructuralGable | Flags | boolean | documented persistence mapping |
| Weight | ComponentHeader.WeightPounds | Geometry | numeric | documented persistence mapping |
| PlyToPlyNote | ComponentTruss.PlyToPly | Geometry | text | documented persistence mapping |
| IsSymmetric | ComponentHeader.IsSymmetrical | Flags | boolean | documented persistence mapping |

Identity and integrity checks that are not rows of the map: truss membership from `Trusses/<name>.tdlTruss` against `ComponentHeader.Name` (membership is a set; pairings are not compared, and JSON never defines membership), `.tdlTruss` MD5 vs `TrussFileCheckSum` (independent of JSON), JSON `Version` vs `AnalysisVersion` (only when the JSON exists). `Project.FilesCheckInId` vs the archive timestamp is compared only on Server mode; on Client mode it is informational-only within the contract.

### Pending

- `CarriedTrusses` and `JigData` — structured mappings pending; DB payloads are not fetched, never string-compared.
- Unmapped nested JSON objects/arrays — not included in output; relational mapping pending.
- FK columns (`*Key`) — never compared to human names; lookup table identities pending.
- Full cross-table mapping — not claimed; the contract currently covers the queried tables only.
