# Advanced

[Back to the quick start](README.md#run)

## Use a separate SQL login

Use this example instead of the Windows sign-in command in [step 2](README.md#2-run-the-report). Replace each `<...>` placeholder, including the angle brackets, with the details you were given. Replace `C:\path\to\projects` with your projects folder path.

```powershell
$env:CSD_SQL_PASSWORD = '<password>'
.\csdirector-database-project-insights.ps1 -SqlUser '<sql-user>' -Server '<server-or-instance>' -ProjectsRoot 'C:\path\to\projects'
```

## Command-line options

| Parameter | Default | Purpose |
|---|---|---|
| `-ProjectsRoot` | `C:\SST\Server\Projects` | Folder with one subfolder per project holding `<ProjectNumber>_<timestamp>.CSDProj` |
| `-Server` / `-Database` | `.\SST` / `CSDatabaseServer` | SQL instance and database |
| `-SqlUser` | none | SQL login (password from `CSD_SQL_PASSWORD`) |
| `-OutputDirectory` | `.\output-insights` | Timestamp is appended |
| `-ProjectId` | all | Limit to specific project numbers |
| `-Limit` | 0 (all) | Stop after N projects |
| `-Client` | off | Inspect the Client database instead of the Server one. Defaults become `C:\SST\Client\Projects` and `CSDatabaseClient`. A project folder with an expanded `Trusses` folder is read directly; otherwise its `.CSDProj` is used. The report and every project page are labelled CLIENT or SERVER |
| `-IncludeDbOnly` | off | Also write pages for DB projects with no archive folder (DB values only). Qualifying pre-cutoff projects appear in the existing "Before 2024.r7" tab; the rest appear in "DB only". Omitted by default |
| `-IncludeTrussJson` | off | Also write `projects/<number>/trusses/<truss>.json`. HTML detail and project-level JSON are always included |
| `-TimingLog` | off | Testing aid: write `timing.csv` with phase durations, row counts, memory, CPU and garbage collection counters; project ordinals only |
| `-ParallelProjects` | 2 | Analyze and write up to two projects at once after each SQL batch loads. Use `1` for sequential processing |
| `-LegacyCutoffVersion` | `2024.7.2.7` | Truss versions strictly lower are before 2024.r7 (DB was the source of truth) |

The latest archive per project folder is used. Projects with no archive are not listed.

## Timing log (testing aid)

Add `-TimingLog` to write `timing.csv` in the timestamped output folder. The CSV records UTC time, elapsed time, phase duration, row and truss counts, report serialization and write time, HTML size, process memory and CPU time, and garbage collection counts. It uses project ordinals and the kind `archive` or `db_only`; it does not contain project names, paths, SQL text or field values. Start rows are written before potentially long SQL, analysis and report operations, so the last row also identifies an operation that has not finished.

The main phases are `sql_count`, `sql_projects`, `sql_trusses_stream` (query, row transfer and conversion into the batch cache), `analysis`, `report`, `progress`, and `index`. With two workers, `parallel_project` also records elapsed time from dispatch through collection; `analysis` and `report` remain separate. Rising `working_mb` or `private_mb` and `gc2` counts can indicate memory pressure. A slow `sql_trusses_stream` row identifies the combined database and row-loading work. In sequential mode, `serialize_ms` and `write_ms` split out the chunked HTML payload work within `report`; the parallel `report` row records its total time and HTML size. Progress is updated at most about once every ten seconds, and the `progress` phase measures each update.
