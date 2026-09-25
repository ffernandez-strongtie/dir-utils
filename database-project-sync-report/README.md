# Database and Project File Sync Report

**Script:** [csdirector-database-project-file-sync-report.ps1](./csdirector-database-project-file-sync-report.ps1)

Reads project archives and database records without modifying either. It writes reports and working copies to separate locations; it does not synchronize or repair data.

## Requirements

- Windows PowerShell 5.1 or later on Windows, with the .NET SQL Server and ZIP APIs available.
- Read access to the project archives and required CSDirector database tables. SQL Server connections use the Windows account running the script.
- Write access to the output location and, unless evidence is retained, the Windows temporary directory.

## Usage

Run from the `database-project-sync-report` directory:

```powershell
# Interactive: press Enter to accept the displayed defaults.
.\csdirector-database-project-file-sync-report.ps1

# Investigate one project and retain supporting evidence.
.\csdirector-database-project-file-sync-report.ps1 -ProjectId SAM-101 -Evidence True

# Compare selected projects, including optional timestamp checks.
.\csdirector-database-project-file-sync-report.ps1 -ProjectId S-1206,S-1328 -Timestamp

# Supply all prompted settings explicitly.
.\csdirector-database-project-file-sync-report.ps1 `
    -ProjectsRoot 'C:\SST\Server\Projects' `
    -Server '.\SST' `
    -Database 'CSDatabaseServer' `
    -OutputDirectory '.\reports\comparison' `
    -Limit 2
```

The script prompts for the projects root, SQL server, database, and output base when those parameters are omitted. Their defaults are `C:\SST\Server\Projects`, `.\SST`, `CSDatabaseServer`, and `output` under the current working directory.

| Option | Behavior |
| --- | --- |
| `-ProjectId` | Restrict the scan to one or more project folder identifiers. |
| `-Limit` | Include at most N projects. Archive projects come first, ranked by their newest eligible csdproj modification time descending; ties use project identifier. All archives for each selected project are included. Only after archive projects are exhausted do DB-only projects fill remaining slots, newest DB modification first. `0` means no limit. |
| `-Evidence True` | Retain copied archives, extracted JSON, and detailed comparison inventories. Default: `False`. |
| `-Timestamp` | Enable timestamp comparisons. Off by default. |
| `-FileTimeZone`, `-DatabaseTimeZone` | Interpret timestamps using the specified Windows time zones. Each defaults to the current machine's time zone. |
| `-ToleranceSeconds` | Timestamp comparison tolerance. Default: `120`. |
| `-NumericTolerance` | Selected numeric-field comparison tolerance. Default: `0.000001`. |

## What is compared

- **Inventory:** canonical `Trusses/<name>.tdlTruss` entries versus database truss names. Retained JSON does not establish truss membership.
- **Selected fields:** 19 truss properties, including quantity, plies, dimensions, pitches, flags, and spacing. This is not a comparison of every design property.
- **Checksums:** MD5 of each uncompressed canonical `.tdlTruss` file versus the checksum recorded in the database.
- **Board footage:** archive and database geometry calculations using the same length basis, alongside Director's stored project total.
- **Suspected duplicates:** repeated piece business data within a DB truss, excluding record identity and audit fields. These are review candidates, not automatically confirmed errors.
- **Optional timestamps:** JSON ZIP-entry times versus database times. Dates are investigation clues, not proof of which design is correct.

Project identifiers come from the first folder beneath `ProjectsRoot`. Each archive receives its own report row. The scan skips `DeletedProjects` folders, directory reparse points, and `.csdproj` filenames beginning with `Attachment_` or `Attachments_`, regardless of capitalization.

## Reading the BDFT columns

| Column | Meaning |
| --- | --- |
| **csdproj** | Footage calculated from archive materials, piece lengths, plies, and layout quantities using the saved truss length mode. |
| **DB (calculated)** | Footage calculated from DB piece dimensions, lengths, and plies using the file's length modes and layout quantities. |
| **Difference** | `csdproj` minus `DB (calculated)`: the comparison on the same basis. |
| **DB (As stored)** | Director's unchanged cached project total, shown for reference. |

`*` marks **Pick Length**. Unmarked totals use **Actual Length**, or both stored length-specific calculations give identical footage. `(Mixed)` identifies a total combining length modes. `(basis unverified)` means the method cannot be established from the available evidence; it does not by itself mean the total is wrong. An em dash means the calculation is unavailable.

The file's explicit truss setting takes precedence over its archive environment setting. If required settings or lengths are missing or invalid, the calculation is unavailable; another length type is not substituted. Both formula totals use nominal dimensions for dimensional lumber and actual catalog dimensions for engineered lumber. They compare geometry and do not reproduce Director pricing exceptions or predict a cache refresh.

Both formula totals use file layout quantities and membership. DB quantity differences and DB-only trusses remain separate checks. The report checks differences per truss as well as in total, so offsetting differences cannot establish agreement.

**Healthy** requires agreement on inventory, selected fields, checksums, comparable BDFT, and quantities, with no suspected duplicate pieces. A different Director stored total or length method alone does not prevent Healthy; a yellow BDFT notice remains visible.

## Output

Each run creates a new directory with a lowercase name using the output base, sanitized company name, UTC timestamp, and unique suffix, including when a custom output base is supplied (for example, `output-spates-fabricators-inc-timestamp-suffix`).

The index and every project page show the company at the right, its unique Server Reference beneath it, and the report creation date/time in UTC below that. All pages use the same timestamp captured once per run, also saved as `ReportCreatedUtc` in `settings.json`. All distinct company names come from `dbo.Company.Name` with company type `Our Company`, sorted alphabetically and shown together. The full list is saved as `Owner.CompanyNames` in `settings.json`; their combined name is used for the sanitized folder name (limited to 80 characters). The reference comes from the first line of `C:\SST\Server\Util\KeyUp.ini`; use `-KeyUpPath` for another installation. For remote or restored databases, select the matching INI file. Missing identity information produces warnings and unavailable labels without blocking the report. Missing company names use `unknown-company` in folder names. Identity sources and lookup notes are saved in `settings.json`.

- `index.html`: overview with links to project details and an expandable legend. The first column shows the source csdproj file’s last-modified time in UTC, or the DB project’s last-modified wall time for projects without archives. Each section is sorted newest first, independently of the optional truss timestamp comparisons.
- `projects/`: individual project HTML reports.
- `summary.json` and `settings.json`: machine-readable results and run settings.
- `evidence/`: supporting copies and inventories when `-Evidence True` is enabled.

Open `index.html` in a browser. Temporary working files are removed by default; originals and SQL records remain unchanged.

[Back to script index](../README.md)

## Legacy projects before 2024.r7

The delayed 2024.r7 release changed the source of truth from database values to project files. The actual customer upgrade date is unknown. The report marks a project/archive as **Possibly Legacy [2024.r7]** only when its archive was read successfully, no truss JSON exists, a unique database project has truss records, and all checked project, component/header/truss, and piece modification dates are known and on or before January 8, 2025 (inclusive, database wall time). Empty projects and failed/missing evidence do not qualify.

These rows appear yellow in their own **Possibly Legacy [2024.r7]** section, newest first. Comparison values are suppressed in HTML, console output, and summary JSON. The explanation and latest DB modification time remain in the summary. Projects with later or unknown dates and no JSON remain unresolved; the release date does not establish when a customer upgraded. Use `-LegacyCutoffDate yyyy-MM-dd` to change the approximate cutoff when better deployment information is available.

## Database projects without archives

After **Empty**, the index lists **DB projects without csdproj**, sorted by DB project modification time descending (unknown dates last). The first column shows the DB project’s last-modified time (database wall time, not creation time). The DB truss count remains in its normal inventory column; unavailable values use gray cells with `-`. Each row links to a project page carrying the same report identity header. These records are also included in `summary.json`.

This check compares database projects against the full eligible archive inventory under `ProjectsRoot`, before `-Limit` is applied. `-Limit` is shared: archive projects use the budget first, then DB-only projects fill remaining slots in descending DB modification order; `-ProjectId` restricts both lists. Existing exclusions apply (DeletedProjects, directory reparse points, and Attachment_/Attachments_ archives). “No csdproj found” describes this scan scope, not proof that no archive exists elsewhere.
