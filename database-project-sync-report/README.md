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
| `-Limit` | Process the first N selected projects in sorted order; `0` means no limit. |
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

Each run creates a new directory using the output base plus a UTC timestamp and unique suffix, including when a custom output base is supplied.

- `index.html`: overview with links to project details and an expandable legend.
- `projects/`: individual project HTML reports.
- `summary.json` and `settings.json`: machine-readable results and run settings.
- `evidence/`: supporting copies and inventories when `-Evidence True` is enabled.

Open `index.html` in a browser. Temporary working files are removed by default; originals and SQL records remain unchanged.

[Back to script index](../README.md)
