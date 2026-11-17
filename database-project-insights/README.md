# CSDirector / TrussStudio Insights

`csdirector-database-project-insights.ps1` compares the trusses stored in the CSDirector database with the TrussStudio files saved inside each `.csdproj` project archive, then writes a report you open in your browser. It only reads: it does not change the database or any project file.

[Download the PowerShell script](csdirector-database-project-insights.ps1?raw=true). If it opens as text, press **Ctrl+S** and save it as `csdirector-database-project-insights.ps1`.

## Run

### 1. Open PowerShell in the script folder

In File Explorer, open the folder that holds `csdirector-database-project-insights.ps1`, click the address bar, type `powershell` and press Enter. A PowerShell window opens in that folder.

### 2. Run the report

For a standard server setup, copy and run this command. It uses your Windows sign-in:

```powershell
.\csdirector-database-project-insights.ps1
```

If PowerShell says running scripts is disabled, paste this into the same window and press Enter:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

This applies only to that window, ends when the window closes, and does not override your organization's Group Policy. If the block persists, ask IT. Then retry the command above.

### 3. Open the report

Back in File Explorer, open the newest `output-insights-<timestamp>` folder in the folder where you opened PowerShell. Double-click `index.html` to view the report in your browser.

[Advanced setup](ADVANCED.md) · [Report reference](REPORT-REFERENCE.md)
