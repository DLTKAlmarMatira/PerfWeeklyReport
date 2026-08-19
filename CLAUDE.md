# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Repo Is

A self-contained PowerShell pipeline that extracts performance-testing data from Deltek's **on-prem TFS** server and produces a weekly HTML dashboard and CSV reports. It runs automatically via GitHub Actions on a self-hosted Windows runner inside the corporate network.

This repo is a nested git repository inside the larger `C:\CLAUDEPROJ\perfprocess` workspace. The parent has its own `CLAUDE.md` and `ADO/CLAUDE.md` with deeper domain context — read those for background on ADO structure, link-type rules, and metric definitions.

## Commands

```powershell
# Full pipeline: extract from ADO + build all reports
Run-AdoExtracts.bat

# Rebuild reports from already-pulled CSVs (no network, no credentials)
powershell -File 6-Build-WeeklyReports.ps1
powershell -File 7-Build-MeetingReport.ps1 -Show     # -Show opens in browser

# Ad-hoc analysis against the joined dataset
.\6-Build-WeeklyReports.ps1 -PassThru | Where-Object Outcome -eq 'Failed' | Group-Object Product

# Diagnostic: verify fields before adding to EXTRA_FIELDS
powershell -File 4-Get-AdoWorkItemTypeFields.ps1 -OutputPath csv\workitemtype_fields.csv

# Re-pull one extract individually
powershell -NoProfile -ExecutionPolicy Bypass -File 1-Get-TestPlanResults.ps1 -OutputPath csv\test_plan_results.csv
```

## Pipeline Architecture

The batch runs scripts in this order: **1 → 2 → 3 → 8 → 6 → 7**. Script 8 was added after 6 and 7 were already numbered; its output (`csv\pbi_bug_links.csv`) is consumed by script 7.

| Step | Script | Network | Produces |
|---|---|---|---|
| 1 | `1-Get-TestPlanResults.ps1` | yes | `csv\test_plan_results.csv` |
| 2 | `2-Get-AdoQueryResults.ps1` | yes | `csv\pbi_task_links.csv` |
| 3 | `3-Get-TaskTestsLinkResults.ps1` | yes | `csv\task_tests_link_results.csv` |
| 4 | `8-Get-PbiBugLinks.ps1` | yes | `csv\pbi_bug_links.csv` |
| 5 | `6-Build-WeeklyReports.ps1` | **no** | 4 derived CSVs |
| 6 | `7-Build-MeetingReport.ps1` | **no** | `weekly_meeting_report.html` |

**Raw extracts** (`csv\test_plan_results.csv`, `pbi_task_links.csv`, `task_tests_link_results.csv`, `pbi_bug_links.csv`) are committed to git — they are point-in-time ADO snapshots that cannot be recovered. The derived CSVs and the HTML report are `.gitignore`d because they are recomputable.

Scripts 1–4 share one auth mechanism: PAT lookup order is `-Pat` arg → DPAPI cache at `%LOCALAPPDATA%\AdoTestPlanExtractor\pat.dat` → `$env:AZURE_DEVOPS_PAT` → prompt. **Press Enter at the prompt to use Windows auth** (the normal path for the on-prem server). Clear a stale PAT with `-ResetPat`.

## GitHub Actions Automation

`.github/workflows/weekly-perf-report.yml` runs the pipeline on a cron schedule (01:07 UTC every Tuesday = 9:07 AM Philippine Time) and emails `weekly_meeting_report.html` to the team.

**Required**: a self-hosted Windows runner with network access to `tfs.deltek.com`, labeled `perf-report`. GitHub-hosted runners cannot reach the on-prem server.

**Secrets needed**: `AZURE_DEVOPS_PAT` (TFS: Test Management Read + Work Items Read) and `SMTP_PASSWORD` (for `smtp.deltek.com`). The workflow can be triggered manually with `workflow_dispatch`.

## PowerShell Conventions

These traps have already bitten this code — don't undo the fixes:

- **`$PSScriptRoot` is empty inside `param()` defaults** under `powershell.exe -File`. Scripts 6, 7, and 8 declare bare params and resolve path defaults in the body using `$PSScriptRoot` after the scope is established.
- **`-ExtraFields` must be one comma-separated string, not multiple tokens.** Across the `cmd.exe → powershell.exe` boundary a multi-token list overflows positionally into `-Pat`/`-ApiVersion`. `Run-AdoExtracts.bat` passes one quoted string; scripts 2 and 3 split it themselves.
- **`Microsoft.VSTS.Common.ActivatedDate` does not exist on this org.** Use `Microsoft.VSTS.Common.StateChangeDate` instead. Do not add `ActivatedDate` back.
- **Script 7 is pure ASCII.** Glyphs are built from `[char]` codes shipped through JSON (`\uXXXX`). PowerShell 5.1 reads BOM-less scripts as Windows-1252, so a literal non-ASCII character in the source becomes mojibake in the output. Don't paste symbols directly into it.
- **Variable name shadowing is silent.** PowerShell variable names are case-insensitive: a loop-local `$neverRun` silently overwrites a constant `$NeverRun`. Script 6 names its outcome constant `$OutcomeNeverRun` to avoid this.
- **Never accumulate rows with `+=` in a loop.** With thousands of test points and hundreds of links this is quadratic. Script 6 uses hashtables of `List[object]`.
- **API version probing.** The on-prem server does not support `api-version=7.1`. All network scripts probe downward (7.1 → 7.0 → 6.0 → 5.1 → …) when `-ApiVersion` is not passed. An empty-but-successful response usually means a version mismatch, not absent data.
