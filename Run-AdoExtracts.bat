@echo off
setlocal

rem Runs the 3 ADO extraction scripts in sequence and then builds the derived
rem weekly reports from them, all output going to the same "csv" folder next
rem to this .bat file. Stops immediately if any step fails instead of
rem continuing with bad/missing data.
rem
rem Steps 1-3 are the extraction (they talk to ADO and may prompt for
rem credentials). Step 4 is pure local CSV transformation - no network, no
rem credentials - so it is safe to re-run on its own at any time via
rem   powershell -File 6-Build-WeeklyReports.ps1
rem
rem Edit the QueryUrl on step 2 below if your "PBI - Active" query URL/ID
rem ever changes.
rem
rem Steps 2 and 3 now also pull extra fields (via -ExtraFields), confirmed to
rem actually exist on Task/Product Backlog Item via
rem 4-Get-AdoWorkItemTypeFields.ps1 (see csv\workitemtype_fields.csv):
rem   - System.ChangedDate / Microsoft.VSTS.Common.StateChangeDate /
rem     Microsoft.VSTS.Common.ClosedDate: for "worked on this week" /
rem     "completed this week" reporting (NOTE: ActivatedDate does NOT exist
rem     on this org's Task/PBI types - do not add it back, use
rem     StateChangeDate instead)
rem   - Microsoft.VSTS.CMMI.Blocked: real "Blocked" field on both types, for
rem     tasks-with-blockers reporting (still need to confirm the team
rem     actually sets this - see workitemtype_fields.csv analysis)
rem   - System.IterationPath: sprint/iteration the item belongs to, needed
rem     to eventually resolve iteration end dates for deadline tracking
rem   - Microsoft.VSTS.Scheduling.RemainingWork (Task) / .Effort (PBI): once
rem     we have iteration end dates (via 5-Get-AdoIterationCapacity.ps1
rem     -ListIterations), Remaining Work vs. days left in the sprint is a
rem     first-pass "on track / at risk" signal for anticipating delay

set "SCRIPT_DIR=%~dp0"
set "OUTDIR=%SCRIPT_DIR%csv"
set "PBI_QUERY_URL=https://tfs.deltek.com/tfs/Deltek/QEAutomation/_queries/query/8a4c81ea-6f2a-4b1b-b2ae-948d4b8b2a12"
rem NOTE: single comma-separated string, passed as ONE quoted argument below.
rem Space-separating them as multiple quoted tokens does NOT reliably bind to
rem -ExtraFields either - PowerShell's array binder only grabs the first token
rem when ExtraFields isn't the script's last parameter, and the rest overflow
rem into -Pat/-ApiVersion positionally ("positional parameter cannot be
rem found" errors). Both 2-Get-AdoQueryResults.ps1 and
rem 3-Get-TaskTestsLinkResults.ps1 now split this string on commas themselves,
rem so a single argument is the reliable way to pass multiple field names
rem across the cmd.exe -> powershell.exe process boundary.
set "EXTRA_FIELDS=System.ChangedDate,Microsoft.VSTS.Common.StateChangeDate,Microsoft.VSTS.Common.ClosedDate,Microsoft.VSTS.CMMI.Blocked,System.IterationPath,Microsoft.VSTS.Scheduling.RemainingWork,Microsoft.VSTS.Scheduling.Effort"

rem Step 2 pulls two EXTRA fields on top of the shared list above, because the
rem deadline dates live on the PBI and only step 2 sees PBIs:
rem   - System.CreatedDate      : used as the work START date. Automatic and
rem     read-only. It is really "when the ticket was made", which drifts from
rem     "when work began" whenever PBIs are created in a planning batch - so
rem     treat it as a trend indicator, not a commitment date.
rem   - Deltek.PlanHotFixRelDt  : repurposed as the work TARGET/END date.
rem     THIS IS A DELIBERATE WORKAROUND. Product Backlog Item has no
rem     StartDate/TargetDate/DueDate field in this org (verified against the
rem     full 112-field inventory), and the Planning*/Completed* fields are
rem     integers, not dates. This field is date-typed and effectively unused
rem     here (1 work item project-wide had a value). ADO still labels it
rem     "Planned Hot Fix Release Date" - the HTML report relabels it to
rem     "Target". If Feature StartDate/TargetDate ever get populated (they
rem     already work on Feature 2834211), switch to those and drop this.
set "PBI_EXTRA_FIELDS=%EXTRA_FIELDS%,System.CreatedDate,Deltek.PlanHotFixRelDt"

if not exist "%OUTDIR%" mkdir "%OUTDIR%"

echo ========================================
echo [1/7] Test Plan run results
echo ========================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%1-Get-TestPlanResults.ps1" -OutputPath "%OUTDIR%\test_plan_results.csv"
if errorlevel 1 (
    echo.
    echo [1/7] FAILED. Stopping.
    goto :end
)

echo.
echo ========================================
echo [2/7] PBI to Task links
echo ========================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%2-Get-AdoQueryResults.ps1" -QueryUrl "%PBI_QUERY_URL%" -OutputPath "%OUTDIR%\pbi_task_links.csv" -ExtraFields "%PBI_EXTRA_FIELDS%"
if errorlevel 1 (
    echo.
    echo [2/7] FAILED. Stopping.
    goto :end
)

echo.
echo ========================================
echo [3/7] Task to Test Case/Test Plan "Tests" links
echo ========================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%3-Get-TaskTestsLinkResults.ps1" -OutputPath "%OUTDIR%\task_tests_link_results.csv" -ExtraFields "%EXTRA_FIELDS%"
if errorlevel 1 (
    echo.
    echo [3/7] FAILED. Stopping.
    goto :end
)

echo.
echo ========================================
echo [4/7] PBI to Bug "Related" links
echo ========================================
rem Bugs hang off a PBI with a plain "Related" link, which the saved query in
rem step 2 does NOT return - it asks only for Parent/Child. This step walks the
rem relations through the REST API instead, so no ADO query edit is needed.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%8-Get-PbiBugLinks.ps1" -CsvDir "%OUTDIR%"
if errorlevel 1 (
    echo.
    echo [4/7] FAILED. Stopping.
    goto :end
)

echo.
echo ========================================
echo [5/7] Latest discussion entries
echo ========================================
rem Fetches the most-recent discussion comment for every PBI and Task so the
rem HTML report can show it on demand. Optional: if this step fails the rest
rem of the pipeline still completes - discussion icons will simply not appear.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%9-Get-WorkItemComments.ps1" -CsvDir "%OUTDIR%"
if errorlevel 1 (
    echo.
    echo [5/7] WARNING: discussion fetch failed - continuing without it.
    echo       Discussion icons will not appear in the HTML report.
    echo       Re-run manually with:
    echo         powershell -File "%SCRIPT_DIR%9-Get-WorkItemComments.ps1"
)

echo.
echo ========================================
echo [6/7] Build derived weekly reports
echo ========================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%6-Build-WeeklyReports.ps1" -CsvDir "%OUTDIR%"
if errorlevel 1 (
    echo.
    echo [6/7] FAILED. The raw extracts above are still good - you can
    echo       re-run just this step with:
    echo         powershell -File "%SCRIPT_DIR%6-Build-WeeklyReports.ps1"
    goto :end
)

echo.
echo ========================================
echo [7/7] Build HTML meeting report
echo ========================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%7-Build-MeetingReport.ps1" -CsvDir "%OUTDIR%"
if errorlevel 1 (
    echo.
    echo [7/7] FAILED. The CSVs above are still good - you can
    echo       re-run just this step with:
    echo         powershell -File "%SCRIPT_DIR%7-Build-MeetingReport.ps1"
    goto :end
)

echo.
echo ========================================
echo All done: extracts + bug links + discussion + weekly CSVs + HTML report.
echo   CSVs        : %OUTDIR%
echo   HTML report : %SCRIPT_DIR%weekly_meeting_report.html
echo ========================================

:end
pause
endlocal
