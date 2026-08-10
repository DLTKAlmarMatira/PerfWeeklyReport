<#
.SYNOPSIS
    Builds the derived weekly performance-testing reports from the three raw
    ADO extracts produced by scripts 1-3.

.DESCRIPTION
    This is "layer 2" of the pipeline. Layer 1 is Run-AdoExtracts.bat, which
    pulls the raw CSVs out of ADO; this script joins them into the reports the
    weekly meeting actually reads.

        Run-AdoExtracts.bat            ->  csv\test_plan_results.csv
          1-Get-TestPlanResults.ps1        csv\pbi_task_links.csv
          2-Get-AdoQueryResults.ps1        csv\task_tests_link_results.csv
          3-Get-TaskTestsLinkResults.ps1
                                            |
        this script                         v
                                       csv\connected_pbi_task_test_results.csv
                                       csv\pbi_summary_weekly.csv
                                       csv\tester_workload_weekly.csv
                                       csv\exceptions_weekly.csv

    These four files used to be built by hand every week. Every column turned
    out to be mechanically derivable, so all of that logic lives here now.
    Nothing in the outputs requires human judgement.

    Unlike scripts 1-5 this one talks to no network and needs no credentials -
    it is pure local CSV transformation, so it is safe to re-run at any time.

.PARAMETER CsvDir
    Folder holding the raw extracts, and where the reports are written.
    Default: the "csv" folder next to this script.

.PARAMETER PassThru
    Also emit the joined rows to the pipeline, for ad-hoc analysis, e.g.
        .\6-Build-WeeklyReports.ps1 -PassThru | Where-Object Outcome -eq 'Failed'

.EXAMPLE
    .\6-Build-WeeklyReports.ps1

.EXAMPLE
    .\6-Build-WeeklyReports.ps1 -CsvDir D:\snapshots\2026-08-05

.NOTES
    HOW THE JOIN WORKS

    Both link extracts come from ADO "work items and direct links" queries, so
    they hold one row per LINK plus a "(root)" row per top-level work item. The
    root rows have an empty Source and must be skipped - they are query
    scaffolding, not relationships. Of 616 rows in task_tests_link_results.csv,
    62 are roots and 554 are real links.

    A Task->Test Suite link fans out to every test point in that suite; a
    Task->Test Case link fans out to every test point for that case (a case can
    sit in several suites). That fan-out turns 554 links into 3922 report rows.

    IMPORTANT: the fan-out joins on Suite ID, never on Suite Path. Suite paths
    are NOT unique - 6 of them are shared by multiple suite IDs, which yields
    100 (Test Case, Suite Path) pairs carrying genuinely different outcomes.
    Suite ID is carried into the output so those rows stay distinguishable; the
    older hand-built file omitted it, which left them impossible to tell apart.
#>

[CmdletBinding()]
param(
    [string]$CsvDir,
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'

# Resolve the default HERE, not in the param() block. Under
# `powershell.exe -File <script>` the parameter defaults are evaluated before
# the script scope is established, so $PSScriptRoot is still empty there and
# `Join-Path $PSScriptRoot 'csv'` dies with "Cannot bind argument to parameter
# 'Path' because it is an empty string". In the body it is populated.
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir) { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $CsvDir)    { $CsvDir = Join-Path $ScriptDir 'csv' }

# Raw ADO link type -> our policy label. This is a function of the link type
# ALONE; the target's work item type is deliberately not consulted.
#
# The "MISTAKE" wording encodes a team policy documented in
# ADO_Performance_Testing_Structure.md: a Task->Test Case/Test Plan
# relationship must use a "Tests" link, because "Child" means work breakdown
# while "Tests" means validation evidence. Using Child breaks traceability.
$LinkClassification = @{
    'System.LinkTypes.Hierarchy-Forward'                     = 'Child (MISTAKE - should be Tests)'
    'Microsoft.VSTS.Common.TestedBy-Reverse'                 = 'Tests (correct)'
    'Microsoft.VSTS.TestCase.SharedStepReferencedBy-Forward' = 'Tests (correct)'
}

$CorrectLink = 'Tests (correct)'
$RootLink    = '(root)'

# NOTE the deliberately distinct name. PowerShell variable names are
# CASE-INSENSITIVE, so a loop-local $neverRun would be the *same variable* as a
# constant named $NeverRun and would silently overwrite it with a row count on
# the first iteration - after which every outcome lookup misses and returns 0.
# Keep this name distinct from any local, and never introduce $neverRunOutcome.
$OutcomeNeverRun = 'Unspecified'

# Outcomes worth a human's attention, alongside any row whose link type is
# wrong. Plain "Unspecified" (never run) is NOT an exception on its own -
# there are ~2000 of those and they would drown out the real signal.
$ExceptionOutcomes = @('Failed', 'Blocked')

# ---------------------------------------------------------------------------
# DEADLINE SOURCE - the single place to change if the source of truth moves.
#
# Product Backlog Item has NO StartDate/TargetDate/DueDate field in this org
# (verified against the full 112-field inventory), and the Planning*/Completed*
# fields are integers, not dates. So the perf team repurposes two existing
# PBI date fields as an interim convention:
#
#     start  <- System.CreatedDate       automatic, read-only
#     target <- Deltek.PlanHotFixRelDt   hand-entered; ADO still labels this
#                                        "Planned Hot Fix Release Date"
#
# Both are read from the SOURCE side of pbi_task_links.csv, because the PBI is
# the source of a PBI->Task link.
#
# If Feature StartDate/TargetDate ever get populated (both already work on
# Feature 2834211), point these two at the Feature instead - nothing else in
# this script needs to change.
# ---------------------------------------------------------------------------
$PbiStartColumn  = 'Source System.CreatedDate'
$PbiTargetColumn = 'Source Deltek.PlanHotFixRelDt'

$Invariant = [System.Globalization.CultureInfo]::InvariantCulture


function Get-CleanIdentity {
    <#
        Pull a person's display name out of either shape ADO identities arrive
        in. Two shapes exist in these extracts because the scripts differ:

          1-Get-TestPlanResults.ps1 projects assignedTo.displayName, so test
          points arrive already clean:
              Almar Matira <ADSDELTEKCOM\AlmarMatira>

          2-/3-Get-*.ps1 hand the whole identity OBJECT to Export-Csv, which
          stringifies it as a PowerShell hashtable literal:
              @{displayName=Almar Matira; url=https://...; uniqueName=...}

        The second is really an extraction-side bug - those scripts should
        project displayName the way script 1 does - but it is parsed here so
        the reports work against the extracts as they exist today.

        Name/login mismatches ("Lotte Morales" vs login "CharlotteMorales")
        come straight out of ADO's own displayName field. We deliberately do
        not reconcile them: guessing at identity mapping would be worse than
        showing what ADO actually holds.
    #>
    param([string]$Raw)

    if ([string]::IsNullOrWhiteSpace($Raw)) { return '' }
    $Raw = $Raw.Trim()

    if ($Raw.StartsWith('@{')) {
        if ($Raw -match 'displayName=([^;}]*)') { return $Matches[1].Trim() }
        return ''
    }
    return ($Raw -split ' <')[0].Trim()
}


function Get-ProductFromIteration {
    <#
        QEAutomation\Costpoint\Sprint 12  ->  Costpoint

        The second segment is the product. A bare "QEAutomation" means the item
        was never filed under a product.
    #>
    param([string]$IterationPath)

    if ([string]::IsNullOrWhiteSpace($IterationPath)) { return '(no product)' }
    $parts = @(($IterationPath -replace '/', '\') -split '\\' | Where-Object { $_.Trim() })
    if ($parts.Count -ge 2) { return $parts[1] }
    return '(no product)'
}


function Get-TaskKind {
    <#
        Classify a task by its title: Scripting authors test cases, Execution
        runs a test plan. This maps cleanly onto the data - Scripting tasks
        link Test Cases and Execution tasks link Test Suites - so it is worth
        having as an explicit column.
    #>
    param([string]$Title)

    $lowered = "$Title".ToLowerInvariant()
    if ($lowered -match 'test execution' -or $lowered -match '-\s*execution\b') { return 'Execution' }
    if ($lowered -match 'scripting') { return 'Scripting' }
    return 'Other'
}


function Format-Rate {
    <#
        Percentage to 1 dp, or N/A when nothing has been run yet.

        Two deliberate choices here:

        1. AwayFromZero (half-up), because these are percentages a human reads
           and cross-checks in Excel, and Excel's ROUND() is half-up. Both
           PowerShell's default [math]::Round and Python's "%.1f" use banker's
           rounding (half-to-EVEN) instead, which turns an exact 31.25 into
           31.2 - correct for summing money, surprising in a status report.
           This only ever moves an exact .x5 midpoint by 0.1.

        2. InvariantCulture, because "{0:0.0}" would emit a comma decimal
           separator under a European locale, making the reports disagree with
           themselves depending on who ran them.
    #>
    param([double]$Numerator, [double]$Denominator)

    if ($Denominator -le 0) { return 'N/A' }
    $pct = 100.0 * $Numerator / $Denominator
    return ([math]::Round($pct, 1, [System.MidpointRounding]::AwayFromZero)).ToString('0.0', $Invariant)
}


function New-Tally {
    # Outcome counters plus the identity sets both summaries need.
    [pscustomobject]@{
        Total    = 0
        Outcomes = @{}
        Mistakes = 0
        Testers  = [System.Collections.Generic.HashSet[string]]::new()
        Pbis     = [System.Collections.Generic.HashSet[string]]::new()
    }
}


function Add-Row {
    param($Tally, $Row)

    $Tally.Total++

    $outcome = $Row.Outcome
    if ($Tally.Outcomes.ContainsKey($outcome)) { $Tally.Outcomes[$outcome]++ }
    else { $Tally.Outcomes[$outcome] = 1 }

    if ($Row.'Link Classification' -ne $CorrectLink) { $Tally.Mistakes++ }

    $tester = Get-CleanIdentity $Row.Tester
    if ($tester) { [void]$Tally.Testers.Add($tester) }

    if (-not [string]::IsNullOrWhiteSpace($Row.'PBI ID')) { [void]$Tally.Pbis.Add($Row.'PBI ID') }
}


function Get-Count {
    param($Tally, [string]$Outcome)
    if ($Tally.Outcomes.ContainsKey($Outcome)) { return $Tally.Outcomes[$Outcome] }
    return 0
}


function Write-Report {
    param([string]$Path, $Rows)

    $count = @($Rows).Count
    if ($count -eq 0) {
        Write-Warning "No rows to write for $([IO.Path]::GetFileName($Path)) - skipped."
        return
    }
    $Rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
    Write-Host ("  {0,-45} {1,5} rows" -f [IO.Path]::GetFileName($Path), $count)
}


# ---- Main ----

try {
    if (-not (Test-Path -LiteralPath $CsvDir)) {
        throw "CSV folder not found: $CsvDir"
    }

    $testPlanPath  = Join-Path $CsvDir 'test_plan_results.csv'
    $pbiTaskPath   = Join-Path $CsvDir 'pbi_task_links.csv'
    $taskTestsPath = Join-Path $CsvDir 'task_tests_link_results.csv'

    $missing = @($testPlanPath, $pbiTaskPath, $taskTestsPath | Where-Object { -not (Test-Path -LiteralPath $_) })
    if ($missing.Count -gt 0) {
        throw ("Missing raw extract(s):`n  {0}`nRun Run-AdoExtracts.bat first." -f ($missing -join "`n  "))
    }

    $testPlan  = @(Import-Csv -LiteralPath $testPlanPath)
    $pbiTask   = @(Import-Csv -LiteralPath $pbiTaskPath)
    $taskTests = @(Import-Csv -LiteralPath $taskTestsPath)
    Write-Host ("Read  test_plan={0}  pbi_task={1}  task_tests={2}" -f $testPlan.Count, $pbiTask.Count, $taskTests.Count)

    # --- Index the test points. Hashtable lookups rather than Where-Object in
    # --- a loop: with ~2400 points and ~550 links the naive version is
    # --- quadratic and takes minutes.
    $pointsBySuite = @{}
    $pointsByCase  = @{}
    foreach ($point in $testPlan) {
        $suiteId = $point.'Suite ID'
        if (-not $pointsBySuite.ContainsKey($suiteId)) {
            $pointsBySuite[$suiteId] = [System.Collections.Generic.List[object]]::new()
        }
        $pointsBySuite[$suiteId].Add($point)

        $caseId = $point.'Test Case ID'
        if (-not $pointsByCase.ContainsKey($caseId)) {
            $pointsByCase[$caseId] = [System.Collections.Generic.List[object]]::new()
        }
        $pointsByCase[$caseId].Add($point)
    }

    # Task ID -> its parent PBI link row. Skip the query's "(root)" rows.
    $pbiByTask = @{}
    foreach ($link in $pbiTask) {
        if ($link.'Source System.WorkItemType' -eq 'Product Backlog Item') {
            $pbiByTask[$link.'Target ID'] = $link
        }
    }

    # --- Fan the links out to test points.
    $connected = [System.Collections.Generic.List[object]]::new()
    $unknownLinkTypes = @{}

    foreach ($link in $taskTests) {
        $rawLink = $link.'Link Type'
        if ($rawLink -eq $RootLink) { continue }

        $targetType = $link.'Target System.WorkItemType'
        $targetId   = $link.'Target ID'

        if ($targetType -eq 'Test Suite') {
            $points = $pointsBySuite[$targetId]
        } else {
            $points = $pointsByCase[$targetId]
        }
        if ($null -eq $points) { continue }

        $classification = $LinkClassification[$rawLink]
        if (-not $classification) {
            # Surface new link types loudly rather than silently mislabelling
            # them as correct - a new type is a process change worth noticing.
            $classification = "UNKNOWN ($rawLink)"
            if ($unknownLinkTypes.ContainsKey($rawLink)) { $unknownLinkTypes[$rawLink]++ }
            else { $unknownLinkTypes[$rawLink] = 1 }
        }

        $taskId    = $link.'Source ID'
        $taskTitle = $link.'Source System.Title'
        $pbi       = $pbiByTask[$taskId]

        # Prefer the task's own iteration; fall back to the PBI's.
        $iteration = $link.'Source System.IterationPath'
        if ([string]::IsNullOrWhiteSpace($iteration) -and $pbi) {
            $iteration = $pbi.'Source System.IterationPath'
        }

        # A Task with no PBI parent yields blanks here. That is a real
        # condition in the data (Task 2831364), not a defect.
        $pbiId    = if ($pbi) { $pbi.'Source ID' }                else { '' }
        $pbiTitle = if ($pbi) { $pbi.'Source System.Title' }       else { '' }
        $pbiState = if ($pbi) { $pbi.'Source System.State' }       else { '' }
        $pbiStart  = if ($pbi) { [string]$pbi.$PbiStartColumn }    else { '' }
        $pbiTarget = if ($pbi) { [string]$pbi.$PbiTargetColumn }   else { '' }

        $product  = Get-ProductFromIteration $iteration
        $kind     = Get-TaskKind $taskTitle
        $assignee = Get-CleanIdentity $link.'Source System.AssignedTo'

        foreach ($point in $points) {
            $outcome = $point.Outcome
            if ([string]::IsNullOrWhiteSpace($outcome)) { $outcome = $OutcomeNeverRun }

            $connected.Add([pscustomobject][ordered]@{
                'PBI ID'              = $pbiId
                'PBI Title'           = $pbiTitle
                'PBI State'           = $pbiState
                # Blank until Run-AdoExtracts.bat is re-run with the new
                # PBI_EXTRA_FIELDS, and until someone fills the target in ADO.
                # Import-Csv yields $null for a column that isn't there, so an
                # older extract degrades to empty rather than throwing.
                'PBI Start'           = $pbiStart
                'PBI Target'          = $pbiTarget
                'Product'             = $product
                'Task ID'             = $taskId
                'Task Title'          = $taskTitle
                'Task State'          = $link.'Source System.State'
                'Task Kind'           = $kind
                'Task Assignee'       = $assignee
                'Link Type (raw)'     = $rawLink
                'Link Classification' = $classification
                'Target ID'           = $targetId
                'Target Type'         = $targetType
                'Target Title'        = $link.'Target System.Title'
                'Suite Path'          = $point.'Suite Path'
                'Suite ID'            = $point.'Suite ID'
                'Test Case ID'        = $point.'Test Case ID'
                'Test Case Title'     = $point.'Test Case Title'
                'Outcome'             = $outcome
                'Configuration'       = $point.Configuration
                'Tester'              = $point.'Assigned To'
                # Placeholder for human commentary. Never populated so far;
                # kept so the column doesn't vanish from downstream consumers.
                'Note'                = ''
            })
        }
    }

    if ($unknownLinkTypes.Count -gt 0) {
        Write-Warning 'Unrecognised link type(s) - $LinkClassification needs updating:'
        foreach ($entry in $unknownLinkTypes.GetEnumerator()) {
            Write-Warning ("   {0}  ({1} links)" -f $entry.Key, $entry.Value)
        }
    }

    if ($connected.Count -eq 0) {
        throw 'Join produced no rows - check that the extracts are all for the same project.'
    }

    # --- Per-PBI summary, in first-appearance order with the no-parent group last.
    $pbiGroups = [System.Collections.Specialized.OrderedDictionary]::new()
    $pbiTitles = @{}
    foreach ($row in $connected) {
        $key = $row.'PBI ID'
        if (-not $pbiGroups.Contains($key)) {
            $pbiGroups[$key] = New-Tally
            $pbiTitles[$key] = @($row.'PBI Title', $row.'PBI State', $row.'PBI Start', $row.'PBI Target')
        }
        Add-Row $pbiGroups[$key] $row
    }

    $orderedKeys = @(
        @($pbiGroups.Keys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) +
        @($pbiGroups.Keys | Where-Object { [string]::IsNullOrWhiteSpace($_) })
    )

    $pbiSummary = [System.Collections.Generic.List[object]]::new()
    foreach ($key in $orderedKeys) {
        $t = $pbiGroups[$key]
        $neverRunCount = Get-Count $t $OutcomeNeverRun
        $run           = $t.Total - $neverRunCount
        $passed        = Get-Count $t 'Passed'

        $pbiSummary.Add([pscustomobject][ordered]@{
            'PBI ID'    = $key
            'PBI Title' = $pbiTitles[$key][0]
            'PBI State' = $pbiTitles[$key][1]
            # Date-only: the raw values carry a time component that is noise
            # for a weekly report and makes the CSV awkward in Excel.
            'Start'     = $(if ($pbiTitles[$key][2]) { ([datetime]$pbiTitles[$key][2]).ToString('yyyy-MM-dd') } else { '' })
            'Target'    = $(if ($pbiTitles[$key][3]) { ([datetime]$pbiTitles[$key][3]).ToString('yyyy-MM-dd') } else { '' })
            'Testers'   = (@($t.Testers) | Sort-Object) -join ', '
            # Row count, not distinct test cases: a case linked via several
            # tasks/suites is genuinely several units of execution work.
            'Total Test Cases'         = $t.Total
            'Passed'                   = $passed
            'Failed'                   = Get-Count $t 'Failed'
            'Blocked'                  = Get-Count $t 'Blocked'
            'Not Applicable'           = Get-Count $t 'NotApplicable'
            'Unspecified (never run)'  = $neverRunCount
            'Run Rate %'               = Format-Rate $run $t.Total
            # NotApplicable counts as run - it is a deliberate verdict, not an
            # absence of one.
            'Pass Rate % (of run)'     = Format-Rate $passed $run
            'Link Type Mistakes'       = $t.Mistakes
        })
    }

    # --- Per-tester workload, busiest first.
    $testerGroups = @{}
    foreach ($row in $connected) {
        $tester = Get-CleanIdentity $row.Tester
        if (-not $testerGroups.ContainsKey($tester)) { $testerGroups[$tester] = New-Tally }
        Add-Row $testerGroups[$tester] $row
    }

    $testerWorkload = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $testerGroups.GetEnumerator()) {
        $t = $entry.Value
        $neverRunCount = Get-Count $t $OutcomeNeverRun
        $run           = $t.Total - $neverRunCount
        $passed        = Get-Count $t 'Passed'

        $testerWorkload.Add([pscustomobject][ordered]@{
            'Tester' = $entry.Key
            # Excludes the no-PBI-parent group; "distinct PBIs" should count
            # real PBIs only.
            'Distinct PBIs'            = $t.Pbis.Count
            'Total Test Cases'         = $t.Total
            'Passed'                   = $passed
            'Failed'                   = Get-Count $t 'Failed'
            'Blocked'                  = Get-Count $t 'Blocked'
            'Unspecified (never run)'  = $neverRunCount
            'Run Rate %'               = Format-Rate $run $t.Total
            'Pass Rate % (of run)'     = Format-Rate $passed $run
        })
    }
    $testerWorkload = @($testerWorkload | Sort-Object -Property 'Total Test Cases' -Descending)

    # --- Exceptions: wrong link type, or a bad outcome.
    $exceptions = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $connected) {
        $isException = ($row.'Link Classification' -ne $CorrectLink) -or
                       ($ExceptionOutcomes -contains $row.Outcome)
        if (-not $isException) { continue }

        $exceptions.Add([pscustomobject][ordered]@{
            'PBI ID'              = $row.'PBI ID'
            'PBI Title'           = $row.'PBI Title'
            'Product'             = $row.Product
            'Task ID'             = $row.'Task ID'
            'Task Title'          = $row.'Task Title'
            'Task Assignee'       = $row.'Task Assignee'
            'Link Classification' = $row.'Link Classification'
            'Test Case ID'        = $row.'Test Case ID'
            'Test Case Title'     = $row.'Test Case Title'
            'Suite Path'          = $row.'Suite Path'
            'Suite ID'            = $row.'Suite ID'
            'Outcome'             = $row.Outcome
            'Tester Clean'        = (Get-CleanIdentity $row.Tester)
            'Note'                = ''
        })
    }

    Write-Host 'Wrote:'
    Write-Report (Join-Path $CsvDir 'connected_pbi_task_test_results.csv') $connected
    Write-Report (Join-Path $CsvDir 'pbi_summary_weekly.csv')              $pbiSummary
    Write-Report (Join-Path $CsvDir 'tester_workload_weekly.csv')          $testerWorkload
    Write-Report (Join-Path $CsvDir 'exceptions_weekly.csv')               $exceptions

    $mistakes = @($connected | Where-Object { $_.'Link Classification' -ne $CorrectLink }).Count
    Write-Host ''
    Write-Host ("Link type mistakes (Child used where Tests belongs): {0}" -f $mistakes)
    Write-Host ("Exceptions needing review: {0}" -f $exceptions.Count)

    if ($PassThru) { $connected }
}
catch {
    Write-Error "Failed: $($_.Exception.Message)"
    exit 1
}
