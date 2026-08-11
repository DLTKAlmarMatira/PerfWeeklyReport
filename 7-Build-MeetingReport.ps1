<#
.SYNOPSIS
    Builds a single self-contained HTML dashboard for the weekly performance
    testing meeting, from the joined dataset that 6-Build-WeeklyReports.ps1
    produces.

.DESCRIPTION
    Answers the three questions the weekly meeting actually asks:

      1. Which active tasks are assigned to which person?
      2. How many test cases/tests are involved for the execution?
      3. Of the execution that came from a test plan - how many Passed,
         Failed, Blocked, Not Applicable?

    Output is ONE .html file with the data, styling, and interactivity all
    embedded. No server, no CDN, no dependencies - double-click it, or attach
    it to an email. Filters (person / product / state / task kind / text) sit
    in a single row and re-scope every number on the page at once, so the
    tiles, the chart, and the table can never disagree.

    SCOPING RULE THAT MATTERS

    The outcome numbers count only test points reached through a Test Suite
    link - i.e. "came from a test plan", which is what an Execution task
    links. Scripting tasks link Test Cases directly, and those same test
    cases usually also sit in an execution suite, so counting both would
    double-count every result. Scripting tasks therefore show their test-case
    count and a dash for outcomes. This is why the table's outcome columns sum
    exactly to the tiles.

.PARAMETER CsvDir
    Folder holding connected_pbi_task_test_results.csv.
    Default: the "csv" folder next to this script.

.PARAMETER OutputPath
    Where to write the HTML.
    Default: weekly_meeting_report.html next to this script.

.PARAMETER Show
    Open the report in the default browser when done.

.EXAMPLE
    .\7-Build-MeetingReport.ps1 -Show

.NOTES
    Requires 6-Build-WeeklyReports.ps1 to have run first (Run-AdoExtracts.bat
    does both, in order). Talks to no network and needs no credentials.

    The colour choices are not arbitrary. Outcomes are a STATUS scale, so they
    use fixed status colours rather than a categorical palette. The stack order
    (Passed, Not applicable, Blocked, Failed, Not started) was chosen because
    putting Passed next to Failed is a red/green pair that colourblind readers
    cannot separate - measured at deltaE 4.1 under deuteranopia, a hard fail.
    Reordered, the worst adjacent pair is 7.9, which is only acceptable
    alongside secondary encoding - hence every status carries a glyph AND a
    text label, the segments are separated by 2px surface gaps, a legend is
    always present, and the chart has a table view. Do not "simplify" any of
    those away.
#>

[CmdletBinding()]
param(
    [string]$CsvDir,
    [string]$OutputPath,
    [switch]$Show
)

$ErrorActionPreference = 'Stop'

# Resolve defaults HERE, not in the param() block. Under
# `powershell.exe -File <script>` the parameter defaults are evaluated before
# the script scope is established, so $PSScriptRoot is still empty there and
# `Join-Path $PSScriptRoot ...` dies with "Cannot bind argument to parameter
# 'Path' because it is an empty string". In the body it is populated.
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir)  { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $CsvDir)     { $CsvDir = Join-Path $ScriptDir 'csv' }
if (-not $OutputPath) { $OutputPath = Join-Path $ScriptDir 'weekly_meeting_report.html' }

function Get-CleanIdentity {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return '' }
    $Raw = $Raw.Trim()
    if ($Raw.StartsWith('@{')) {
        if ($Raw -match 'displayName=([^;}]*)') { return $Matches[1].Trim() }
        return ''
    }
    return ($Raw -split ' <')[0].Trim()
}

# These two must stay byte-identical to the versions in
# 6-Build-WeeklyReports.ps1. They classify the tasks that script 6 never sees
# (the ones with no linked tests), so if the two drift, the same task would be
# labelled differently depending on which file you read.
function Get-ProductFromIteration {
    # QEAutomation\Costpoint\Sprint 12 -> Costpoint
    param([string]$IterationPath)
    if ([string]::IsNullOrWhiteSpace($IterationPath)) { return '(no product)' }
    $parts = @(($IterationPath -replace '/', '\') -split '\\' | Where-Object { $_.Trim() })
    if ($parts.Count -ge 2) { return $parts[1] }
    return '(no product)'
}

# Ages are computed HERE, at build time, against a single reference instant
# embedded in the page. If the browser computed "days ago" against its own
# clock, a report opened next Tuesday would silently relabel last week's work
# as current. A generated report should be a fixed snapshot.
$AsOf = Get-Date

function Get-DaysSince {
    param([string]$IsoDate)
    if ([string]::IsNullOrWhiteSpace($IsoDate)) { return -1 }   # -1 = no date
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($IsoDate, [ref]$parsed)) {
        $days = [int][math]::Floor(($AsOf - $parsed).TotalDays)
        if ($days -lt 0) { return 0 }
        return $days
    }
    return -1
}

function Get-DaysUntil {
    # Days from the frozen asOf to a future date. NEGATIVE means overdue, so
    # unlike Get-DaysSince this must NOT clamp at zero.
    param([string]$IsoDate)
    if ([string]::IsNullOrWhiteSpace($IsoDate)) { return $null }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($IsoDate, [ref]$parsed)) {
        return [int][math]::Floor(($parsed.Date - $AsOf.Date).TotalDays)
    }
    return $null
}

function Get-DateOnly {
    param([string]$IsoDate)
    if ([string]::IsNullOrWhiteSpace($IsoDate)) { return '' }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($IsoDate, [ref]$parsed)) { return $parsed.ToString('yyyy-MM-dd') }
    return ''
}

function Get-TaskKind {
    # Scripting authors test cases; Execution runs a test plan.
    param([string]$Title)
    $lowered = "$Title".ToLowerInvariant()
    if ($lowered -match 'test execution' -or $lowered -match '-\s*execution\b') { return 'Execution' }
    if ($lowered -match 'scripting') { return 'Scripting' }
    return 'Other'
}

try {
    $connectedPath = Join-Path $CsvDir 'connected_pbi_task_test_results.csv'
    if (-not (Test-Path -LiteralPath $connectedPath)) {
        throw ("Not found: $connectedPath`n" +
               "Run 6-Build-WeeklyReports.ps1 first (or just run Run-AdoExtracts.bat, which does both).")
    }

    $connected = @(Import-Csv -LiteralPath $connectedPath)
    if ($connected.Count -eq 0) { throw "No rows in $connectedPath" }

    # The connected dataset is keyed on TEST links, so a task that has no test
    # case or test plan linked yet produces no rows and would vanish from this
    # report entirely. That silently understates everyone's workload - 22 active
    # tasks were missing before this was added. The task universe therefore
    # comes from the PBI->Task extract, and the test data is overlaid onto it.
    $pbiTaskPath = Join-Path $CsvDir 'pbi_task_links.csv'
    $pbiTaskRows = @()
    if (Test-Path -LiteralPath $pbiTaskPath) {
        $pbiTaskRows = @(Import-Csv -LiteralPath $pbiTaskPath |
                         Where-Object { $_.'Source System.WorkItemType' -eq 'Product Backlog Item' })
    } else {
        Write-Warning "pbi_task_links.csv not found - tasks with no linked tests will be missing from the report."
    }

    # --- Bugs linked to each PBI, from 8-Get-PbiBugLinks.ps1. Optional: an
    # --- older csv folder simply yields no bug counts rather than an error.
    # --- Only Target Type = 'Bug' counts - a "Related" link also points at
    # --- other PBIs and Tasks, so counting all of them would be wrong.
    $bugsByPbi = @{}
    $bugUnreadable = 0
    $bugLinksPath = Join-Path $CsvDir 'pbi_bug_links.csv'
    if (Test-Path -LiteralPath $bugLinksPath) {
        foreach ($r in (Import-Csv -LiteralPath $bugLinksPath)) {
            if ($r.Readable -eq 'False') { $bugUnreadable++; continue }
            if ($r.'Target Type' -ne 'Bug') { continue }
            # NOT $pid - that is a READ-ONLY automatic variable (the current
            # process id), and PowerShell variable names are case-insensitive,
            # so assigning to it throws "Cannot overwrite variable PID".
            # Same family of trap as the $neverRun/$NeverRun collision in
            # script 6; see the PowerShell traps section in ADO/CLAUDE.md.
            $bugPbiId = $r.'PBI ID'
            if (-not $bugsByPbi.ContainsKey($bugPbiId)) {
                $bugsByPbi[$bugPbiId] = [System.Collections.Generic.HashSet[string]]::new()
            }
            [void]$bugsByPbi[$bugPbiId].Add($r.'Target ID')
        }
        Write-Host ("Bug links: {0} PBI(s) carry bugs; {1} related target(s) unreadable" -f $bugsByPbi.Count, $bugUnreadable)
    } else {
        Write-Warning "pbi_bug_links.csv not found - bug counts will be blank. Run 8-Get-PbiBugLinks.ps1."
    }

    # --- Activity dates. These live on the work item itself, so they come from
    # --- the link extracts rather than the connected dataset (script 6 doesn't
    # --- carry them). StateChangeDate and ClosedDate are 100% / correctly
    # --- populated; RemainingWork and Blocked are NOT (empty on every row in
    # --- this org), which is why there is no at-risk or blocker reporting here.
    $taskDates = @{}
    foreach ($link in $pbiTaskRows) {
        $taskDates[$link.'Target ID'] = @{
            changed = $link.'Target Microsoft.VSTS.Common.StateChangeDate'
            closed  = $link.'Target Microsoft.VSTS.Common.ClosedDate'
        }
    }
    # Fallback for tasks with no PBI parent, which never appear above.
    $taskTestsPath = Join-Path $CsvDir 'task_tests_link_results.csv'
    if (Test-Path -LiteralPath $taskTestsPath) {
        foreach ($link in (Import-Csv -LiteralPath $taskTestsPath)) {
            if ($link.'Link Type' -eq '(root)') { continue }
            $id = $link.'Source ID'
            if ($taskDates.ContainsKey($id)) { continue }
            $taskDates[$id] = @{
                changed = $link.'Source Microsoft.VSTS.Common.StateChangeDate'
                closed  = $link.'Source Microsoft.VSTS.Common.ClosedDate'
            }
        }
    }

    $required = @('Task ID','Task Title','Task State','Task Kind','Task Assignee',
                  'Product','PBI ID','PBI Title','Target Type','Outcome',
                  'Test Case ID','Tester','Link Classification')
    $present = $connected[0].PSObject.Properties.Name
    $absent = @($required | Where-Object { $present -notcontains $_ })
    if ($absent.Count -gt 0) {
        throw ("$connectedPath is missing column(s): {0}`nRe-run 6-Build-WeeklyReports.ps1 to regenerate it." -f ($absent -join ', '))
    }

    Write-Host ("Read {0} joined rows" -f $connected.Count)

    # --- Aggregate to one record per task. Filters are all task-level
    # --- attributes, so the browser can re-scope everything by filtering this
    # --- small array - no need to ship 3900 raw rows to the page.
    $byTask = [System.Collections.Specialized.OrderedDictionary]::new()

    function New-TaskRecord {
        param([string]$Id, [string]$Title, [string]$State, [string]$Kind,
              [string]$Assignee, [string]$Product, [string]$PbiId, [string]$PbiTitle,
              [string]$PbiStart, [string]$PbiTarget)
        if ([string]::IsNullOrWhiteSpace($Assignee)) { $Assignee = '(unassigned)' }
        [pscustomobject]@{
            id = $Id; title = $Title; state = $State; kind = $Kind
            assignee = $Assignee; product = $Product; pbiId = $PbiId; pbiTitle = $PbiTitle
            pbiStart = $PbiStart; pbiTarget = $PbiTarget
            CaseSet   = [System.Collections.Generic.HashSet[string]]::new()
            TesterSet = [System.Collections.Generic.HashSet[string]]::new()
            exec = 0; passed = 0; failed = 0; blocked = 0; na = 0; never = 0; mistakes = 0
        }
    }

    # Seed every task that exists under a PBI, including ones with no tests yet.
    foreach ($link in $pbiTaskRows) {
        $id = $link.'Target ID'
        if ($byTask.Contains($id)) { continue }
        $title = $link.'Target System.Title'
        $byTask[$id] = New-TaskRecord -Id $id -Title $title `
            -State $link.'Target System.State' `
            -Kind (Get-TaskKind $title) `
            -Assignee (Get-CleanIdentity $link.'Target System.AssignedTo') `
            -Product (Get-ProductFromIteration $link.'Target System.IterationPath') `
            -PbiId $link.'Source ID' -PbiTitle $link.'Source System.Title' `
            -PbiStart  ([string]$link.'Source System.CreatedDate') `
            -PbiTarget ([string]$link.'Source Deltek.PlanHotFixRelDt')
    }

    foreach ($row in $connected) {
        $id = $row.'Task ID'
        if (-not $byTask.Contains($id)) {
            $assignee = $row.'Task Assignee'
            if ([string]::IsNullOrWhiteSpace($assignee)) { $assignee = '(unassigned)' }
            $byTask[$id] = [pscustomobject]@{
                id        = $id
                title     = $row.'Task Title'
                state     = $row.'Task State'
                kind      = $row.'Task Kind'
                assignee  = $assignee
                product   = $row.Product
                pbiId     = $row.'PBI ID'
                pbiTitle  = $row.'PBI Title'
                pbiStart  = [string]$row.'PBI Start'
                pbiTarget = [string]$row.'PBI Target'
                CaseSet   = [System.Collections.Generic.HashSet[string]]::new()
                TesterSet = [System.Collections.Generic.HashSet[string]]::new()
                exec      = 0
                passed    = 0
                failed    = 0
                blocked   = 0
                na        = 0
                never     = 0
                mistakes  = 0
            }
        }
        $t = $byTask[$id]

        [void]$t.CaseSet.Add($row.'Test Case ID')
        if ($row.'Link Classification' -ne 'Tests (correct)') { $t.mistakes++ }

        # Outcomes only from test-plan-sourced points. See SCOPING RULE above.
        if ($row.'Target Type' -eq 'Test Suite') {
            $t.exec++
            switch ($row.Outcome) {
                'Passed'        { $t.passed++ }
                'Failed'        { $t.failed++ }
                'Blocked'       { $t.blocked++ }
                'NotApplicable' { $t.na++ }
                default         { $t.never++ }
            }
            $tester = Get-CleanIdentity $row.Tester
            if ($tester) { [void]$t.TesterSet.Add($tester) }
        }
    }

    $tasks = foreach ($t in $byTask.Values) {
        $d = $taskDates[$t.id]
        $changedIso = if ($d) { $d.changed } else { '' }
        $closedIso  = if ($d) { $d.closed }  else { '' }

        [pscustomobject][ordered]@{
            id       = $t.id
            title    = $t.title
            state    = $t.state
            kind     = $t.kind
            assignee = $t.assignee
            product  = $t.product
            pbiId    = $t.pbiId
            pbiTitle = $t.pbiTitle
            changedOn   = Get-DateOnly  $changedIso
            changedDays = Get-DaysSince $changedIso
            closedOn    = Get-DateOnly  $closedIso
            closedDays  = Get-DaysSince $closedIso
            # Deadline pair. NO_TARGET keeps daysLeft numeric so the column
            # stays sortable and filterable; the UI checks targetOn to decide
            # whether to show anything at all.
            startOn     = Get-DateOnly $t.pbiStart
            targetOn    = Get-DateOnly $t.pbiTarget
            # Bug IDs travel as a list, not just a count, so the browser can
            # de-duplicate at group level: one bug linked to two of a person's
            # PBIs must count once for that person, not twice.
            bugIds      = $(
                if ($t.pbiId -and $bugsByPbi.ContainsKey($t.pbiId)) {
                    (@($bugsByPbi[$t.pbiId]) | Sort-Object) -join ','
                } else { '' }
            )
            daysLeft    = $(
                $d = Get-DaysUntil $t.pbiTarget
                if ($null -eq $d) { 99999 } else { $d }
            )
            # Length of the whole start->target window, so the page can draw an
            # elapsed meter without parsing dates in the browser (and without
            # drifting off the frozen asOf). 0 means "cannot draw a meter".
            windowDays  = $(
                $s = Get-DaysUntil $t.pbiStart
                $e = Get-DaysUntil $t.pbiTarget
                if ($null -eq $s -or $null -eq $e) { 0 }
                elseif (($e - $s) -le 0) { 0 }
                else { [int]($e - $s) }
            )
            cases    = $t.CaseSet.Count
            testers  = ((@($t.TesterSet) | Sort-Object) -join ', ')
            exec     = $t.exec
            passed   = $t.passed
            failed   = $t.failed
            blocked  = $t.blocked
            na       = $t.na
            never    = $t.never
            mistakes = $t.mistakes
        }
    }
    $tasks = @($tasks)

    $execTotal = ($tasks | Measure-Object -Property exec -Sum).Sum
    Write-Host ("Aggregated to {0} tasks; {1} test-plan-sourced test points" -f $tasks.Count, $execTotal)

    # Built from char codes so this script file contains no non-ASCII bytes at
    # all. ConvertTo-Json emits them as \uXXXX escapes, so the generated HTML is
    # pure ASCII too and cannot be mangled by a codepage mismatch.
    $glyphs = [pscustomobject][ordered]@{
        check   = [string][char]0x2713   # heavy check
        dash    = [string][char]0x2013   # en dash
        blocked = [string][char]0x2298   # circled division slash
        cross   = [string][char]0x2717   # ballot X
        circle  = [string][char]0x25CB   # hollow circle
        mdash   = [string][char]0x2014   # em dash
        dot     = [string][char]0x00B7   # middot separator
        up      = [string][char]0x2191
        down    = [string][char]0x2193
    }

    $payload = [pscustomobject][ordered]@{
        meta = [pscustomobject][ordered]@{
            generated  = $AsOf.ToString('yyyy-MM-dd HH:mm')
            asOf       = $AsOf.ToString('yyyy-MM-dd')
            sourceRows = $connected.Count
            taskCount  = $tasks.Count
            execTotal  = [int]$execTotal
            glyphs     = $glyphs
            # Labels are here, not hardcoded in the page, because the fields
            # behind them are a workaround: ADO calls the target field
            # "Planned Hot Fix Release Date". If the source moves to Feature
            # StartDate/TargetDate, only these two strings and the two column
            # constants in 6-Build-WeeklyReports.ps1 change.
            startLabel  = 'Start'
            targetLabel = 'Target'
            targetNote  = 'Start = PBI created date. Target = the PBI field ADO labels "Planned Hot Fix Release Date", repurposed by the perf team as the delivery target.'
            withTarget  = [int](@($tasks | Where-Object { $_.targetOn }).Count)
            bugUnreadable = [int]$bugUnreadable
        }
        tasks = $tasks
    }

    $json = $payload | ConvertTo-Json -Depth 6 -Compress
    # Escape every '<' as its JSON unicode form. No '<' appears in JSON outside
    # string literals, so a blanket replace is safe, and it stops a task title
    # containing "</script>" from breaking out of the embedded data block.
    # Built from the char code so the literal backslash is unmistakable.
    $jsonLessThan = [string][char]0x5C + 'u003c'      # -> <
    $json = $json.Replace('<', $jsonLessThan)

    $template = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Performance Testing - Weekly Meeting Report</title>
<style>
  /* Light is the default; dark is a selected set of steps for the dark
     surface, not an automatic inversion. Declared under both the media query
     (OS setting) and the [data-theme] scope (explicit toggle) so the toggle
     wins either way. */
  :root {
    color-scheme: light;
    --surface:      #fcfcfb;
    --plane:        #f9f9f7;
    --ink:          #0b0b0b;
    --ink-2:        #52514e;
    --ink-muted:    #898781;
    --grid:         #e1e0d9;
    --axis:         #c3c2b7;
    --border:       rgba(11,11,11,0.10);
    /* Status scale - fixed, never themed by hue. */
    --st-good:      #0ca30c;
    --st-na:        #898781;
    --st-serious:   #ec835a;
    --st-critical:  #d03b3b;
    --st-never:     #e1e0d9;
    /* Task states are ORDERED stages (To Do -> In Progress -> Done), so they
       get a single-hue ordinal ramp, light->dark, NOT the status colours and
       NOT a categorical set. Validated: monotone lightness, adjacent dL >=
       0.06, light end clears the surface. */
    --ts-todo:      #86b6ef;
    --ts-doing:     #3987e5;
    --ts-done:      #1c5cab;
    /* Elapsed-meter levels. Classic traffic light, chosen over green/orange
       after validation: green vs the orange status step measures dE 5.6 under
       protanopia (fail), while green vs amber measures 11.3 (pass) and 27.6
       for normal vision. Green vs red is still only 4.1 under deuteranopia,
       which is why the meter always carries its day label. */
    --mtr-ok:       #0ca30c;
    --mtr-soon:     #fab219;
    --mtr-late:     #d03b3b;
  }
  @media (prefers-color-scheme: dark) {
    :root:where(:not([data-theme="light"])) {
      color-scheme: dark;
      --surface:   #1a1a19;
      --plane:     #0d0d0d;
      --ink:       #ffffff;
      --ink-2:     #c3c2b7;
      --ink-muted: #898781;
      --grid:      #2c2c2a;
      --axis:      #383835;
      --border:    rgba(255,255,255,0.10);
      --st-never:  #2c2c2a;
      --ts-todo:   #9ec5f4;
      --ts-doing:  #5598e7;
      --ts-done:   #256abf;
    }
  }
  :root[data-theme="dark"] {
    color-scheme: dark;
    --surface:   #1a1a19;
    --plane:     #0d0d0d;
    --ink:       #ffffff;
    --ink-2:     #c3c2b7;
    --ink-muted: #898781;
    --grid:      #2c2c2a;
    --axis:      #383835;
    --border:    rgba(255,255,255,0.10);
    --st-never:  #2c2c2a;
    --ts-todo:   #9ec5f4;
    --ts-doing:  #5598e7;
    --ts-done:   #256abf;
  }

  * { box-sizing: border-box; }
  body {
    margin: 0; padding: 24px 28px 56px;
    background: var(--plane); color: var(--ink);
    font: 14px/1.5 system-ui, -apple-system, "Segoe UI", sans-serif;
  }
  h1 { font-size: 20px; font-weight: 600; margin: 0 0 4px; }
  h2 { font-size: 14px; font-weight: 600; margin: 0 0 2px; }
  .sub { color: var(--ink-2); font-size: 12.5px; margin: 0; }
  .muted { color: var(--ink-muted); }

  header { display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; margin-bottom: 18px; }
  button {
    font: inherit; color: var(--ink); background: var(--surface);
    border: 1px solid var(--border); border-radius: 6px;
    padding: 6px 11px; cursor: pointer; min-height: 32px;
  }
  button:hover { background: var(--plane); }
  button[aria-pressed="true"] { border-color: var(--ink-muted); font-weight: 600; }

  .card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 10px; padding: 16px 18px; margin-bottom: 16px;
  }

  /* One filter row, above everything it scopes.
     Two zones, side by side and never stacked: the controls wrap among
     themselves on the left, the scope readout is pinned right. Previously
     everything shared one wrapping flex, so the readout was just another item
     and dropped to a second line as soon as the controls filled the width. */
  .filters { display: flex; flex-wrap: nowrap; gap: 16px; align-items: flex-end; }
  .filter-controls {
    display: flex; flex-wrap: wrap; gap: 12px 14px; align-items: flex-end;
    flex: 1 1 auto; min-width: 0;
  }

  /* Pinned to the top so the controls stay reachable while reading rows far
     down the page - change a filter and watch the table update without
     scrolling back. z-index sits below #tip (50) so chart tooltips still draw
     over it, and above ordinary content so rows pass underneath.
     The solid surface is load-bearing: without it, scrolling rows would show
     through the bar. */
  .card.filters {
    position: sticky;
    top: 0;
    z-index: 30;
    background: var(--grid);
    border-color: var(--axis);
    box-shadow: 0 4px 20px rgba(0,0,0,.38);
  }
  /* Fills the body's top padding, so nothing is briefly visible above the bar
     as it pins. The pseudo-element always matches the page background, not the
     bar background, so it stays invisible as content scrolls under it. */
  .card.filters::before {
    content: ""; position: absolute; left: -1px; right: -1px;
    top: -26px; height: 26px; background: var(--plane);
  }
  .filter-controls > label { display: flex; flex-direction: column; gap: 4px; font-size: 12px; color: var(--ink-2); }
  select, input[type="search"] {
    font: inherit; color: var(--ink); background: var(--surface);
    border: 1px solid var(--border); border-radius: 6px;
    padding: 6px 8px; min-height: 32px; min-width: 150px;
  }
  /* Right-hand readout. Allowed to use two short lines rather than one long
     one - stacking inside its own column is what keeps it off a row of its own. */
  .scope {
    flex: 0 0 auto; font-size: 12px; color: var(--ink-2);
    text-align: right; line-height: 1.45; padding-bottom: 4px; white-space: nowrap;
  }
  /* Multi-select checkbox dropdown widget used by the filter bar.
     Chart-section selects (fGroup, fLoadGroup, fLoadColor) keep the base
     select style above and are not affected by any of these rules. */
  .multi-sel { position: relative; display: inline-block; }
  .multi-btn {
    font: inherit; color: var(--ink); background: var(--surface);
    border: 1px solid var(--border); border-radius: 6px;
    padding: 6px 28px 6px 8px; min-height: 32px; min-width: 150px; max-width: 172px;
    cursor: pointer; text-align: left; position: relative;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  }
  /* \25BE = small down-pointing triangle (ASCII-safe in a PS1 file). */
  .multi-btn::after {
    content: "\25BE"; position: absolute; right: 8px; top: 50%;
    transform: translateY(-50%); font-size: 10px; color: var(--ink-2); pointer-events: none;
  }
  .multi-panel {
    position: absolute; top: calc(100% + 4px); left: 0;
    background: var(--surface); border: 1px solid var(--border); border-radius: 8px;
    box-shadow: 0 4px 16px rgba(0,0,0,.18); z-index: 100;
    min-width: 160px; max-width: 220px; padding: 6px 0;
    max-height: 260px; overflow-y: auto;
  }
  .multi-panel[hidden] { display: none; }
  .cb-item {
    display: flex; flex-direction: row; align-items: center; gap: 8px;
    padding: 6px 12px; cursor: pointer; font-size: 13px; color: var(--ink);
    white-space: nowrap; text-align: left;
  }
  .cb-item:hover { background: var(--grid); }
  .cb-item input[type="checkbox"] { margin: 0; cursor: pointer; accent-color: var(--blue); }
  .cb-sep { height: 1px; background: var(--border); margin: 4px 8px; }
  /* Date-range sub-section inside the Activity panel. */
  .range-inputs { display: flex; flex-direction: column; gap: 6px; padding: 4px 12px 8px; }
  .range-inputs[hidden] { display: none; }
  .range-lbl { display: flex; flex-direction: row; align-items: center; gap: 6px; font-size: 12px; color: var(--ink-2); white-space: nowrap; }
  .range-date { font: inherit; font-size: 12px; color: var(--ink); background: var(--surface); border: 1px solid var(--border); border-radius: 6px; padding: 3px 5px; flex: 1; min-width: 0; }

  /* Hero + tiles. Proportional figures on big numbers - tabular-nums makes
     them look loose at display sizes. */
  .kpis { display: flex; flex-wrap: wrap; gap: 14px; align-items: stretch; }
  .hero { min-width: 268px; }   /* wide enough to keep the rate line unwrapped */
  .hero .value { font-size: 48px; font-weight: 600; line-height: 1.05; letter-spacing: -0.5px; }
  .hero .label { font-size: 12.5px; color: var(--ink-2); margin-top: 2px; }
  .hero .rate { font-size: 12.5px; color: var(--ink-2); margin-top: 8px; }
  .tile {
    flex: 1 1 132px; min-width: 132px;
    border-left: 3px solid var(--tile-color, var(--axis));
    padding-left: 11px;
  }
  .tile .value { font-size: 26px; font-weight: 600; line-height: 1.15; }
  .tile .label { font-size: 12.5px; color: var(--ink-2); display: flex; align-items: center; gap: 6px; }
  .tile .pct { font-size: 11.5px; color: var(--ink-muted); }
  /* Status colour never carries meaning alone - it always ships with this
     glyph and the text label beside it. */
  .glyph {
    width: 15px; height: 15px; flex: 0 0 15px; border-radius: 50%;
    display: inline-flex; align-items: center; justify-content: center;
    font-size: 10px; line-height: 1; font-weight: 700;
    background: var(--tile-color); color: #fff;
  }
  .glyph.on-light { color: var(--ink); }

  .banner {
    display: flex; gap: 10px; align-items: baseline;
    border-left: 3px solid var(--st-serious); padding: 10px 12px;
    background: var(--surface); border-radius: 6px; font-size: 13px;
  }

  /* Chart */
  .chart-head { display: flex; justify-content: space-between; align-items: center; gap: 12px; margin-bottom: 12px; }
  .legend { display: flex; flex-wrap: wrap; gap: 6px 16px; margin: 0 0 14px; padding: 0; list-style: none; font-size: 12.5px; color: var(--ink-2); }
  .legend li { display: flex; align-items: center; gap: 6px; }
  .swatch { width: 11px; height: 11px; border-radius: 2px; background: var(--sw); flex: 0 0 11px; }
  /* Header swatch: ties each count column back to the legend above, so the
     table view is readable as the same encoding as the bars. */
  th .hsw {
    display: inline-block; width: 8px; height: 8px; border-radius: 2px;
    background: var(--sw); margin-right: 5px; vertical-align: middle;
  }

  .bars { display: flex; flex-direction: column; gap: 12px; }
  .barrow { display: grid; grid-template-columns: 152px 1fr 92px; gap: 12px; align-items: center; }
  .barrow.wide { grid-template-columns: 152px 1fr 70px 78px; }
  .barrow:focus-visible { outline: 2px solid var(--ink-muted); outline-offset: 3px; border-radius: 4px; }
  .barname { font-size: 13px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  /* 2px surface gaps do the separating between segments - never a border. */
  .track { display: flex; gap: 2px; height: 20px; align-items: stretch; }
  .seg { background: var(--seg); position: relative; min-width: 2px; }
  .seg:first-child { border-radius: 2px 0 0 2px; }
  .seg:last-child  { border-radius: 0 4px 4px 0; }  /* rounded data-end, square at baseline */
  /* Enlarge the hover/hit area beyond the painted 20px. */
  .seg::after { content: ""; position: absolute; inset: -5px 0; }
  .barrow:hover .seg { filter: brightness(1.06); }
  .bartotal { font-size: 12.5px; color: var(--ink-2); font-variant-numeric: tabular-nums; }
  .scale { display: flex; justify-content: space-between; margin-top: 10px; padding-left: 164px; border-top: 1px solid var(--grid); }
  .scale span { font-size: 11px; color: var(--ink-muted); font-variant-numeric: tabular-nums; padding-top: 4px; }

  /* Tables. tabular-nums here, where digits must align vertically. */
  table { border-collapse: collapse; width: 100%; font-size: 13px; }
  th, td { text-align: left; padding: 7px 9px; border-bottom: 1px solid var(--grid); vertical-align: top; }
  th { font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.04em; color: var(--ink-muted); font-weight: 600; white-space: nowrap; }
  th.num, td.num { text-align: right; font-variant-numeric: tabular-nums; }
  /* Centred column. Kept separate from .num because a column can be numeric
     for sorting and filtering while not being right-aligned. */
  th.ctr, td.ctr { text-align: center; font-variant-numeric: tabular-nums; }
  th.sortable { cursor: pointer; user-select: none; }
  th.sortable:hover { color: var(--ink); }
  tbody tr:hover { background: var(--plane); }
  .tasktitle { font-weight: 500; }
  td.nowrap { white-space: nowrap; }

  /* Elapsed meter: how much of the start->target window is used up.
     Per the meter spec the unfilled track is a lighter step of the FILL's own
     hue (done here with an opacity wash of the same colour, so it adapts to
     both surfaces) - state then reads across the whole bar, not just the fill.
     The short label beside it is NOT decoration: green vs red measures dE 4.1
     under deuteranopia, so colour alone cannot carry this. Do not remove it. */
  /* Bar stacked ABOVE its label. Stacking keeps every bar's right edge on the
     same line no matter how long the label is ("done 113d early" is far wider
     than "55d"), which a side-by-side layout cannot do. */
  .meter { display: flex; flex-direction: column; align-items: center; gap: 3px; }
  .meter-track {
    position: relative; width: 76px; height: 6px; flex: 0 0 6px;
    border-radius: 3px; overflow: hidden;
  }
  /* The unfilled track must stay visible - it is what shows how much of the
     window is LEFT. At .22 it vanished on the dark surface, so a 60% bar and a
     100% bar looked the same length and the meter told you nothing without
     reading its label. */
  .meter-bg   { position: absolute; inset: 0; background: var(--mtr); opacity: .32; }
  .meter-fill { position: absolute; left: 0; top: 0; bottom: 0; background: var(--mtr); border-radius: 3px; }
  .meter-lab  { font-size: 11.5px; font-variant-numeric: tabular-nums; text-align: right; color: var(--ink-2); line-height: 1.25; }
  .meter-lab.over { color: var(--st-critical); font-weight: 600; }

  /* Per-column filter row. Sits under the sortable header, inside the same
     sticky-free thead; only the tbody re-renders as you type. */
  tr.filterrow th { padding: 4px 6px 8px; border-bottom: 1px solid var(--axis); }
  .colf {
    width: 100%; min-width: 0; box-sizing: border-box;
    font: inherit; font-size: 12px; color: var(--ink);
    background: var(--plane); border: 1px solid var(--border);
    border-radius: 5px; padding: 3px 6px; min-height: 26px;
  }
  .colf::placeholder { color: var(--ink-muted); }
  .colf:focus-visible { outline: 2px solid var(--ink-muted); outline-offset: 1px; }
  /* Narrow numeric filter boxes, or their min-width steals the Task column
     and titles wrap to three lines. Scoped to the task table so the chart
     tables keep their own sizing. */
  th.num .colf { text-align: right; width: 56px; }
  th.ctr .colf { text-align: center; width: 56px; }
  #taskTable th:first-child { min-width: 300px; }
  /* Give the meter columns real width. Without this the free space all goes to
     the Task/PBI column and the meter ends up squeezed against the number
     block - centred, but visibly cramped. */
  #taskTable th:nth-child(7), #taskTable td:nth-child(7),
  #loadTable th:nth-child(3), #loadTable td:nth-child(3) { min-width: 104px; }
  .colfoot { display: flex; align-items: center; gap: 12px; margin-top: 10px; }
  .colfoot .sub { margin: 0; }
  .colfoot button { margin-left: auto; }
  .meta-line { color: var(--ink-muted); font-size: 11.5px; margin-top: 2px; }
  .pill { display: inline-block; font-size: 11px; padding: 1px 7px; border-radius: 999px; border: 1px solid var(--border); color: var(--ink-2); white-space: nowrap; }
  .empty { padding: 28px 8px; text-align: center; color: var(--ink-muted); }
  .hidden { display: none !important; }

  /* Tooltip enhances; every value is also in a table, never gated behind hover. */
  #tip {
    position: fixed; z-index: 50; pointer-events: none; opacity: 0;
    transition: opacity .08s; background: var(--surface); color: var(--ink);
    border: 1px solid var(--border); border-radius: 8px; padding: 9px 11px;
    box-shadow: 0 6px 20px rgba(0,0,0,.16); font-size: 12.5px; min-width: 172px;
  }
  #tip .tt-head { font-weight: 600; margin-bottom: 6px; }
  #tip .tt-row { display: flex; align-items: center; gap: 8px; margin-top: 3px; }
  #tip .tt-key { width: 10px; height: 2px; border-radius: 1px; background: var(--sw); flex: 0 0 10px; }
  #tip .tt-val { margin-left: auto; font-weight: 600; font-variant-numeric: tabular-nums; }
  #tip .tt-name { color: var(--ink-2); }

  footer { margin-top: 26px; font-size: 12px; color: var(--ink-muted); line-height: 1.7; }

  @media print {
    body { background: #fff; padding: 0; }
    .card { break-inside: avoid; border-color: #ccc; }
    .filters, #themeBtn, #tableBtn, #tip { display: none !important; }
  }
</style>
</head>
<body>

<header>
  <div>
    <h1>Performance Testing - Weekly Meeting Report</h1>
    <p class="sub">Active work by person, and execution results from test plans.
       <span class="muted">Generated <span id="genAt"></span></span></p>
  </div>
  <button id="themeBtn" type="button" title="Switch light / dark">Theme</button>
</header>

<div class="card filters" role="group" aria-label="Filters">
 <div class="filter-controls">
  <label>Person
    <div class="multi-sel" id="fPerson">
      <button class="multi-btn" type="button" aria-haspopup="true" aria-expanded="false"><span class="multi-label">All people</span></button>
      <div class="multi-panel" hidden></div>
    </div>
  </label>
  <label>Product
    <div class="multi-sel" id="fProduct">
      <button class="multi-btn" type="button" aria-haspopup="true" aria-expanded="false"><span class="multi-label">All products</span></button>
      <div class="multi-panel" hidden></div>
    </div>
  </label>
  <label>Task state
    <div class="multi-sel" id="fState">
      <button class="multi-btn" type="button" aria-haspopup="true" aria-expanded="false"><span class="multi-label">3 selected</span></button>
      <div class="multi-panel" hidden>
        <label class="cb-item"><input type="checkbox" value="To Do" checked><span>To Do</span></label>
        <label class="cb-item"><input type="checkbox" value="In Progress" checked><span>In Progress</span></label>
        <label class="cb-item"><input type="checkbox" value="Done" checked><span>Done</span></label>
      </div>
    </div>
  </label>
  <label>Task kind
    <div class="multi-sel" id="fKind">
      <button class="multi-btn" type="button" aria-haspopup="true" aria-expanded="false"><span class="multi-label">All kinds</span></button>
      <div class="multi-panel" hidden></div>
    </div>
  </label>
  <label>Activity
    <div class="multi-sel" id="fActivity">
      <button class="multi-btn" type="button" aria-haspopup="true" aria-expanded="false"><span class="multi-label">Worked on, last 7 days</span></button>
      <div class="multi-panel" hidden>
        <label class="cb-item"><input type="checkbox" value="w7" checked><span>Worked on, last 7 days</span></label>
        <label class="cb-item"><input type="checkbox" value="w14"><span>Worked on, last 14 days</span></label>
        <label class="cb-item"><input type="checkbox" value="c7"><span>Completed, last 7 days</span></label>
        <label class="cb-item"><input type="checkbox" value="c30"><span>Completed, last 30 days</span></label>
        <div class="cb-sep"></div>
        <label class="cb-item"><input type="checkbox" value="range" id="cbRange"><span>Custom range</span></label>
        <div class="range-inputs" id="actRange" hidden>
          <label class="range-lbl">From <input type="date" id="rangeStart" class="range-date"></label>
          <label class="range-lbl">To&#160;&#160;&#160;<input type="date" id="rangeEnd" class="range-date"></label>
        </div>
      </div>
    </div>
  </label>
  <label>Search
    <input id="fText" type="search" placeholder="task, PBI, product...">
  </label>
  <button id="fReset" type="button">Reset</button>
 </div>
 <div class="scope" id="scopeNote"></div>
</div>

<div class="card">
  <h2>Activity <span class="muted" style="font-weight:400">- what moved, as of <span id="asOf"></span></span></h2>
  <p class="sub" style="margin-bottom:14px">All filters apply, including task state.</p>
  <div class="kpis" id="activity"></div>
</div>

<div class="card">
  <h2>Execution results <span class="muted" style="font-weight:400">- test points that came from a test plan</span></h2>
  <p class="sub" style="margin-bottom:14px">Scripting tasks are excluded here: they link test cases directly, which would double-count the same results.</p>
  <div class="kpis" id="kpis"></div>
</div>

<div id="mistakeBanner" class="banner hidden" style="margin-bottom:16px">
  <strong id="mistakeCount"></strong>
  <span>link(s) in scope use a <code>Child</code> relationship where <code>Tests</code> is required, so they are missing from ADO traceability. See <code>exceptions_weekly.csv</code>.</span>
</div>

<div class="card">
  <div class="chart-head">
    <div>
      <h2 id="chartTitle">Test Points</h2>
      <p class="sub">Bar length is volume; segments are the outcome mix.</p>
    </div>
    <div style="display:flex;gap:8px;align-items:center">
      <label style="font-size:12px;color:var(--ink-2)">Group by
        <select id="fGroup" style="min-width:120px">
          <option value="assignee" selected>Person</option>
          <option value="product">Product</option>
          <option value="state">Task state</option>
        </select>
      </label>
      <button id="tableBtn" type="button" aria-pressed="false">Table view</button>
    </div>
  </div>
  <ul class="legend" id="legend"></ul>
  <div id="chartWrap"><div class="bars" id="bars"></div><div class="scale" id="scale"></div></div>
  <div id="chartTable" class="hidden"></div>
</div>

<div class="card">
  <div class="chart-head">
    <div>
      <h2 id="loadTitle">Assigned Work Items</h2>
      <p class="sub" id="loadSub">How many tasks each person holds, split by state, and how many PBIs they span.</p>
    </div>
    <div style="display:flex;gap:8px;align-items:center">
      <label style="font-size:12px;color:var(--ink-2)">Group by
        <select id="fLoadGroup" style="min-width:110px">
          <option value="assignee" selected>Person</option>
          <option value="product">Product</option>
        </select>
      </label>
      <label style="font-size:12px;color:var(--ink-2)">Colour by
        <select id="fLoadColor" style="min-width:110px">
          <option value="state" selected>Task state</option>
          <option value="urgency">Time to target</option>
        </select>
      </label>
      <button id="loadTableBtn" type="button" aria-pressed="false">Table view</button>
    </div>
  </div>
  <ul class="legend" id="loadLegend"></ul>
  <div id="loadWrap"><div class="bars" id="loadBars"></div><div class="scale" id="loadScale"></div></div>
  <div id="loadTable" class="hidden"></div>
</div>

<div class="card">
  <h2>Task Details</h2>
  <p class="sub" style="margin-bottom:12px">Outcome columns cover test-plan-sourced points, so they total to the tiles above.
     Scripting tasks show <span class="muted">-</span> because they author test cases rather than execute a plan.</p>
  <div id="taskTable"></div>
</div>

<footer>
  <div id="provenance"></div>
  <div>Rebuild: <code>Run-AdoExtracts.bat</code> (extract + reports + this page), or
       <code>powershell -File 7-Build-MeetingReport.ps1</code> to regenerate just this page from the existing CSVs.</div>
</footer>

<div id="tip" role="tooltip" aria-hidden="true"></div>

<script id="payload" type="application/json">/*__DATA__*/</script>
<script>
(function () {
  "use strict";

  var DATA  = JSON.parse(document.getElementById("payload").textContent);
  var TASKS = DATA.tasks || [];
  var META  = DATA.meta || {};

  // Status scale. Order is deliberate: Passed must NOT sit beside Failed -
  // that red/green pair is indistinguishable under deuteranopia. Every entry
  // carries a glyph and a label so hue is never the only channel.
  // Glyphs arrive through the JSON payload rather than being written literally
  // here, ON PURPOSE. This .ps1 has no UTF-8 BOM, so PowerShell 5.1 reads it as
  // Windows-1252 and a literal checkmark in the source gets double-encoded on
  // write, rendering as mojibake. The generator builds them from char codes and
  // ConvertTo-Json escapes them to \uXXXX, so this file stays pure ASCII and is
  // immune to however it happens to be saved.
  var GL = META.glyphs || {};

  var OUTCOMES = [
    { key: "passed",  label: "Passed",         glyph: GL.check,   varName: "--st-good"     },
    { key: "na",      label: "Not applicable", glyph: GL.dash,    varName: "--st-na"       },
    { key: "blocked", label: "Blocked",        glyph: GL.blocked, varName: "--st-serious"  },
    { key: "failed",  label: "Failed",         glyph: GL.cross,   varName: "--st-critical" },
    { key: "never",   label: "Not started",    glyph: GL.circle,  varName: "--st-never"    }
  ];
  var SEP = "  " + GL.dot + "  ";
  var LIGHT_GLYPH = { never: true, na: false };

  // Precomputed so the Run rate column is sortable like any other number.
  // -1 parks outcome-less Scripting tasks at the bottom instead of pretending 0%.
  TASKS.forEach(function (t) {
    var run = t.exec - t.never;
    t.runRate  = t.exec ? run / t.exec : -1;
    t.passRate = run > 0 ? t.passed / run : -1;   // -1 sorts "nothing run" last
  });

  var $ = function (id) { return document.getElementById(id); };
  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text !== undefined && text !== null) n.textContent = String(text);
    return n;
  }
  function fmt(n) { return (n || 0).toLocaleString("en-US"); }
  function rate(num, den) { return den > 0 ? (Math.round(num / den * 1000) / 10).toFixed(1) + "%" : "n/a"; }

  // ---- filter plumbing -----------------------------------------------------
  function uniq(key) {
    var seen = Object.create(null), out = [];
    TASKS.forEach(function (t) {
      var v = t[key];
      if (v && !seen[v]) { seen[v] = 1; out.push(v); }
    });
    return out.sort();
  }
  // Multi-select helpers ---------------------------------------------------
  // Returns a Set of checked values from a .multi-panel element.
  function setFromPanel(panelEl) {
    var s = new Set();
    panelEl.querySelectorAll("input[type=checkbox]:checked").forEach(function (cb) { s.add(cb.value); });
    return s;
  }
  // Returns the button label text: joined values if short, "N selected" if long.
  function labelFromSet(s, allLabel) {
    if (!s.size) return allLabel;
    var arr = []; s.forEach(function (v) { arr.push(v); });
    var text = arr.join(", ");
    return text.length > 26 ? arr.length + " selected" : text;
  }
  // Toggles the named panel; closes all others first.
  function openPanel(containerId) {
    var cont = $(containerId), panel = cont.querySelector(".multi-panel");
    var willOpen = panel.hidden;
    document.querySelectorAll(".multi-panel").forEach(function (p) {
      p.hidden = true;
      p.closest(".multi-sel").querySelector(".multi-btn").setAttribute("aria-expanded", "false");
    });
    if (willOpen) {
      panel.hidden = false;
      cont.querySelector(".multi-btn").setAttribute("aria-expanded", "true");
    }
  }
  // Populates a panel with checkbox rows from a values array.
  // defaultSet: Set of values that start checked (pass new Set() for none).
  function fillCheckboxes(containerId, values, defaultSet) {
    var panel = $(containerId).querySelector(".multi-panel");
    values.forEach(function (v) {
      var lbl = document.createElement("label");
      lbl.className = "cb-item";
      var cb = document.createElement("input");
      cb.type = "checkbox"; cb.value = v;   // textContent on span: labels are untrusted data
      if (defaultSet && defaultSet.has(v)) cb.checked = true;
      var sp = document.createElement("span");
      sp.textContent = v;
      lbl.appendChild(cb); lbl.appendChild(sp);
      panel.appendChild(lbl);
    });
  }
  // Wires a dynamically-populated panel to update state[stateKey] on change.
  function wirePanel(id, stateKey, allLabel) {
    var cont = $(id), panel = cont.querySelector(".multi-panel");
    cont.querySelector(".multi-btn").addEventListener("click", function (e) {
      e.stopPropagation(); openPanel(id);
    });
    panel.addEventListener("change", function () {
      state[stateKey] = setFromPanel(panel);
      cont.querySelector(".multi-label").textContent = labelFromSet(state[stateKey], allLabel);
      render();
    });
  }

  var state = { person: new Set(), product: new Set(),
                state: new Set(["To Do", "In Progress", "Done"]),
                kind: new Set(), text: "", activity: new Set(["w7"]),
                activityRange: { start: "", end: "" },
                group: "assignee", tableView: false, sortKey: "exec", sortDir: -1,
                loadGroup: "assignee", loadColor: "state", loadTableView: false, colFilters: {} };

  // changedDays / closedDays are -1 when the date is absent, so a plain
  // "<= 7" test would wrongly match those. Always require >= 0 first.
  function withinDays(value, limit) { return value >= 0 && value <= limit; }

  var ACTIVITY = {
    w7:  function (t) { return withinDays(t.changedDays, 7); },
    w14: function (t) { return withinDays(t.changedDays, 14); },
    c7:  function (t) { return withinDays(t.closedDays, 7); },
    c30: function (t) { return withinDays(t.closedDays, 30); },
    // ISO string comparison works for YYYY-MM-DD dates (lexicographic = chronological).
    range: function (t) {
      var s = state.activityRange.start, e = state.activityRange.end;
      if (!s && !e) return false;
      var d = t.changedOn;
      if (!d) return false;
      if (s && d < s) return false;
      if (e && d > e) return false;
      return true;
    }
  };
  // Short display labels for the Activity button summary.
  var ACT_LABEL = { w7: "Last 7d", w14: "Last 14d", c7: "Closed 7d", c30: "Closed 30d", range: "Custom" };
  function activityLabel() {
    if (!state.activity.size) return "Any time";
    var arr = []; state.activity.forEach(function (a) { arr.push(ACT_LABEL[a] || a); });
    return arr.length > 2 ? arr.length + " selected" : arr.join(", ");
  }

  function visible(opts) {
    opts = opts || {};
    var q = state.text.trim().toLowerCase();
    return TASKS.filter(function (t) {
      if (state.person.size  && !state.person.has(t.assignee))  return false;
      if (state.product.size && !state.product.has(t.product))  return false;
      if (!opts.ignoreState  && state.state.size && !state.state.has(t.state)) return false;
      if (state.kind.size    && !state.kind.has(t.kind))         return false;
      if (state.activity.size) {
        var actPass = false;
        state.activity.forEach(function (a) { if (ACTIVITY[a] && ACTIVITY[a](t)) actPass = true; });
        if (!actPass) return false;
      }
      if (q) {
        var hay = (t.title + " " + t.pbiTitle + " " + t.product + " " + t.id + " " + t.pbiId + " " + t.assignee).toLowerCase();
        if (hay.indexOf(q) === -1) return false;
      }
      return true;
    });
  }
  function totals(rows) {
    var acc = { exec: 0, cases: 0, mistakes: 0 };
    OUTCOMES.forEach(function (o) { acc[o.key] = 0; });
    rows.forEach(function (t) {
      acc.exec += t.exec; acc.cases += t.cases; acc.mistakes += t.mistakes;
      OUTCOMES.forEach(function (o) { acc[o.key] += t[o.key] || 0; });
    });
    return acc;
  }

  // ---- KPI row -------------------------------------------------------------
  function renderKpis(rows, sum) {
    var wrap = $("kpis");
    wrap.textContent = "";

    // The hero is PASS rate, not run rate. Run rate is a coverage measure -
    // it is dragged down by tests nobody has executed yet (20 suites here were
    // never started at all), which says nothing about quality and reads as
    // alarming in a management deck. Pass rate already excludes never-run by
    // construction. Run rate is kept as context below, not hidden.
    var run = sum.exec - sum.never;
    var hero = el("div", "tile hero");
    hero.style.borderLeftColor = "var(--st-good)";
    hero.appendChild(el("div", "value", rate(sum.passed, run)));
    hero.appendChild(el("div", "label", "Pass rate, of tests actually run"));
    // Two explicit lines rather than one that wraps unpredictably at the tile's
    // width - an accidental orphan looks sloppy on a page that goes in front
    // of management.
    hero.appendChild(el("div", "rate", fmt(sum.passed) + " passed of " + fmt(run) + " run"));
    var r2 = el("div", "rate");
    r2.style.marginTop = "2px";
    r2.textContent = fmt(sum.exec) + " planned " + GL.dot + " " + rate(run, sum.exec) + " executed";
    hero.appendChild(r2);
    wrap.appendChild(hero);

    OUTCOMES.forEach(function (o) {
      var tile = el("div", "tile");
      tile.style.setProperty("--tile-color", "var(" + o.varName + ")");
      var lab = el("div", "label");
      var g = el("span", "glyph" + (LIGHT_GLYPH[o.key] ? " on-light" : ""), o.glyph);
      g.setAttribute("aria-hidden", "true");
      lab.appendChild(g);
      lab.appendChild(el("span", null, o.label));
      tile.appendChild(el("div", "value", fmt(sum[o.key])));
      tile.appendChild(lab);
      tile.appendChild(el("div", "pct", sum.exec ? rate(sum[o.key], sum.exec) + " of total" : GL.mdash));
      wrap.appendChild(tile);
    });
  }

  // ---- activity tiles ------------------------------------------------------
  function renderActivity() {
    var rows = visible();
    var wrap = $("activity");
    wrap.textContent = "";

    function count(fn) { var n = 0; rows.forEach(function (t) { if (fn(t)) n++; }); return n; }

    var tiles = [
      { label: "Worked on, last 7 days",  value: count(ACTIVITY.w7),
        sub: count(ACTIVITY.w14) + " in last 14", color: "var(--st-good)" },
      { label: "Completed, last 7 days",  value: count(ACTIVITY.c7),
        sub: count(ACTIVITY.c30) + " in last 30", color: "var(--axis)" },
      { label: "No change in 30+ days",   value: count(function (t) { return t.changedDays > 30; }),
        sub: "of " + fmt(rows.length) + " tasks", color: "var(--st-serious)" },
      // Only shown when it happens. Every task currently has a state-change
      // date, so a permanent "0" tile would just be noise in a meeting - but
      // if an extract ever comes back without them, this surfaces it.
      { label: "Never touched",           value: count(function (t) { return t.changedDays < 0; }),
        sub: "no state-change date", color: "var(--st-na)", hideWhenZero: true }
    ];

    // Deadline tiles only appear once somebody has actually set a target.
    // Showing "0 overdue" when NOTHING has a deadline would read as good news
    // when it really means "no data".
    var haveTarget = rows.filter(function (t) { return t.targetOn; });
    if (haveTarget.length) {
      // Done tasks are excluded from both. A completed task cannot be
      // "overdue" - it was finished late, which is history, not a risk. The
      // rest of this card deliberately spans all states; these two do not.
      var live = function (t) { return t.targetOn && t.state !== "Done"; };
      var overdue = count(function (t) { return live(t) && t.daysLeft < 0; });
      var soon    = count(function (t) { return live(t) && t.daysLeft >= 0 && t.daysLeft <= 7; });
      tiles.push({ label: "Overdue", value: overdue,
                   sub: "past " + (META.targetLabel || "target"),
                   color: "var(--st-critical)", hideWhenZero: true });
      tiles.push({ label: "Due within 7 days", value: soon,
                   sub: fmt(haveTarget.length) + " of " + fmt(rows.length) + " have a target",
                   color: "var(--st-serious)" });
    } else {
      tiles.push({ label: "No target dates set", value: 0,
                   sub: "nothing to track against", color: "var(--st-na)", isNote: true });
    }

    tiles.filter(function (t) { return !(t.hideWhenZero && t.value === 0); }).forEach(function (t) {
      var tile = el("div", "tile");
      tile.style.setProperty("--tile-color", t.color);
      tile.appendChild(el("div", "value", t.isNote ? GL.mdash : fmt(t.value)));
      tile.appendChild(el("div", "label", t.label));
      tile.appendChild(el("div", "pct", t.sub));
      wrap.appendChild(tile);
    });
  }

  // ---- legend --------------------------------------------------------------
  function renderLegend() {
    var ul = $("legend");
    ul.textContent = "";
    OUTCOMES.forEach(function (o) {
      var li = el("li");
      var sw = el("span", "swatch");
      sw.style.setProperty("--sw", "var(" + o.varName + ")");
      li.appendChild(sw);
      li.appendChild(el("span", null, o.glyph + "  " + o.label));
      ul.appendChild(li);
    });
  }

  // ---- grouped stacked bars ------------------------------------------------
  function groupRows(rows) {
    var map = Object.create(null), order = [];
    rows.forEach(function (t) {
      var k = t[state.group] || "(none)";
      if (!map[k]) { map[k] = { name: k, exec: 0 }; OUTCOMES.forEach(function (o) { map[k][o.key] = 0; }); order.push(k); }
      map[k].exec += t.exec;
      OUTCOMES.forEach(function (o) { map[k][o.key] += t[o.key] || 0; });
    });
    return order.map(function (k) { return map[k]; })
                .filter(function (g) { return g.exec > 0; })
                .sort(function (a, b) { return b.exec - a.exec; });
  }

  function renderBars(rows) {
    var groups = groupRows(rows);
    var bars = $("bars"), scale = $("scale");
    bars.textContent = ""; scale.textContent = "";

    if (!groups.length) {
      bars.appendChild(el("div", "empty", "No test-plan-sourced results in this slice."));
      return groups;
    }
    var max = groups[0].exec;

    groups.forEach(function (g) {
      var row = el("div", "barrow");
      row.tabIndex = 0;
      row.appendChild(el("div", "barname", g.name));

      var track = el("div", "track");
      track.style.width = (g.exec / max * 100).toFixed(2) + "%";
      OUTCOMES.forEach(function (o) {
        var v = g[o.key];
        if (!v) return;                       // never render a 0-width segment
        var seg = el("div", "seg");
        seg.style.setProperty("--seg", "var(" + o.varName + ")");
        seg.style.width = (v / g.exec * 100).toFixed(3) + "%";
        track.appendChild(seg);
      });
      row.appendChild(track);
      // Only the total is direct-labelled; per-segment values live in the
      // tooltip and the table view. A number on every segment is noise.
      row.appendChild(el("div", "bartotal", fmt(g.exec)));

      bindTip(row, g);
      bars.appendChild(row);
    });

    [0, Math.round(max / 2), max].forEach(function (v) { scale.appendChild(el("span", null, fmt(v))); });
    return groups;
  }

  // ---- tooltip (hover AND keyboard focus show the same thing) -------------
  var tip = $("tip");
  function tipHtml(g) {
    tip.textContent = "";
    tip.appendChild(el("div", "tt-head", g.name));
    OUTCOMES.forEach(function (o) {
      var r = el("div", "tt-row");
      var k = el("span", "tt-key");
      k.style.setProperty("--sw", "var(" + o.varName + ")");
      r.appendChild(k);
      r.appendChild(el("span", "tt-name", o.label));
      r.appendChild(el("span", "tt-val", fmt(g[o.key])));
      tip.appendChild(r);
    });
    var tot = el("div", "tt-row");
    tot.style.marginTop = "6px";
    tot.appendChild(el("span", "tt-name", "Total"));
    tot.appendChild(el("span", "tt-val", fmt(g.exec)));
    tip.appendChild(tot);
  }
  function place(x, y) {
    var pad = 14, r = tip.getBoundingClientRect();
    var left = Math.min(x + pad, window.innerWidth - r.width - 8);
    var top = Math.min(y + pad, window.innerHeight - r.height - 8);
    tip.style.left = Math.max(8, left) + "px";
    tip.style.top = Math.max(8, top) + "px";
  }
  function show(g, x, y) { tipHtml(g); tip.style.opacity = "1"; tip.setAttribute("aria-hidden", "false"); place(x, y); }
  function hide() { tip.style.opacity = "0"; tip.setAttribute("aria-hidden", "true"); }
  function bindTip(row, g) {
    row.addEventListener("pointermove", function (e) { show(g, e.clientX, e.clientY); });
    row.addEventListener("pointerleave", hide);
    row.addEventListener("focus", function () {
      var r = row.getBoundingClientRect();
      show(g, r.left + r.width / 2, r.bottom);
    });
    row.addEventListener("blur", hide);
  }

  // ---- chart's table twin --------------------------------------------------
  function renderChartTable(groups) {
    var host = $("chartTable");
    host.textContent = "";
    var t = el("table");
    var thead = el("thead"), hr = el("tr");
    hr.appendChild(el("th", null, state.group === "assignee" ? "Person" : (state.group === "product" ? "Product" : "Task state")));
    OUTCOMES.forEach(function (o) { hr.appendChild(el("th", "num", o.label)); });
    hr.appendChild(el("th", "num", "Total"));
    thead.appendChild(hr); t.appendChild(thead);

    var tb = el("tbody");
    groups.forEach(function (g) {
      var tr = el("tr");
      tr.appendChild(el("td", null, g.name));
      OUTCOMES.forEach(function (o) { tr.appendChild(el("td", "num", fmt(g[o.key]))); });
      tr.appendChild(el("td", "num", fmt(g.exec)));
      tb.appendChild(tr);
    });
    t.appendChild(tb);
    host.appendChild(t);
  }

  // ---- Assigned Work Items: count of tasks per person, split by state -----
  // Task states are ordered stages, so they use the ordinal ramp (light ->
  // dark = earlier -> later), never the status colours. A state is not a
  // verdict; painting "Done" green would imply a quality judgement.
  var STATES = [
    { key: "To Do",       varName: "--ts-todo"  },
    { key: "In Progress", varName: "--ts-doing" },
    { key: "Done",        varName: "--ts-done"  }
  ];

  // Urgency: the SAME bars, coloured by time-to-target instead of workflow
  // state. Deliberately four buckets, not five - an "overdue / <=7d / <=30d /
  // later / none" scale put amber next to orange, which measures deltaE 13.6
  // for normal vision (below the 15 floor). Merging the middle buckets fixes
  // it at the source rather than mitigating an unreadable pair.
  //
  // These are status colours because the buckets ARE states of concern, so
  // each ships a glyph and a text label - hue never carries it alone.
  var URGENCY = [
    { key: "overdue", label: "Overdue",      glyph: GL.cross,   varName: "--st-critical",
      test: function (t) { return t.targetOn && t.daysLeft < 0; } },
    { key: "soon",    label: "Due in 7 days", glyph: GL.blocked, varName: "--st-serious",
      test: function (t) { return t.targetOn && t.daysLeft >= 0 && t.daysLeft <= 7; } },
    { key: "later",   label: "Later",        glyph: GL.dash,    varName: "--st-na",
      test: function (t) { return t.targetOn && t.daysLeft > 7; } },
    { key: "none",    label: "No target",    glyph: GL.circle,  varName: "--st-never",
      test: function ()  { return true; } }          // catch-all, must stay last
  ];

  // ---- elapsed meter -------------------------------------------------------
  // Draws "how much of the start->target window is gone". Three levels:
  //   green  < 75% elapsed        amber 75-100%        red past target
  // A Done item is neutral grey regardless of dates - finished work cannot be
  // at risk, and painting it red would put false alarm in the meeting.
  function meterCell(item, isDone) {
    // "ctr" must match the header cell's class, or the header and the meters
    // sit on different alignments - invisible while the cell holds a flex
    // meter, but wrong the moment anyone puts plain text here.
    var td = el("td", "ctr nowrap");
    if (!item.targetOn) {
      td.className += " muted";
      td.textContent = GL.mdash;
      return td;
    }

    var win = item.windowDays || 0;
    var colour, over = false, days, pct, label;

    if (isDone) {
      // FREEZE at completion. daysLeft is measured from today, so leaving a
      // finished item on it would keep the meter ticking down for months and
      // eventually flip a long-since-delivered PBI to "50d over".
      //   target - closed = (target - today) + (today - closed)
      //                   = daysLeft + closedDays
      // Positive = finished early, negative = finished late.
      var atClose = (item.closedDays !== undefined && item.closedDays !== null && item.closedDays >= 0)
                  ? item.daysLeft + item.closedDays
                  : null;
      colour = "var(--st-na)";                       // history, never alarming
      if (atClose === null) {
        pct = 1; label = "done";                     // done, close date unknown
      } else {
        pct = win > 0 ? (win - atClose) / win : 1;
        label = atClose > 0 ? "done " + atClose + "d early"
              : atClose === 0 ? "done on time"
              : "done " + Math.abs(atClose) + "d late";
      }
    } else {
      days = item.daysLeft;
      pct  = win > 0 ? (win - days) / win : (days < 0 ? 1.2 : 1);
      if (days < 0)         { colour = "var(--mtr-late)"; over = true; }
      else if (pct >= 0.75) { colour = "var(--mtr-soon)"; }
      else                  { colour = "var(--mtr-ok)"; }
      label = days < 0 ? Math.abs(days) + "d over" : days === 0 ? "today" : days + "d";
    }

    var wrap = el("div", "meter");
    var track = el("div", "meter-track");
    track.style.setProperty("--mtr", colour);
    track.appendChild(el("div", "meter-bg"));
    var fill = el("div", "meter-fill");
    fill.style.width = (Math.max(0, Math.min(1, pct)) * 100).toFixed(1) + "%";
    track.appendChild(fill);
    wrap.appendChild(track);

    var lab = el("span", "meter-lab" + (over ? " over" : ""), label);
    wrap.appendChild(lab);

    // The dates the meter is built from stay reachable without hovering a
    // tooltip - native title works on touch-less keyboard focus too.
    td.title = (item.startOn ? item.startOn : "?") + "  ->  " + item.targetOn +
               (win > 0 ? "   (" + win + " day window, " + Math.round(pct * 100) + "% elapsed)" : "") +
               (isDone && item.closedOn ? "   closed " + item.closedOn : "");
    td.appendChild(wrap);
    return td;
  }

  function urgencyOf(t) {
    for (var i = 0; i < URGENCY.length; i++) { if (URGENCY[i].test(t)) return URGENCY[i].key; }
    return "none";
  }

  // One place that decides what the load bars are segmented by.
  function loadDims() {
    if (state.loadColor === "urgency") {
      return URGENCY.map(function (u) {
        return { field: "u_" + u.key, label: u.label, varName: u.varName, glyph: u.glyph };
      });
    }
    return STATES.map(function (s) {
      return { field: s.key, label: s.key, varName: s.varName, glyph: null };
    });
  }

  function renderLoadLegend() {
    var ul = $("loadLegend");
    ul.textContent = "";
    loadDims().forEach(function (d) {
      var li = el("li");
      var sw = el("span", "swatch");
      sw.style.setProperty("--sw", "var(" + d.varName + ")");
      li.appendChild(sw);
      li.appendChild(el("span", null, (d.glyph ? d.glyph + "  " : "") + d.label));
      ul.appendChild(li);
    });
  }

  function loadGroups(rows) {
    var map = Object.create(null), order = [];
    rows.forEach(function (t) {
      var k = t[state.loadGroup] || "(none)";
      if (!map[k]) {
        map[k] = { name: k, total: 0, pbis: Object.create(null), soonest: null, soonestOn: "",
                   bugSet: Object.create(null) };
        STATES.forEach(function (s) { map[k][s.key] = 0; });
        URGENCY.forEach(function (u) { map[k]["u_" + u.key] = 0; });
        order.push(k);
      }
      var g = map[k];
      g.total++;
      if (g[t.state] === undefined) g[t.state] = 0;   // tolerate an unseen state
      g[t.state]++;
      g["u_" + urgencyOf(t)]++;
      // Union, not sum: the same bug linked to two of this person's PBIs is
      // one bug for them.
      if (t.bugIds) { t.bugIds.split(",").forEach(function (b) { if (b) g.bugSet[b] = 1; }); }
      // Soonest live deadline in this group. Done tasks are skipped - a
      // finished task's date is history, not a thing to plan around.
      if (t.targetOn && t.state !== "Done" && (g.soonest === null || t.daysLeft < g.soonest)) {
        g.soonest = t.daysLeft; g.soonestOn = t.targetOn;
      }
      var pk = t.pbiId || "(no PBI parent)";
      if (!g.pbis[pk]) {
        g.pbis[pk] = { id: t.pbiId, title: t.pbiTitle || "(no PBI parent)", tasks: 0,
                       targetOn: t.targetOn, daysLeft: t.daysLeft, startOn: t.startOn,
                       windowDays: t.windowDays, closedDays: null, closedOn: "",
                       // Bugs hang off the PBI, so every task of a PBI reports
                       // the same set - take it, don't accumulate it.
                       bugs: t.bugIds ? t.bugIds.split(",").filter(Boolean).length : 0 };
        STATES.forEach(function (s) { g.pbis[pk][s.key] = 0; });
        URGENCY.forEach(function (u) { g.pbis[pk]["u_" + u.key] = 0; });
      }
      g.pbis[pk].tasks++;
      if (g.pbis[pk][t.state] === undefined) g.pbis[pk][t.state] = 0;
      g.pbis[pk][t.state]++;
      g.pbis[pk]["u_" + urgencyOf(t)]++;
      // A PBI finishes when its LAST task does, i.e. the most recent close =
      // the SMALLEST closedDays (fewest days ago).
      if (t.closedDays >= 0 && (g.pbis[pk].closedDays === null || t.closedDays < g.pbis[pk].closedDays)) {
        g.pbis[pk].closedDays = t.closedDays;
        g.pbis[pk].closedOn   = t.closedOn;
      }
    });
    return order.map(function (k) { return map[k]; })
                .sort(function (a, b) { return b.total - a.total; });
  }

  function renderLoad(rows) {
    var groups = loadGroups(rows);
    var bars = $("loadBars"), scale = $("loadScale");
    bars.textContent = ""; scale.textContent = "";
    if (!groups.length) {
      bars.appendChild(el("div", "empty", "No tasks match these filters."));
      return groups;
    }
    var max = groups[0].total;

    groups.forEach(function (g) {
      var row = el("div", "barrow wide");
      row.tabIndex = 0;
      row.appendChild(el("div", "barname", g.name));

      var track = el("div", "track");
      track.style.width = (g.total / max * 100).toFixed(2) + "%";
      loadDims().forEach(function (d) {
        var v = g[d.field];
        if (!v) return;
        var seg = el("div", "seg");
        seg.style.setProperty("--seg", "var(" + d.varName + ")");
        seg.style.width = (v / g.total * 100).toFixed(3) + "%";
        track.appendChild(seg);
      });
      row.appendChild(track);
      row.appendChild(el("div", "bartotal", fmt(g.total) + " task" + (g.total === 1 ? "" : "s")));
      var np = Object.keys(g.pbis).length;
      row.appendChild(el("div", "bartotal muted", fmt(np) + " PBI" + (np === 1 ? "" : "s")));

      bindLoadTip(row, g);
      bars.appendChild(row);
    });

    [0, Math.round(max / 2), max].forEach(function (v) { scale.appendChild(el("span", null, fmt(v))); });
    return groups;
  }

  function loadTipHtml(g) {
    tip.textContent = "";
    tip.appendChild(el("div", "tt-head", g.name));
    loadDims().forEach(function (d) {
      var r = el("div", "tt-row");
      var k = el("span", "tt-key");
      k.style.setProperty("--sw", "var(" + d.varName + ")");
      r.appendChild(k);
      r.appendChild(el("span", "tt-name", d.label));
      r.appendChild(el("span", "tt-val", fmt(g[d.field])));
      tip.appendChild(r);
    });
    var tot = el("div", "tt-row"); tot.style.marginTop = "6px";
    tot.appendChild(el("span", "tt-name", "Tasks"));
    tot.appendChild(el("span", "tt-val", fmt(g.total)));
    tip.appendChild(tot);
    var pb = el("div", "tt-row");
    pb.appendChild(el("span", "tt-name", "Across PBIs"));
    pb.appendChild(el("span", "tt-val", fmt(Object.keys(g.pbis).length)));
    tip.appendChild(pb);
    // The single most actionable number for a person: their nearest live
    // deadline. Shown in both colour modes, not just the urgency one.
    if (g.soonest !== null) {
      var s = el("div", "tt-row");
      s.appendChild(el("span", "tt-name", "Soonest target"));
      s.appendChild(el("span", "tt-val",
        g.soonest < 0 ? Math.abs(g.soonest) + "d overdue"
        : g.soonest === 0 ? "today" : "in " + g.soonest + "d"));
      tip.appendChild(s);
      var s2 = el("div", "tt-row");
      s2.appendChild(el("span", "tt-name", g.soonestOn));
      tip.appendChild(s2);
    }
  }
  function bindLoadTip(row, g) {
    row.addEventListener("pointermove", function (e) { loadTipHtml(g); tip.style.opacity = "1"; tip.setAttribute("aria-hidden","false"); place(e.clientX, e.clientY); });
    row.addEventListener("pointerleave", hide);
    row.addEventListener("focus", function () {
      var r = row.getBoundingClientRect();
      loadTipHtml(g); tip.style.opacity = "1"; tip.setAttribute("aria-hidden","false");
      place(r.left + r.width / 2, r.bottom);
    });
    row.addEventListener("blur", hide);
  }

  // The table twin drills one level further than the chart, to the PBIs behind
  // each bar. There are ~34 PBIs, far past the ~7 that colour can carry, so
  // this is a table by design rather than more segments.
  function renderLoadTable(groups) {
    var host = $("loadTable");
    host.textContent = "";
    var t = el("table"), thead = el("thead"), hr = el("tr");
    // Target (index 2) is centred to match the meter cells meterCell() emits.
    [state.loadGroup === "assignee" ? "Person" : "Product", "PBI", "Target", "Tasks", "Bugs"].forEach(function (h, i) {
      hr.appendChild(el("th", i === 2 ? "ctr" : (i >= 3 ? "num" : null), h));
    });
    // Each count column carries the legend's own swatch, so the table view is
    // readable as the SAME encoding as the bars above it. Without this the
    // legend floats over a table with no colour in it and means nothing.
    loadDims().forEach(function (d) {
      var th = el("th", "num");
      var sw = el("span", "hsw");
      sw.style.setProperty("--sw", "var(" + d.varName + ")");
      sw.setAttribute("aria-hidden", "true");
      th.appendChild(sw);
      th.appendChild(document.createTextNode(d.label));
      hr.appendChild(th);
    });
    thead.appendChild(hr); t.appendChild(thead);

    var tb = el("tbody");
    groups.forEach(function (g) {
      var pbis = Object.keys(g.pbis).map(function (k) { return g.pbis[k]; })
                       .sort(function (a, b) { return b.tasks - a.tasks; });
      // Subtotal row per group, then one row per PBI beneath it.
      var sr = el("tr");
      var nameCell = el("td", null, g.name);
      nameCell.style.fontWeight = "600";
      sr.appendChild(nameCell);
      sr.appendChild(el("td", "muted", fmt(pbis.length) + " PBI" + (pbis.length === 1 ? "" : "s")));
      // Target stays BLANK on the group row. It is a PBI-level date, so putting
      // one PBI's deadline on a row that represents many would read as though
      // the whole group shared it. Blank rather than a dash, because a dash
      // means "no target set" on the PBI rows below. The soonest live deadline
      // is still available - it is in the bar's tooltip, labelled as such.
      sr.appendChild(el("td", "ctr"));
      var totCell = el("td", "num", fmt(g.total)); totCell.style.fontWeight = "600";
      sr.appendChild(totCell);
      var gBugs = Object.keys(g.bugSet).length;
      var gb = el("td", "num" + (gBugs ? "" : " muted"), fmt(gBugs));
      if (gBugs) gb.style.fontWeight = "600";
      sr.appendChild(gb);
      loadDims().forEach(function (d) { sr.appendChild(el("td", "num", fmt(g[d.field]))); });
      tb.appendChild(sr);

      pbis.forEach(function (p) {
        var tr = el("tr");
        tr.appendChild(el("td", null, ""));
        var c = el("td");
        c.appendChild(el("div", null, p.title));
        if (p.id) c.appendChild(el("div", "meta-line", "PBI " + p.id));
        tr.appendChild(c);
        // Target belongs on the PBI row - that is the level the date is set at.
        // A PBI counts as done only when every one of its tasks is.
        tr.appendChild(meterCell(p, p["Done"] === p.tasks));
        tr.appendChild(el("td", "num", fmt(p.tasks)));
        tr.appendChild(el("td", "num" + (p.bugs ? "" : " muted"), fmt(p.bugs)));
        // Real per-PBI counts. A zero is dimmed rather than dashed - a dash
        // reads as "no data", but 0 here is a genuine, known count.
        loadDims().forEach(function (d) {
          tr.appendChild(el("td", "num" + (p[d.field] ? "" : " muted"), fmt(p[d.field])));
        });
        tb.appendChild(tr);
      });
    });
    t.appendChild(tb);
    host.appendChild(t);
  }

  // ---- task table ----------------------------------------------------------
  // "Cases" is DISTINCT test case IDs; "Points" is test points from test plans.
  // Both are needed: a case sitting in several suites is several points, so
  // Passed can legitimately exceed Cases. Showing Points makes the outcome
  // columns visibly sum to something, which stops that looking like a bug.
  var COLS = [
    { key: "title",    label: "Task",      num: false },
    { key: "kind",     label: "Kind",      num: false, options: function () { return uniq("kind"); } },
    { key: "state",    label: "State",     num: false, options: function () { return uniq("state"); } },
    { key: "assignee", label: "Assignee",  num: false, options: function () { return uniq("assignee"); } },
    { key: "product",  label: "Product",   num: false, options: function () { return uniq("product"); } },
    { key: "changedDays", label: "Last change", num: true },
    // num:true keeps it sortable and filterable with >/< expressions;
    // align:"center" only changes where it sits in the column.
    { key: "daysLeft",    label: "Target",      num: true, align: "center" },
    { key: "cases",    label: "Cases",     num: true },
    { key: "exec",     label: "Points",    num: true },
    { key: "passed",   label: "Passed",    num: true },
    { key: "failed",   label: "Failed",    num: true },
    { key: "blocked",  label: "Blocked",   num: true },
    { key: "na",       label: "N/A",       num: true },
    { key: "never",    label: "Not started", num: true }
  ];

  // Column filters live on the table because the table is the drill-down
  // surface: the row at the top of the page scopes everything, these narrow
  // this table only. The header and filter row are built ONCE and only the
  // tbody re-renders - rebuilding an <input> the user is mid-keystroke in
  // would steal focus on every character typed.
  var taskUI = null, lastRows = [];

  function colValue(task, key) {
    // Sentinels must not be compared as numbers: passRate -1 means "nothing
    // run" and changedDays -1 means "no date", neither of which is a value.
    if (key === "passRate")    return task.passRate < 0 ? null : task.passRate * 100;
    if (key === "changedDays") return task.changedDays < 0 ? null : task.changedDays;
    if (key === "daysLeft")    return task.targetOn ? task.daysLeft : null;   // no target = no value
    return task[key];
  }

  function numMatch(value, expr) {
    expr = (expr || "").trim();
    if (!expr) return true;
    var m = /^(>=|<=|>|<|=)?\s*(-?\d+(?:\.\d+)?)$/.exec(expr);
    if (!m) return true;                                      // unparseable: don't filter
    if (value === null || value === undefined) return false;  // no value never matches
    var n = parseFloat(m[2]);
    switch (m[1] || "=") {
      case ">":  return value >  n;
      case ">=": return value >= n;
      case "<":  return value <  n;
      case "<=": return value <= n;
      default:   return value === n;
    }
  }

  function passesColFilters(task) {
    for (var i = 0; i < COLS.length; i++) {
      var c = COLS[i], f = state.colFilters[c.key];
      if (!f) continue;
      if (c.num) {
        if (!numMatch(colValue(task, c.key), f)) return false;
      } else {
        var v = String(task[c.key] === undefined || task[c.key] === null ? "" : task[c.key]);
        // The Task cell also renders the ID and PBI, so search those too -
        // otherwise typing a PBI number you can see would find nothing.
        if (c.key === "title") v += " " + task.id + " " + (task.pbiId || "") + " " + (task.pbiTitle || "");
        if (v.toLowerCase().indexOf(f.toLowerCase()) === -1) return false;
      }
    }
    return true;
  }

  function buildTaskUI() {
    var host = $("taskTable");
    host.textContent = "";
    var t = el("table"), thead = el("thead"), hr = el("tr"), labels = {};

    COLS.forEach(function (c) {
      var th = el("th", (c.align === "center" ? "ctr " : (c.num ? "num " : "")) + "sortable");
      var span = el("span", null, c.label);
      labels[c.key] = span;
      th.appendChild(span);
      th.tabIndex = 0;
      function doSort() {
        if (state.sortKey === c.key) state.sortDir = -state.sortDir;
        else { state.sortKey = c.key; state.sortDir = c.num ? -1 : 1; }
        renderTasks(lastRows);
      }
      th.addEventListener("click", doSort);
      th.addEventListener("keydown", function (e) { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); doSort(); } });
      hr.appendChild(th);
    });
    thead.appendChild(hr);

    var fr = el("tr", "filterrow");
    COLS.forEach(function (c) {
      var cell = el("th", c.align === "center" ? "ctr" : (c.num ? "num" : null)), input;
      if (c.options) {
        input = document.createElement("select");
        var blank = document.createElement("option");
        blank.value = ""; blank.textContent = "All";
        input.appendChild(blank);
        c.options().forEach(function (v) {
          var o = document.createElement("option");
          o.value = v; o.textContent = v;      // textContent: untrusted data
          input.appendChild(o);
        });
      } else {
        input = document.createElement("input");
        input.type = "text";
        input.placeholder = c.num ? ">0" : "contains";
      }
      input.className = "colf";
      input.setAttribute("aria-label", "Filter by " + c.label);
      function apply(e) { state.colFilters[c.key] = e.target.value; renderTasks(lastRows); }
      input.addEventListener("input", apply);
      input.addEventListener("change", apply);
      cell.appendChild(input);
      fr.appendChild(cell);
    });
    thead.appendChild(fr);
    t.appendChild(thead);

    var tb = el("tbody");
    t.appendChild(tb);
    host.appendChild(t);

    var foot = el("div", "colfoot");
    var note = el("span", "sub");
    var clear = el("button", null, "Clear column filters");
    clear.type = "button";
    clear.addEventListener("click", function () {
      state.colFilters = {};
      var fields = fr.querySelectorAll("input, select");
      for (var i = 0; i < fields.length; i++) fields[i].value = "";
      renderTasks(lastRows);
    });
    foot.appendChild(note);
    foot.appendChild(clear);
    host.appendChild(foot);

    taskUI = { tbody: tb, labels: labels, note: note };
  }

  function renderTasks(rows) {
    lastRows = rows;
    if (!taskUI) buildTaskUI();

    COLS.forEach(function (c) {
      taskUI.labels[c.key].textContent = c.label +
        (state.sortKey === c.key ? (state.sortDir === 1 ? " " + GL.up : " " + GL.down) : "");
    });

    var filtered = rows.filter(passesColFilters);
    var sorted = filtered.slice().sort(function (a, b) {
      var x = a[state.sortKey], y = b[state.sortKey];
      if (typeof x === "number" && typeof y === "number") return (x - y) * state.sortDir;
      return String(x).localeCompare(String(y)) * state.sortDir;
    });

    taskUI.note.textContent = filtered.length === rows.length
      ? fmt(rows.length) + " task" + (rows.length === 1 ? "" : "s")
      : "Showing " + fmt(filtered.length) + " of " + fmt(rows.length) + " tasks (column filters active)";

    var tb = taskUI.tbody;
    tb.textContent = "";
    if (!sorted.length) {
      var er = el("tr"), ec = el("td", "empty", "No tasks match these filters.");
      ec.colSpan = COLS.length;
      er.appendChild(ec); tb.appendChild(er);
      return;
    }

    sorted.forEach(function (task) {
      var tr = el("tr");

      var td = el("td");
      td.appendChild(el("div", "tasktitle", task.title));
      var meta = "#" + task.id;
      if (task.pbiId) meta += SEP + "PBI " + task.pbiId;
      else meta += SEP + "no PBI parent";
      // Start rides the meta line rather than taking a 15th column: it is
      // auto-derived (PBI created date) and lower-value than the target, but
      // it is what makes the target a window rather than a bare deadline.
      if (task.startOn) meta += SEP + (META.startLabel || "Start").toLowerCase() + " " + task.startOn;
      if (task.testers) meta += SEP + "tested by " + task.testers;
      td.appendChild(el("div", "meta-line", meta));
      tr.appendChild(td);

      var kindTd = el("td");
      kindTd.appendChild(el("span", "pill", task.kind));
      tr.appendChild(kindTd);

      var stTd = el("td");
      stTd.appendChild(el("span", "pill", task.state));
      tr.appendChild(stTd);

      tr.appendChild(el("td", "nowrap", task.assignee));
      tr.appendChild(el("td", "nowrap", task.product));

      var chg = el("td", "num nowrap");
      if (task.changedDays < 0) {
        chg.className += " muted";
        chg.textContent = GL.mdash;
      } else {
        chg.appendChild(el("div", null, task.changedOn));
        var age = task.changedDays === 0 ? "today"
                : task.changedDays === 1 ? "1 day ago"
                : task.changedDays + " days ago";
        chg.appendChild(el("div", "meta-line", age));
      }
      tr.appendChild(chg);

      tr.appendChild(meterCell(task, task.state === "Done"));

      tr.appendChild(el("td", "num", fmt(task.cases)));

      // A Scripting task links test cases, not a plan - it has no execution
      // outcomes of its own, and a 0 would read as "all failed to run".
      // Both branches must emit exactly 6 cells to match the 6 trailing
      // columns (Points..Not started). A mismatch silently shifts every
      // column after it, which looks like bad data rather than a layout bug.
      if (task.exec === 0) {
        ["points", "passed", "failed", "blocked", "na", "never"].forEach(function () {
          tr.appendChild(el("td", "num muted", GL.mdash));
        });
      } else {
        tr.appendChild(el("td", "num", fmt(task.exec)));
        tr.appendChild(el("td", "num", fmt(task.passed)));
        tr.appendChild(el("td", "num", fmt(task.failed)));
        tr.appendChild(el("td", "num", fmt(task.blocked)));
        tr.appendChild(el("td", "num", fmt(task.na)));
        tr.appendChild(el("td", "num", fmt(task.never)));
      }
      tb.appendChild(tr);
    });
  }

  // ---- render-all (filters scope every number at once) --------------------
  function render() {
    var rows = visible();
    var sum = totals(rows);

    renderActivity();
    renderKpis(rows, sum);
    var groups = renderBars(rows);
    renderChartTable(groups);

    var lg = renderLoad(rows);
    renderLoadTable(lg);
    // Title is fixed; the "Group by" control already states the grouping, so
    // restating it here just made the heading jitter when the toggle changed.
    $("loadTitle").textContent = "Assigned Work Items";
    $("loadSub").textContent = state.loadColor === "urgency"
      ? "Bar length is task count; segments are time remaining until the target date. Calendar days, not effort."
      : "Bar length is task count; segments are workflow state. Switch Colour by to see time remaining.";
    $("loadWrap").classList.toggle("hidden", state.loadTableView);
    $("loadTable").classList.toggle("hidden", !state.loadTableView);
    $("loadTableBtn").setAttribute("aria-pressed", state.loadTableView ? "true" : "false");

    renderTasks(rows);

    // Fixed title; the "Group by" control beside it already states the
    // grouping, so restating it here only made the heading jitter on change.
    $("chartTitle").textContent = "Test Points";

    var b = $("mistakeBanner");
    if (sum.mistakes > 0) { $("mistakeCount").textContent = fmt(sum.mistakes); b.classList.remove("hidden"); }
    else b.classList.add("hidden");

    // Two short stacked lines, not one long one: a single line was wide enough
    // to push this readout onto a row of its own.
    // "case links" is deliberately NOT called "distinct" - it is the sum of
    // each task's distinct-case count, so a case linked from two tasks counts
    // twice.
    var sn = $("scopeNote");
    sn.textContent = "";
    sn.appendChild(el("div", null, fmt(rows.length) + " of " + fmt(TASKS.length) + " tasks in scope"));
    sn.appendChild(el("div", null, fmt(sum.cases) + " case links" + SEP + fmt(sum.exec) + " test points"));

    $("chartWrap").classList.toggle("hidden", state.tableView);
    $("chartTable").classList.toggle("hidden", !state.tableView);
    $("tableBtn").setAttribute("aria-pressed", state.tableView ? "true" : "false");
  }

  // ---- wire up -------------------------------------------------------------
  // Populate dynamic panels and wire the three generic filters.
  fillCheckboxes("fPerson",  uniq("assignee"), new Set());
  fillCheckboxes("fProduct", uniq("product"),  new Set());
  fillCheckboxes("fKind",    uniq("kind"),     new Set());
  wirePanel("fPerson",  "person",  "All people");
  wirePanel("fProduct", "product", "All products");
  wirePanel("fKind",    "kind",    "All kinds");

  // fState: static panel (To Do / In Progress / Done).
  var fStatePanel = $("fState").querySelector(".multi-panel");
  $("fState").querySelector(".multi-btn").addEventListener("click", function (e) {
    e.stopPropagation(); openPanel("fState");
  });
  fStatePanel.addEventListener("change", function () {
    state.state = setFromPanel(fStatePanel);
    $("fState").querySelector(".multi-label").textContent = labelFromSet(state.state, "All states");
    render();
  });

  // fActivity: static panel; checking a completion filter clears Task state
  // (Done tasks are excluded by the default Active state, so completion+active
  // would always yield nothing - same guard as before, now clears checkboxes).
  var fActPanel = $("fActivity").querySelector(".multi-panel");
  $("fActivity").querySelector(".multi-btn").addEventListener("click", function (e) {
    e.stopPropagation(); openPanel("fActivity");
  });
  fActPanel.addEventListener("change", function (e) {
    // Show/hide date inputs when the Custom range checkbox is toggled.
    if (e.target.value === "range") { $("actRange").hidden = !e.target.checked; }
    state.activity = setFromPanel(fActPanel);
    if (state.activity.has("c7") || state.activity.has("c30")) {
      fStatePanel.querySelectorAll("input[type=checkbox]").forEach(function (cb) { cb.checked = false; });
      state.state = new Set();
      $("fState").querySelector(".multi-label").textContent = "All states";
    }
    $("fActivity").querySelector(".multi-label").textContent = activityLabel();
    render();
  });
  // Date inputs fire "change" (not "input") when the picker commits a value.
  // They live inside .multi-sel so the stopPropagation on that container already
  // prevents them from closing the panel.
  $("rangeStart").addEventListener("change", function (e) {
    state.activityRange.start = e.target.value;
    $("fActivity").querySelector(".multi-label").textContent = activityLabel();
    render();
  });
  $("rangeEnd").addEventListener("change", function (e) {
    state.activityRange.end = e.target.value;
    $("fActivity").querySelector(".multi-label").textContent = activityLabel();
    render();
  });

  // Close all panels when clicking anywhere outside them.
  // Clicks inside a panel stop propagation so checkboxes don't trigger this.
  document.querySelectorAll(".multi-sel").forEach(function (sel) {
    sel.addEventListener("click", function (e) { e.stopPropagation(); });
  });
  document.addEventListener("click", function () {
    document.querySelectorAll(".multi-panel").forEach(function (p) {
      p.hidden = true;
      p.closest(".multi-sel").querySelector(".multi-btn").setAttribute("aria-expanded", "false");
    });
  });

  renderLegend();
  renderLoadLegend();

  $("fGroup").addEventListener("change",   function (e) { state.group = e.target.value; render(); });
  $("fText").addEventListener("input",     function (e) { state.text = e.target.value; render(); });
  $("fReset").addEventListener("click", function () {
    ["fPerson","fProduct","fKind"].forEach(function (id) {
      $(id).querySelectorAll("input[type=checkbox]").forEach(function (cb) { cb.checked = false; });
    });
    $("fState").querySelectorAll("input[type=checkbox]").forEach(function (cb) {
      cb.checked = (cb.value === "To Do" || cb.value === "In Progress");
    });
    $("fActivity").querySelectorAll("input[type=checkbox]").forEach(function (cb) { cb.checked = cb.value === "w7"; });
    $("actRange").hidden = true;
    $("rangeStart").value = ""; $("rangeEnd").value = "";
    state.person = new Set(); state.product = new Set();
    state.state = new Set(["To Do", "In Progress", "Done"]);
    state.kind = new Set(); state.activity = new Set(["w7"]);
    state.activityRange = { start: "", end: "" }; state.text = "";
    $("fText").value = ""; $("fGroup").value = "assignee"; state.group = "assignee";
    $("fPerson").querySelector(".multi-label").textContent   = "All people";
    $("fProduct").querySelector(".multi-label").textContent  = "All products";
    $("fState").querySelector(".multi-label").textContent    = "3 selected";
    $("fKind").querySelector(".multi-label").textContent     = "All kinds";
    $("fActivity").querySelector(".multi-label").textContent = "Last 7d";
    render();
  });
  $("tableBtn").addEventListener("click", function () { state.tableView = !state.tableView; render(); });
  $("fLoadGroup").addEventListener("change", function (e) { state.loadGroup = e.target.value; render(); });
  $("fLoadColor").addEventListener("change", function (e) {
    state.loadColor = e.target.value;
    renderLoadLegend();          // the legend changes with the dimension
    render();
  });
  $("loadTableBtn").addEventListener("click", function () { state.loadTableView = !state.loadTableView; render(); });

  $("themeBtn").addEventListener("click", function () {
    var cur = document.documentElement.getAttribute("data-theme");
    var next = cur === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", next);
    try { localStorage.setItem("perfReportTheme", next); } catch (err) {}
  });
  try {
    var saved = localStorage.getItem("perfReportTheme");
    if (saved) document.documentElement.setAttribute("data-theme", saved);
  } catch (err) {}

  $("genAt").textContent = META.generated || "";
  $("asOf").textContent = META.asOf || "";
  $("provenance").textContent =
    "Source: csv/connected_pbi_task_test_results.csv - " + fmt(META.sourceRows) +
    " joined rows across " + fmt(META.taskCount) + " tasks, " + fmt(META.execTotal) +
    " test points from test plans.";
  // State the field mapping openly. Anyone cross-checking a Target against ADO
  // will find a box labelled "Planned Hot Fix Release Date"; without this note
  // that looks like a bug in the report.
  if (META.targetNote) {
    var pn = document.createElement("div");
    pn.textContent = META.targetNote;
    $("provenance").parentNode.insertBefore(pn, $("provenance").nextSibling);
  }

  render();
})();
</script>
</body>
</html>
'@

    $html = $template.Replace('/*__DATA__*/', $json)
    # -Encoding UTF8 writes a BOM in PS 5.1; the <meta charset> covers browsers
    # either way, and the BOM keeps Notepad/Excel happy with non-ASCII names.
    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8

    $size = [math]::Round((Get-Item -LiteralPath $OutputPath).Length / 1KB, 1)
    Write-Host ("Wrote {0} ({1} KB)" -f $OutputPath, $size)

    if ($Show) { Start-Process $OutputPath }
}
catch {
    Write-Error "Failed: $($_.Exception.Message)"
    exit 1
}
