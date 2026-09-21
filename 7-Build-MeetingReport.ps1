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
$AsOf = [datetime]::UtcNow.AddHours(8)   # Philippine Time (UTC+8)

# Parse an ISO date/datetime string as UTC regardless of the machine's local
# timezone. TryParse without RoundtripKind treats the Z suffix as local time
# on Windows, which shifts day counts by the UTC offset of the runner.
$_IsoStyles = [System.Globalization.DateTimeStyles]::RoundtripKind
$_IcCulture = [System.Globalization.CultureInfo]::InvariantCulture
function Parse-IsoUtc {
    param([string]$IsoDate, [ref]$Result)
    return [datetime]::TryParse($IsoDate, $_IcCulture, $_IsoStyles, $Result)
}

function Get-DaysSince {
    param([string]$IsoDate)
    if ([string]::IsNullOrWhiteSpace($IsoDate)) { return -1 }   # -1 = no date
    $parsed = [datetime]::MinValue
    if (Parse-IsoUtc $IsoDate ([ref]$parsed)) {
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
    if (Parse-IsoUtc $IsoDate ([ref]$parsed)) {
        return [int][math]::Floor(($parsed.Date - $AsOf.Date).TotalDays)
    }
    return $null
}

function Get-WorkingDaysUntil {
    # Same contract as Get-DaysUntil but counts Mon-Fri days only.
    # Negative return means overdue; null means no/unparseable date.
    param([string]$IsoDate)
    if ([string]::IsNullOrWhiteSpace($IsoDate)) { return $null }
    $parsed = [datetime]::MinValue
    if (-not (Parse-IsoUtc $IsoDate ([ref]$parsed))) { return $null }
    $target = $parsed.Date
    $start  = $AsOf.Date
    if ($target -eq $start) { return 0 }
    $step = if ($target -gt $start) { 1 } else { -1 }
    $count = 0
    $cur = $start.AddDays($step)
    while (($step -eq 1 -and $cur -le $target) -or ($step -eq -1 -and $cur -ge $target)) {
        if ($cur.DayOfWeek -ne [DayOfWeek]::Saturday -and $cur.DayOfWeek -ne [DayOfWeek]::Sunday) {
            $count += $step
        }
        $cur = $cur.AddDays($step)
    }
    return $count
}

function Get-WorkingDaysSince {
    # Mirrors Get-DaysSince but counts Mon-Fri days only. -1 = no date.
    param([string]$IsoDate)
    if ([string]::IsNullOrWhiteSpace($IsoDate)) { return -1 }
    $parsed = [datetime]::MinValue
    if (-not (Parse-IsoUtc $IsoDate ([ref]$parsed))) { return -1 }
    $target = $parsed.Date
    $start  = $AsOf.Date
    if ($target -ge $start) { return 0 }   # future/same date -> clamp
    $count = 0
    $cur = $target.AddDays(1)
    while ($cur -le $start) {
        if ($cur.DayOfWeek -ne [DayOfWeek]::Saturday -and $cur.DayOfWeek -ne [DayOfWeek]::Sunday) {
            $count++
        }
        $cur = $cur.AddDays(1)
    }
    return $count
}

function Get-DateOnly {
    param([string]$IsoDate)
    if ([string]::IsNullOrWhiteSpace($IsoDate)) { return '' }
    $parsed = [datetime]::MinValue
    if (Parse-IsoUtc $IsoDate ([ref]$parsed)) { return $parsed.ToString('yyyy-MM-dd') }
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

function Get-Numeric {
    param([string]$v)
    $n = 0.0
    if (-not [string]::IsNullOrWhiteSpace($v) -and
        [double]::TryParse($v.Trim(), [System.Globalization.NumberStyles]::Any,
            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { $n }
    else { 0.0 }
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
    # --- Previous bug counts for delta detection. Written after each build so
    # --- the next run can show +N when a new bug is linked.
    $bugCountsPrevPath = Join-Path $CsvDir 'bug_counts_prev.json'
    $bugCountsPrev = @{}
    if (Test-Path -LiteralPath $bugCountsPrevPath) {
        try {
            $prevJson = Get-Content -LiteralPath $bugCountsPrevPath -Raw -Encoding UTF8
            (ConvertFrom-Json $prevJson).PSObject.Properties | ForEach-Object {
                $bugCountsPrev[$_.Name] = [int]$_.Value
            }
            Write-Host ("Previous bug counts loaded: {0} task(s)" -f $bugCountsPrev.Count)
        } catch {
            Write-Warning "Could not read bug_counts_prev.json - delta will not be shown this run."
        }
    }

    # --- Rolling history of weekly status counts for the trend chart.
    # --- Each run appends one entry keyed by asOf date; capped at 12 weeks.
    # --- The previous entry (second-to-last after append) drives the delta tiles.
    $statusHistPath = Join-Path $CsvDir 'status_counts_history.json'
    $statusHistory  = [System.Collections.Generic.List[object]]::new()
    if (Test-Path -LiteralPath $statusHistPath) {
        try {
            $loaded = ConvertFrom-Json (Get-Content -LiteralPath $statusHistPath -Raw -Encoding UTF8)
            foreach ($e in @($loaded)) { $statusHistory.Add($e) }
            Write-Host ("Status history loaded: {0} week(s)" -f $statusHistory.Count)
        } catch {
            Write-Warning "Could not read status_counts_history.json - starting fresh history."
        }
    }
    # Seed from the legacy single-snapshot file when history is empty on first
    # migration. Use asOf-7d as the approximate date since the file has none.
    if ($statusHistory.Count -eq 0) {
        $legacyPrevPath = Join-Path $CsvDir 'status_counts_prev.json'
        if (Test-Path -LiteralPath $legacyPrevPath) {
            try {
                $legacyPrev = ConvertFrom-Json (Get-Content -LiteralPath $legacyPrevPath -Raw -Encoding UTF8)
                $seedDate   = $AsOf.AddDays(-7).ToString('yyyy-MM-dd')
                $statusHistory.Add([pscustomobject]@{
                    date       = $seedDate
                    toDo       = [int]$legacyPrev.toDo
                    inProgress = [int]$legacyPrev.inProgress
                    done       = [int]$legacyPrev.done
                    bugs       = [int]$legacyPrev.bugs
                })
                Write-Host ("Seeded history from legacy snapshot (date: {0})" -f $seedDate)
            } catch {
                Write-Warning "Could not seed from status_counts_prev.json."
            }
        }
    }
    # $statusPrev is set after $weekTuesdayStr is known (see below).

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

    # --- Bugs linked directly to Tasks (task_bug_links.csv from 8-Get-PbiBugLinks.ps1).
    # --- These are merged with PBI-level bugs at task-object build time below.
    $bugsByTask = @{}
    $taskBugLinksPath = Join-Path $CsvDir 'task_bug_links.csv'
    if (Test-Path -LiteralPath $taskBugLinksPath) {
        $taskBugUnreadable = 0
        foreach ($r in (Import-Csv -LiteralPath $taskBugLinksPath)) {
            if ($r.Readable -eq 'False') { $taskBugUnreadable++; continue }
            if ($r.'Target Type' -ne 'Bug') { continue }
            $bugTaskId = $r.'Task ID'
            if (-not $bugsByTask.ContainsKey($bugTaskId)) {
                $bugsByTask[$bugTaskId] = [System.Collections.Generic.HashSet[string]]::new()
            }
            [void]$bugsByTask[$bugTaskId].Add($r.'Target ID')
        }
        Write-Host ("Task bug links: {0} Task(s) carry bugs; {1} related target(s) unreadable" -f $bugsByTask.Count, $taskBugUnreadable)
    } else {
        Write-Warning "task_bug_links.csv not found - task-level bug counts will be blank. Run 8-Get-PbiBugLinks.ps1."
    }

    # --- All discussion entries per work item, from 9-Get-WorkItemComments.ps1.
    # --- Rows arrive oldest-first (the order script 9 emits them). The list is
    # --- kept in that order; discDays uses the last (most-recent) entry.
    # --- Optional: if the CSV is absent the pipeline still completes and the
    # --- HTML report renders normally - discussion icons simply do not appear.
    $commentsByItem = @{}
    $commentsPath = Join-Path $CsvDir 'workitem_comments.csv'
    if (Test-Path -LiteralPath $commentsPath) {
        foreach ($r in (Import-Csv -LiteralPath $commentsPath)) {
            if ($r.WorkItemId -and $r.CommentHtml) {
                if (-not $commentsByItem.ContainsKey($r.WorkItemId)) {
                    $commentsByItem[$r.WorkItemId] = [System.Collections.Generic.List[object]]::new()
                }
                $commentsByItem[$r.WorkItemId].Add([pscustomobject]@{
                    author = [string]$r.Author
                    date   = Get-DateOnly $r.Date
                    html   = [string]$r.CommentHtml
                })
            }
        }
        $totalComments = ($commentsByItem.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
        Write-Host ("Discussion entries loaded: {0} comment(s) across {1} work item(s)" -f $totalComments, $commentsByItem.Count)
    } else {
        Write-Warning "workitem_comments.csv not found - discussion icons will be absent. Run 9-Get-WorkItemComments.ps1."
    }

    # --- Remaining Work change history from 10-Get-WorkItemRemainingWork.ps1.
    # --- Optional: if the CSV is absent the report renders without RW data.
    # --- Only the most-recent change per task is kept (CSV is oldest-first,
    # --- so each write overwrites the previous, leaving the last entry).
    $remWorkByTask = @{}
    $remWorkPath = Join-Path $CsvDir 'workitem_remwork_history.csv'
    if (Test-Path -LiteralPath $remWorkPath) {
        foreach ($r in (Import-Csv -LiteralPath $remWorkPath)) {
            if ($r.WorkItemId) {
                $remWorkByTask[$r.WorkItemId] = @{
                    old  = $r.OldValue
                    new  = $r.NewValue
                    date = $r.RevisionDate
                }
            }
        }
        Write-Host ("Remaining Work history loaded: {0} task(s) with changes" -f $remWorkByTask.Count)
    }

    # --- Test Case state change history from 11-Get-TestCaseStateHistory.ps1.
    # --- Optional: if the CSV is absent the report renders without TC state
    # --- change data. All revisions are kept (oldest-first per TC) so the
    # --- per-task aggregation can replay history to any cutoff date.
    $tcHistByTc = @{}   # TestCaseId -> List[object] of @{date;old;new}, oldest-first
    $tcStateHistPath = Join-Path $CsvDir 'workitem_tc_state_history.csv'
    if (Test-Path -LiteralPath $tcStateHistPath) {
        foreach ($r in (Import-Csv -LiteralPath $tcStateHistPath)) {
            if ($r.TestCaseId) {
                # Treat the 9999-01-01 TFS sentinel as no date - it means the
                # revision is the current live one and carries no real timestamp.
                # Get-DaysSince would clamp it to 0, causing false filter hits.
                $safeDate = if ($r.RevisionDate -notlike '9999*') { $r.RevisionDate } else { '' }
                if (-not $tcHistByTc.ContainsKey($r.TestCaseId)) {
                    $tcHistByTc[$r.TestCaseId] = [System.Collections.Generic.List[object]]::new()
                }
                $tcHistByTc[$r.TestCaseId].Add(@{ date = $safeDate; old = $r.OldState; new = $r.NewState })
            }
        }
        Write-Host ("TC state history loaded: {0} test case(s) with changes" -f $tcHistByTc.Count)
    }

    # --- Activity dates. These live on the work item itself, so they come from
    # --- the link extracts rather than the connected dataset (script 6 doesn't
    # --- carry them). StateChangeDate and ClosedDate are 100% / correctly
    # --- populated; RemainingWork and Blocked are NOT (empty on every row in
    # --- this org), which is why there is no at-risk or blocker reporting here.
    $taskDates = @{}
    foreach ($link in $pbiTaskRows) {
        $taskDates[$link.'Target ID'] = @{
            changed = $link.'Target System.ChangedDate'
            closed  = $link.'Target Microsoft.VSTS.Common.ClosedDate'
        }
    }
    # Fallback for tasks with no PBI parent, which never appear above.
    # Also builds $tcStatesByTask: per scripting task, how many linked test
    # cases are in "Ready" state (scripted) vs "Design" (still being authored).
    $taskTestsPath = Join-Path $CsvDir 'task_tests_link_results.csv'
    $tcStatesByTask = @{}
    if (Test-Path -LiteralPath $taskTestsPath) {
        foreach ($link in (Import-Csv -LiteralPath $taskTestsPath)) {
            if ($link.'Link Type' -eq '(root)') { continue }
            $id = $link.'Source ID'
            if (-not $taskDates.ContainsKey($id)) {
                $taskDates[$id] = @{
                    changed = $link.'Source System.ChangedDate'
                    closed  = $link.'Source Microsoft.VSTS.Common.ClosedDate'
                }
            }
            $tid = $link.'Target ID'
            if (-not $id -or -not $tid) { continue }
            if (-not $tcStatesByTask.ContainsKey($id)) {
                $tcStatesByTask[$id] = @{
                    Seen     = [System.Collections.Generic.HashSet[string]]::new()
                    Ready    = 0
                    Design   = 0
                    TcStates = @{}
                }
            }
            $bucket = $tcStatesByTask[$id]
            if ($bucket.Seen.Add($tid)) {
                $curState = $link.'Target System.State'
                $bucket.TcStates[$tid] = $curState
                if ($curState -eq 'Ready')  { $bucket.Ready++ }
                elseif ($curState -eq 'Design') { $bucket.Design++ }
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

    # Deduplicate Test Suite rows: per (Task ID, Test Case ID) keep only the row
    # with the highest Suite ID (highest ID = latest suite = most-recent run).
    # Scripting rows (Target Type = Test Case) are left untouched - they don't
    # contribute to outcome counts and deduping them would hide valid case links.
    $suiteRowList    = [System.Collections.Generic.List[object]]::new()
    $nonSuiteRowList = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $connected) {
        if ($r.'Target Type' -eq 'Test Suite') { $suiteRowList.Add($r) }
        else                                   { $nonSuiteRowList.Add($r) }
    }
    $bestSuite = @{}
    foreach ($r in $suiteRowList) {
        $key = $r.'Task ID' + '|' + $r.'Test Case ID'
        $sid = 0; [int]::TryParse($r.'Suite ID', [ref]$sid) | Out-Null
        if (-not $bestSuite.ContainsKey($key) -or $sid -gt $bestSuite[$key].sid) {
            $bestSuite[$key] = @{ sid = $sid; row = $r }
        }
    }
    $connected = @($nonSuiteRowList) + @($bestSuite.Values | ForEach-Object { $_.row })
    Write-Host ("After latest-suite dedup: {0} rows ({1} suite-linked TCs, {2} scripting-linked)" -f `
        $connected.Count, $bestSuite.Count, $nonSuiteRowList.Count)

    # --- Aggregate to one record per task. Filters are all task-level
    # --- attributes, so the browser can re-scope everything by filtering this
    # --- small array - no need to ship 3900 raw rows to the page.
    $byTask = [System.Collections.Specialized.OrderedDictionary]::new()

    function New-TaskRecord {
        param([string]$Id, [string]$Title, [string]$State, [string]$Kind,
              [string]$Assignee, [string]$Product, [string]$PbiId, [string]$PbiTitle,
              [string]$PbiStart, [string]$PbiTarget, [string]$RemWork = '')
        if ([string]::IsNullOrWhiteSpace($Assignee)) { $Assignee = '(unassigned)' }
        [pscustomobject]@{
            id = $Id; title = $Title; state = $State; kind = $Kind
            assignee = $Assignee; product = $Product; pbiId = $PbiId; pbiTitle = $PbiTitle
            pbiStart = $PbiStart; pbiTarget = $PbiTarget
            remWork   = Get-Numeric $RemWork
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
            -PbiTarget ([string]$link.'Source Deltek.PlanHotFixRelDt') `
            -RemWork   ([string]$link.'Target Microsoft.VSTS.Scheduling.RemainingWork')
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
                remWork   = 0.0
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

    $cutoff7 = $AsOf.AddDays(-7)

    $tasks = foreach ($t in $byTask.Values) {
        $d = $taskDates[$t.id]
        $changedIso = if ($d) { $d.changed } else { '' }
        $closedIso  = if ($d) { $d.closed }  else { '' }

        # Aggregate TC state change history for this task's linked test cases.
        # Computes how many TCs were in Ready/Design state 7 days ago vs. now
        # so the display can show "Ready: 28 -> 29 . Design: 6 -> 5".
        # tcStateChangeDays (-1 sentinel) is used by the activity filter.
        $tcReady7Ago = 0; $tcDesign7Ago = 0; $tcStateChangeDays = -1
        if ($tcStatesByTask.ContainsKey($t.id)) {
            $bucket7 = $tcStatesByTask[$t.id]
            foreach ($tcId in $bucket7.Seen) {
                # --- State 7 days ago ---
                $hist7  = $tcHistByTc[$tcId]
                $state7 = $null
                if ($hist7 -and $hist7.Count) {
                    $stateAtCutoff = $null
                    $foundBefore   = $false
                    foreach ($h in $hist7) {           # oldest-first; last match wins
                        if (-not $h.date) { continue }
                        $hp = [datetime]::MinValue
                        if ([datetime]::TryParse($h.date, [ref]$hp) -and $hp -le $cutoff7) {
                            $stateAtCutoff = $h.new
                            $foundBefore   = $true
                        }
                    }
                    $state7 = if ($foundBefore) { $stateAtCutoff } else { $hist7[0].old }
                }
                if ($null -eq $state7) { $state7 = $bucket7.TcStates[$tcId] }

                if     ($state7 -eq 'Ready')  { $tcReady7Ago++ }
                elseif ($state7 -eq 'Design') { $tcDesign7Ago++ }

                # --- Most-recent change date for activity filter ---
                if ($hist7 -and $hist7.Count) {
                    $lastH = $hist7[$hist7.Count - 1]
                    if ($lastH.date) {
                        $dAgo = Get-DaysSince $lastH.date
                        if ($tcStateChangeDays -lt 0 -or $dAgo -lt $tcStateChangeDays) {
                            $tcStateChangeDays = $dAgo
                        }
                    }
                }
            }
        }

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
            closedOn      = Get-DateOnly       $closedIso
            closedDays    = Get-DaysSince       $closedIso
            workClosedDays = Get-WorkingDaysSince $closedIso
            # Deadline pair. NO_TARGET keeps daysLeft numeric so the column
            # stays sortable and filterable; the UI checks targetOn to decide
            # whether to show anything at all.
            startOn     = Get-DateOnly $t.pbiStart
            targetOn    = Get-DateOnly $t.pbiTarget
            # Bug IDs travel as a list, not just a count, so the browser can
            # de-duplicate at group level: one bug linked to two of a person's
            # PBIs must count once for that person, not twice.
            bugIds      = $(
                # Merge PBI-level and task-level Related bugs, de-duplicated.
                # A bug linked to both the PBI and the task must count once.
                $bugSet = [System.Collections.Generic.HashSet[string]]::new()
                if ($t.pbiId -and $bugsByPbi.ContainsKey($t.pbiId)) {
                    foreach ($b in $bugsByPbi[$t.pbiId]) { [void]$bugSet.Add($b) }
                }
                if ($bugsByTask.ContainsKey($t.id)) {
                    foreach ($b in $bugsByTask[$t.id]) { [void]$bugSet.Add($b) }
                }
                if ($bugSet.Count) { (@($bugSet) | Sort-Object) -join ',' } else { '' }
            )
            bugDelta    = $(
                $curCount = $bugSet.Count
                $prevCount = if ($bugCountsPrev.ContainsKey($t.id)) { $bugCountsPrev[$t.id] } else { $null }
                if ($null -ne $prevCount -and $curCount -gt $prevCount) { $curCount - $prevCount } else { 0 }
            )
            daysLeft    = $(
                $d = Get-DaysUntil $t.pbiTarget
                if ($null -eq $d) { 99999 } else { $d }
            )
            workDaysLeft = $(
                $d = Get-WorkingDaysUntil $t.pbiTarget
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
            workWindowDays = $(
                $s = Get-WorkingDaysUntil $t.pbiStart
                $e = Get-WorkingDaysUntil $t.pbiTarget
                if ($null -eq $s -or $null -eq $e) { 0 }
                elseif (($e - $s) -le 0) { 0 }
                else { [int]($e - $s) }
            )
            cases    = $t.CaseSet.Count
            tcReady  = if ($tcStatesByTask.ContainsKey($t.id)) { $tcStatesByTask[$t.id].Ready }  else { 0 }
            tcDesign = if ($tcStatesByTask.ContainsKey($t.id)) { $tcStatesByTask[$t.id].Design } else { 0 }
            testers  = ((@($t.TesterSet) | Sort-Object) -join ', ')
            exec     = $t.exec
            passed   = $t.passed
            failed   = $t.failed
            blocked  = $t.blocked
            na       = $t.na
            never    = $t.never
            mistakes = $t.mistakes
            remWork    = [double]$t.remWork
            remWorkOld = if ($remWorkByTask.ContainsKey($t.id)) { $remWorkByTask[$t.id].old }  else { $null }
            remWorkNew = if ($remWorkByTask.ContainsKey($t.id)) { $remWorkByTask[$t.id].new }  else { $null }
            tcReady7Ago       = $tcReady7Ago
            tcDesign7Ago      = $tcDesign7Ago
            tcStateChangeDays = $tcStateChangeDays
            discDays = $(
                $discList = $commentsByItem[$t.id]
                if ($discList -and $discList.Count) { Get-DaysSince $discList[$discList.Count - 1].date } else { -1 }
            )
            taskDisc = $(
                $tc = $commentsByItem[$t.id]
                if ($tc -and $tc.Count) { @($tc) } else { $null }
            )
            pbiDisc = $(
                if ($t.pbiId) {
                    $pc = $commentsByItem[$t.pbiId]
                    if ($pc -and $pc.Count) { @($pc) } else { $null }
                } else { $null }
            )
        }
    }
    $tasks = @($tasks)

    $execTotal = ($tasks | Measure-Object -Property exec -Sum).Sum
    Write-Host ("Aggregated to {0} tasks; {1} test-plan-sourced test points" -f $tasks.Count, $execTotal)

    # --- Weekly activity counts (not cumulative totals).
    # --- toDo/inProgress = tasks in that state touched this week (changedDays <= 7).
    # --- done            = tasks closed this week (closedDays <= 7).
    # --- cntBugs         = current total bugs (used to compute weekly new-bug delta).
    $weeklyToDo  = @($tasks | Where-Object { $_.state -eq 'To Do'       -and $_.changedDays -ge 0 -and $_.changedDays -le 7 }).Count
    $weeklyInPrg = @($tasks | Where-Object { $_.state -eq 'In Progress' -and $_.changedDays -ge 0 -and $_.changedDays -le 7 }).Count
    $weeklyDone  = @($tasks | Where-Object { $_.closedDays -ge 0 -and $_.closedDays -le 7 }).Count
    # Snapshot counts: total tasks in each state regardless of activity window.
    $snapToDo  = @($tasks | Where-Object { $_.state -eq 'To Do' }).Count
    $snapInPrg = @($tasks | Where-Object { $_.state -eq 'In Progress' }).Count
    $snapDone  = @($tasks | Where-Object { $_.state -eq 'Done' }).Count
    $allBugIdSet = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($t in $tasks) {
        if ($t.bugIds) { foreach ($b in ($t.bugIds -split ',')) { [void]$allBugIdSet.Add($b) } }
    }
    $cntBugs = $allBugIdSet.Count
    # New bugs this week = current total minus the previous week's total (floor 0).
    $prevTotalBugs = if ($statusPrev -and $null -ne $statusPrev.totalBugs) { [int]$statusPrev.totalBugs } `
                     elseif ($statusPrev -and $null -ne $statusPrev.bugs)   { [int]$statusPrev.bugs }     `
                     else { $cntBugs }
    $weeklyBugs = [math]::Max(0, $cntBugs - $prevTotalBugs)

    function Get-StatusDelta([int]$cur, $prev, [string]$key) {
        if ($null -eq $prev) { return $null }
        $p = $prev.$key
        if ($null -eq $p) { return $null }
        return $cur - [int]$p
    }

    # Anchor history entries to the closing Tuesday of the current week.
    # Tuesday is the week-close, so mid-week runs key forward to the coming Tuesday
    # rather than backward to the past one. DayOfWeek: Sun=0, Mon=1, Tue=2, ..., Sat=6.
    $daysToNextTuesday = (2 - [int]$AsOf.DayOfWeek + 7) % 7   # 0 when today IS Tuesday
    $weekTuesdayDate   = $AsOf.AddDays($daysToNextTuesday).Date
    $weekTuesdayStr    = $weekTuesdayDate.ToString('yyyy-MM-dd')
    # Days elapsed since the previous Tuesday (for retroactive window math).
    $daysSincePrevTue  = if ($daysToNextTuesday -eq 0) { 7 } else { 7 - $daysToNextTuesday }
    # Previous week's entry — skip the current week's own entry so re-runs don't
    # compare this week's count against itself and produce a bogus delta of 0.
    $statusPrev = $statusHistory | Where-Object { $_.date -ne $weekTuesdayStr } | Select-Object -Last 1

    # Retroactive seeding: when no history exists, approximate prior-week counts
    # from changedDays/closedDays windows anchored to Tuesday boundaries.
    # Week N-back runs from Tuesday N*7 days ago to the Tuesday (N-1)*7 days ago.
    # changedDays window: [daysSincePrevTue+(N-1)*7+1 .. daysSincePrevTue+N*7].
    # Bugs delta is omitted for retroactive entries (no prior snapshot to diff).
    if ($statusHistory.Count -eq 0) {
        for ($wb = 3; $wb -ge 1; $wb--) {
            $winStart = $daysSincePrevTue + ($wb - 1) * 7 + 1
            $winEnd   = $daysSincePrevTue + $wb * 7
            $seedDate = $weekTuesdayDate.AddDays(-($wb * 7)).ToString('yyyy-MM-dd')
            $sToDo  = @($tasks | Where-Object { $_.state -eq 'To Do'       -and $_.changedDays -ge $winStart -and $_.changedDays -le $winEnd }).Count
            $sInPrg = @($tasks | Where-Object { $_.state -eq 'In Progress' -and $_.changedDays -ge $winStart -and $_.changedDays -le $winEnd }).Count
            $sDone  = @($tasks | Where-Object { $_.closedDays -ge $winStart -and $_.closedDays -le $winEnd }).Count
            $statusHistory.Add([pscustomobject]@{ date=$seedDate; toDo=$sToDo; inProgress=$sInPrg; done=$sDone; bugs=0; totalBugs=$cntBugs; snapToDo=$snapToDo; snapInProgress=$snapInPrg; snapDone=0 })
            Write-Host ("Retroactive seed: {0}  ToDo={1}  InPrg={2}  Done={3}" -f $seedDate, $sToDo, $sInPrg, $sDone)
        }
    }

    # Build updated history: dedup by Tuesday date, append this week's entry, cap at 12.
    $asOfStr = $AsOf.ToString('yyyy-MM-dd')

    # Per-person weekly breakdown for filtered chart and tile deltas.
    $byPerson = @{}
    $byPersonBugSets = @{}
    foreach ($t in $tasks) {
        $name = $t.assignee
        if (-not $byPerson.ContainsKey($name)) {
            $byPerson[$name] = [pscustomobject]@{ toDo=0; inProgress=0; done=0; totalBugs=0; snapToDo=0; snapInProgress=0; snapDone=0 }
            $byPersonBugSets[$name] = [System.Collections.Generic.HashSet[string]]::new()
        }
        if ($t.state -eq 'To Do'       -and $t.changedDays -ge 0 -and $t.changedDays -le 7) { $byPerson[$name].toDo++ }
        if ($t.state -eq 'In Progress' -and $t.changedDays -ge 0 -and $t.changedDays -le 7) { $byPerson[$name].inProgress++ }
        if ($t.closedDays -ge 0 -and $t.closedDays -le 7)                                   { $byPerson[$name].done++ }
        if ($t.bugIds) { foreach ($b in ($t.bugIds -split ',')) { [void]$byPersonBugSets[$name].Add($b) } }
        if ($t.state -eq 'To Do')        { $byPerson[$name].snapToDo++ }
        if ($t.state -eq 'In Progress')  { $byPerson[$name].snapInProgress++ }
        if ($t.state -eq 'Done')         { $byPerson[$name].snapDone++ }
    }
    foreach ($pName in @($byPerson.Keys)) { $byPerson[$pName].totalBugs = $byPersonBugSets[$pName].Count }

    # Same-week dedup: any existing entry for this Tuesday key is removed before
    # appending the fresh one, so re-pulls within the week update rather than grow.
    $updatedHistory = [System.Collections.Generic.List[object]]::new()
    foreach ($e in @($statusHistory)) {
        if ($e.date -eq $weekTuesdayStr) { continue }
        $updatedHistory.Add($e)
    }
    $updatedHistory.Add([pscustomobject]@{ date=$weekTuesdayStr; toDo=$weeklyToDo; inProgress=$weeklyInPrg; done=$weeklyDone; bugs=$weeklyBugs; totalBugs=$cntBugs; snapToDo=$snapToDo; snapInProgress=$snapInPrg; snapDone=$snapDone; byPerson=$byPerson })
    while ($updatedHistory.Count -gt 12) { $updatedHistory.RemoveAt(0) }

    Write-Host ("Status history: keyed to Tuesday {0} ({1} week(s) total)" -f $weekTuesdayStr, $updatedHistory.Count)

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
        times   = [string][char]0x00D7   # multiplication sign (dismiss button)
        up      = [string][char]0x2191
        arrow   = [string][char]0x2192   # right arrow (A -> B)
        down    = [string][char]0x2193
    }

    $payload = [pscustomobject][ordered]@{
        meta = [pscustomobject][ordered]@{
            generated  = $AsOf.ToString('yyyy-MM-dd hh:mm tt') + ' PHT'
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
            tfsBase       = 'https://tfs.deltek.com/tfs/Deltek/QEAutomation/_workitems/edit/'
            statusCounts = [pscustomobject]@{
                toDo       = $weeklyToDo
                inProgress = $weeklyInPrg
                done       = $weeklyDone
                bugs       = $weeklyBugs
                totalBugs  = $cntBugs
            }
            statusDeltas = [pscustomobject]@{
                toDo       = (Get-StatusDelta $weeklyToDo  $statusPrev 'toDo')
                inProgress = (Get-StatusDelta $weeklyInPrg $statusPrev 'inProgress')
                done       = (Get-StatusDelta $weeklyDone  $statusPrev 'done')
            }
            statusHistory = @($updatedHistory)
        }
        tasks = $tasks
    }

    $json = $payload | ConvertTo-Json -Depth 10 -Compress
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
  .scripting-status { text-align: left; vertical-align: middle; padding-left: 4em; }
  .rw-history, .tc-state-hist { font-size: 11.5px; margin-top: 4px; color: var(--ink-2); }
  .rw-label   { color: var(--ink-muted); }
  .rw-val     { font-weight: 500; }
  .rw-none    { color: var(--ink-muted); font-style: italic; }
  .tc-ready  { color: var(--st-good); font-weight: 500; }
  .tc-design { color: var(--ink-muted); }
  .tc-sep    { margin: 0 0.5em; color: var(--ink-muted); }

  header { display: flex; align-items: center; gap: 12px; margin-bottom: 16px; }
  button {
    font: inherit; color: var(--ink); background: var(--surface);
    border: 1px solid var(--border); border-radius: 6px;
    padding: 6px 11px; cursor: pointer; min-height: 32px;
  }
  .icon-btn[aria-pressed="true"] { background: rgba(59,130,246,.14); border-color: #3b82f6; color: #3b82f6; }
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
  .filters { display: flex; flex-wrap: nowrap; gap: 12px; align-items: center; }
  .filter-controls {
    display: flex; flex-wrap: wrap; gap: 6px 8px; align-items: center;
    flex: 1 1 auto; min-width: 0;
  }

  /* Sticky filter strip — floating pill bar look. Solid surface prevents
     scrolling content from showing through; shadow is subtle (not heavy). */
  .card.filters {
    position: sticky; top: 0; z-index: 30;
    background: var(--surface);
    border-color: var(--rule);
    border-top: none;
    border-radius: 0 0 10px 10px;
    box-shadow: 0 4px 18px rgba(0,0,0,.10);
    padding: 10px 14px;
  }
  /* Hide the text label above each control — the placeholder text inside the
     control already identifies it, so the label is visual noise. The <label>
     element is kept for screen-reader association. */
  .filter-controls > label { display: contents; font-size: 0; }

  /* Base form controls — scoped to filter bar so chart-section selects are unaffected. */
  select, input[type="search"] {
    font: inherit; color: var(--ink); background: var(--surface);
    border: 1px solid var(--border); border-radius: 6px;
    padding: 6px 8px; min-height: 32px; min-width: 150px;
  }
  .filter-controls select, .filter-controls input[type="search"] {
    background: var(--plane); border-color: var(--rule); border-radius: 20px;
    padding: 3px 12px; min-height: 28px; min-width: 120px; font-size: 12px;
  }
  .scope {
    flex: 0 0 auto; font-size: 12px; color: var(--ink-2);
    text-align: right; line-height: 1.45; white-space: nowrap;
  }
  /* Multi-select checkbox dropdown — pill style. */
  .multi-sel { position: relative; display: inline-block; }
  .multi-btn {
    font: inherit; font-size: 12px; color: var(--ink); background: var(--plane);
    border: 1px solid var(--rule); border-radius: 20px;
    padding: 3px 26px 3px 12px; min-height: 28px; min-width: 110px; max-width: 160px;
    cursor: pointer; text-align: left; position: relative;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  }
  .multi-btn::after {
    content: "\25BE"; position: absolute; right: 10px; top: 50%;
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
  .delta-good { color: var(--st-good); font-weight: 600; }
  .delta-bad  { color: var(--st-critical); font-weight: 600; }
  /* Status tiles: 2x2 grid beside the chart */
  #statusDash { display: grid; grid-template-columns: 1fr 1fr; width: 280px; flex-shrink: 0; gap: 20px 8px; align-content: center; }
  #statusDash .tile { flex: none; min-width: 0; }
  #statusDash .tile .value { font-size: 20px; }
  #statusDash .tile .label { font-size: 11px; }
  #statusDash .tile .pct   { font-size: 10px; }
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
  .pbi-row { cursor: pointer; }
  a.tfs-link { color: var(--ink-muted); margin-left: 5px; vertical-align: middle;
               text-decoration: none; opacity: .55; display: inline-flex; }
  a.tfs-link:hover { opacity: 1; color: var(--ts-doing); }
  .pbi-selected { background: rgba(57,135,229,0.10); box-shadow: inset 3px 0 0 var(--ts-doing); }
  .pbi-selected:hover { background: rgba(57,135,229,0.15); }
  .disc-btn { background: none; border: none; cursor: pointer; padding: 2px 5px; color: var(--ink-2); border-radius: 4px; line-height: 1; }
  .disc-btn:hover { color: var(--ink); background: var(--plane); }
  .disc-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.45); z-index: 200; display: flex; align-items: center; justify-content: center; }
  .disc-dialog { background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 20px 24px; max-width: 600px; width: 90%; max-height: 80vh; overflow-y: auto; box-shadow: 0 8px 32px rgba(0,0,0,0.22); position: relative; }
  .disc-dialog h3 { font-size: 15px; font-weight: 600; margin: 0; padding-right: 36px; line-height: 1.3; }
  .disc-close { position: absolute; top: 14px; right: 16px; background: none; border: none; cursor: pointer; font-size: 20px; color: var(--ink-2); padding: 0 6px; border-radius: 4px; line-height: 1; }
  .disc-close:hover { color: var(--ink); background: var(--plane); }
  .disc-section { margin-top: 14px; padding-top: 14px; border-top: 1px solid var(--border); }
  .disc-who { font-size: 11px; font-weight: 600; color: var(--ink-2); text-transform: uppercase; letter-spacing: 0.05em; }
  .disc-meta { font-size: 11.5px; color: var(--ink-muted); margin: 3px 0 8px; }
  .disc-body { font-size: 13.5px; line-height: 1.65; color: var(--ink); font-family: inherit; }
  .disc-body * { color: inherit !important; background-color: transparent !important; font-family: inherit !important; }
  .disc-body p { margin: 0 0 6px; }
  .disc-body p:last-child { margin-bottom: 0; }
  .disc-none { color: var(--ink-muted); font-style: italic; font-size: 13px; }
  .disc-entry + .disc-entry { margin-top: 12px; padding-top: 12px; border-top: 1px dashed var(--border); }
  /* ---- Team Capacity bars ---- */
  .cap-section { display: flex; flex-direction: column; gap: 4px; }
  .cap-row  { display: grid; grid-template-columns: 152px 1fr 200px; gap: 12px; align-items: center; padding: 6px 4px; border-radius: 6px; }
  .cap-info  { font-size: 12px; font-variant-numeric: tabular-nums; }
  .cap-rate  { font-weight: 600; }
  .cap-rate.ok   { color: var(--st-good); }
  .cap-rate.warn { color: #fab219; }
  .cap-rate.over { color: var(--st-critical); }
  .cap-detail { color: var(--ink-muted); font-size: 11px; }
  .cap-empty  { color: var(--ink-muted); font-style: italic; font-size: 13px; }
  .cap-scale  { display: flex; justify-content: space-between; padding-left: 164px; margin-top: 8px; font-size: 11px; color: var(--ink-muted); border-top: 1px solid var(--grid); padding-top: 4px; }
  .cap-col-head { text-align: right; padding-right: 8px; white-space: nowrap; }
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
  tr.filterrow, .colfoot { display: none; }
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
  #taskTable td { vertical-align: middle; }
  /* Give the meter columns real width. Without this the free space all goes to
     the Task/PBI column and the meter ends up squeezed against the number
     block - centred, but visibly cramped. */
  #taskTable th:nth-child(7), #taskTable td:nth-child(7),
  #loadTable th:nth-child(4), #loadTable td:nth-child(4) { min-width: 104px; }
  .colfoot { display: flex; align-items: center; gap: 12px; margin-top: 10px; }
  .colfoot .sub { margin: 0; }
  .colfoot button { margin-left: auto; }
  .meta-line { color: var(--ink-muted); font-size: 11.5px; margin-top: 2px; }
  .bug-delta { margin-left: 4px; font-size: 10.5px; font-weight: 600; color: #c0392b; }
  .flag-row { border-left: 3px solid #fab219; }
  .flag-chips { display: flex; flex-wrap: wrap; gap: 4px; margin-top: 4px; }
  .flag-chip { font-size: 10.5px; padding: 1px 6px; border-radius: 999px; background: #fab21922; border: 1px solid #fab219; color: #b07800; white-space: nowrap; }
  .pill { display: inline-block; font-size: 11px; padding: 1px 7px; border-radius: 999px; border: 1px solid var(--border); color: var(--ink-2); white-space: nowrap; }
  .empty { padding: 28px 8px; text-align: center; color: var(--ink-muted); }
  .hidden { display: none !important; }
  /* Card-level collapse: a header row with an h2 and a Show/Hide toggle. */
  .card-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 8px; }
  .collapse-btn { flex-shrink: 0; font-size: 12px; padding: 4px 10px; min-height: 28px; }

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

  /* Minimalistic icon buttons in the header toolbar */
  .icon-btn {
    background: none; border: 1px solid transparent; cursor: pointer;
    width: 34px; height: 34px; display: flex; align-items: center; justify-content: center;
    border-radius: 7px; color: var(--ink-2); padding: 0; position: relative; flex: 0 0 34px;
  }
  .icon-btn:hover { background: var(--surface-2); border-color: var(--rule); }
  .icon-btn svg { width: 18px; height: 18px; stroke: currentColor; fill: none;
    stroke-width: 1.75; stroke-linecap: round; stroke-linejoin: round; display: block; }
  .chart-mode-btn { width:32px; height:32px; display:flex; align-items:center; justify-content:center; border:1px solid var(--rule); border-radius:4px; background:var(--surface); color:var(--ink-2); cursor:pointer; opacity:0.55; transition:opacity .15s; }
  .chart-mode-btn:hover { opacity:0.85; }
  .chart-mode-btn.chart-mode-active { background:var(--accent,#6366f1); color:#fff; border-color:transparent; opacity:1; }
  .icon-badge {
    position: absolute; top: -3px; right: -3px;
    background: #ef4444; color: #fff; font-size: 9px; font-weight: 700;
    min-width: 15px; height: 15px; border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    padding: 0 3px; line-height: 1; pointer-events: none;
  }
  /* Activity drawer */
  #activityDrawer {
    position: fixed; top: 0; right: 0; bottom: 0; width: min(420px, 95vw);
    background: var(--surface); box-shadow: -4px 0 28px rgba(0,0,0,.18);
    z-index: 60; transform: translateX(100%); transition: transform .22s ease;
    overflow-y: auto; padding: 24px 20px;
    border-left: 1px solid var(--rule);
  }
  #activityDrawer.open { transform: translateX(0); }
  #drawerOverlay {
    display: none; position: fixed; inset: 0; z-index: 59;
    background: rgba(0,0,0,.25);
  }
  #drawerOverlay.open { display: block; }
  /* Exceptions popover */
  #exceptPopover {
    display: none; position: fixed; z-index: 55;
    background: var(--surface); border: 1px solid var(--rule);
    border-radius: 8px; padding: 12px 14px; max-width: 340px;
    box-shadow: 0 4px 18px rgba(0,0,0,.14); font-size: 13px; line-height: 1.5;
  }
  #exceptPopover.open { display: block; }

  @media print {
    body { background: #fff; padding: 0; }
    .card { break-inside: avoid; border-color: #ccc; }
    .filters, header, #tableBtn, #tip, #activityDrawer, #drawerOverlay, #exceptPopover { display: none !important; }
  }


</style>
</head>
<body>

<header>
  <h1 style="margin:0;flex:1;min-width:0;font-size:16px;font-weight:600">
    Performance Testing &mdash; Weekly Meeting Report
    <span class="muted" style="font-size:11px;font-weight:400;margin-left:10px;white-space:nowrap">Generated <span id="genAt"></span></span>
  </h1>
  <div style="display:flex;gap:6px;align-items:center;flex-shrink:0">
    <button id="exceptBtn" class="icon-btn" type="button" title="Link exceptions" hidden>
      <svg viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
      <span class="icon-badge" id="exceptBadge"></span>
    </button>
    <button id="activityBtn" class="icon-btn" type="button" title="Activity">
      <svg viewBox="0 0 24 24"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>
    </button>
    <button id="themeBtn" class="icon-btn" type="button" title="Switch light / dark">
      <svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>
    </button>
  </div>
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
      <button class="multi-btn" type="button" aria-haspopup="true" aria-expanded="false"><span class="multi-label">2 selected</span></button>
      <div class="multi-panel" hidden>
        <label class="cb-item"><input type="checkbox" value="To Do"><span>To Do</span></label>
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
        <label class="cb-item"><input type="checkbox" value="c7"><span>Completed, last 7 days</span></label>
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
  <button id="fReset" class="icon-btn" type="button" title="Reset filters">
    <svg viewBox="0 0 24 24"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 .49-3"/></svg>
  </button>
 </div>
 <div class="scope" id="scopeNote"></div>
</div>

<!-- Activity slide-in drawer (toggled by #activityBtn in header) -->
<div id="drawerOverlay"></div>
<div id="activityDrawer" role="dialog" aria-label="Activity">
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px">
    <h2 style="margin:0">Activity <span class="muted" style="font-weight:400;font-size:14px">as of <span id="asOf"></span></span></h2>
    <button id="drawerClose" class="icon-btn" type="button" title="Close">
      <svg viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
    </button>
  </div>
  <p class="sub" style="margin-bottom:14px">All filters apply, including task state.</p>
  <div class="kpis" id="activity"></div>
</div>

<!-- Exceptions popover (anchored near #exceptBtn) -->
<div id="exceptPopover">
  <strong id="mistakeCount"></strong>
  <span> link(s) use a <code>Child</code> relationship where <code>Tests</code> is required, missing from ADO traceability. See <code>exceptions_weekly.csv</code>.</span>
</div>

<div class="card" style="margin-bottom:16px">
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">
    <h2 style="margin:0">Weekly Status</h2>
    <div style="display:flex;gap:4px;align-items:center">
      <button id="chartModeActivity" class="chart-mode-btn chart-mode-active" type="button" title="Activity view">
        <svg viewBox="0 0 20 14" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><polyline points="0,7 4,7 6,2 9,12 12,4 14,7 20,7"/></svg>
      </button>
      <button id="chartModeSnapshot" class="chart-mode-btn" type="button" title="Snapshot view">
        <svg viewBox="0 0 16 14" width="20" height="20" fill="currentColor"><rect x="0" y="8" width="4" height="6" rx="0.5"/><rect x="6" y="3" width="4" height="11" rx="0.5"/><rect x="12" y="5" width="4" height="9" rx="0.5"/></svg>
      </button>
    </div>
  </div>
  <div style="display:flex;gap:20px;align-items:flex-start;margin-top:8px">
    <div style="flex:1;min-width:0">
      <div id="statusChart"></div>
    </div>
    <div style="display:flex;flex-direction:column;align-items:center;gap:8px">
      <span id="statusScope" style="font-size:14px;font-weight:600;color:var(--ink-2);text-align:center"></span>
      <div class="kpis" id="statusDash"></div>
    </div>
  </div>
</div>

<div class="card">
  <div class="chart-head">
    <div>
      <h2 id="loadTitle">Assigned Work Items</h2>
      <p class="sub" id="loadSub">A PBI appears only if it has at least one matching task. Bar length = task count; segments show workflow state.</p>
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
          <option value="state">Task state</option>
          <option value="urgency" selected>Time to target</option>
        </select>
      </label>
      <button id="loadTableBtn" class="icon-btn" type="button" aria-pressed="false" title="Table view">
        <svg viewBox="0 0 24 24"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="9" x2="9" y2="21"/></svg>
      </button>
    </div>
  </div>
  <ul class="legend" id="loadLegend"></ul>
  <div id="loadWrap"><div class="bars" id="loadBars"></div><div class="scale" id="loadScale"></div></div>
  <div id="loadTable" class="hidden"></div>
</div>

<div class="card">
  <h2>Task Details</h2>
  <p class="sub" style="margin-bottom:12px">One row per active task, scoped by the filters above. Scripting tasks show ADO test case state (Ready / Design); Execution tasks show test plan outcomes.</p>
  <div id="taskTable"></div>
</div>

<div class="card">
  <div class="card-head">
    <h2>Execution results <span class="muted" style="font-weight:400">- test points that came from a test plan</span></h2>
    <button id="kpisToggle" class="icon-btn" type="button" aria-expanded="false" title="Show / hide">
      <svg viewBox="0 0 24 24"><polyline points="6 9 12 15 18 9"/></svg>
    </button>
  </div>
  <div id="kpisBody" hidden>
    <p class="sub" style="margin-bottom:14px">Scripting tasks are excluded here: they link test cases directly, which would double-count the same results.</p>
    <div class="kpis" id="kpis"></div>
  </div>
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
      <button id="tableBtn" class="icon-btn" type="button" aria-pressed="false" title="Table view">
        <svg viewBox="0 0 24 24"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="9" x2="9" y2="21"/></svg>
      </button>
      <button id="chartToggle" class="icon-btn" type="button" aria-expanded="false" title="Show / hide">
        <svg viewBox="0 0 24 24"><polyline points="6 9 12 15 18 9"/></svg>
      </button>
    </div>
  </div>
  <div id="chartBody" hidden>
    <ul class="legend" id="legend"></ul>
    <div id="chartWrap"><div class="bars" id="bars"></div><div class="scale" id="scale"></div></div>
    <div id="chartTable" class="hidden"></div>
  </div>
</div>

<footer>
  <div id="provenance"></div>
  <div>Rebuild: <code>Run-AdoExtracts.bat</code> (extract + reports + this page), or
       <code>powershell -File 7-Build-MeetingReport.ps1</code> to regenerate just this page from the existing CSVs.</div>
</footer>

<div id="tip" role="tooltip" aria-hidden="true"></div>

<script id="payload" type="application/json">/*__DATA__*/</script>
<div id="discOverlay" class="disc-overlay hidden" role="dialog" aria-modal="true" aria-labelledby="discTitle">
  <div class="disc-dialog">
    <button id="discClose" class="disc-close" type="button" aria-label="Close discussion">&times;</button>
    <h3 id="discTitle"></h3>
    <div id="discContent"></div>
  </div>
</div>
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
                state: new Set(["In Progress", "Done"]), chartMode: "activity",
                kind: new Set(), text: "", activity: new Set(["w7"]),
                activityRange: { start: "", end: "" },
                group: "assignee", tableView: false, sortKey: "exec", sortDir: -1,
                loadGroup: "assignee", loadColor: "urgency", loadTableView: false,
                selectedPbiKey: null, _tableAutoSelected: false, colFilters: {} };

  // changedDays / closedDays are -1 when the date is absent, so a plain
  // "<= 7" test would wrongly match those. Always require >= 0 first.
  function withinDays(value, limit) { return value >= 0 && value <= limit; }
  function daysAgoDate(n) {
    if (n < 0) return "";
    var d = new Date(META.asOf + "T00:00:00");
    d.setDate(d.getDate() - n);
    return d.toISOString().slice(0, 10);
  }
  // Most recent activity across: task state change, latest comment, TC state change.
  function lastActivity(task) {
    var days = task.changedDays, on = task.changedOn;
    if (task.discDays >= 0 && (days < 0 || task.discDays < days)) {
      days = task.discDays; on = daysAgoDate(task.discDays);
    }
    if (task.state !== "Done" && task.tcStateChangeDays >= 0 && (days < 0 || task.tcStateChangeDays < days)) {
      days = task.tcStateChangeDays; on = daysAgoDate(task.tcStateChangeDays);
    }
    return { days: days, on: on };
  }

  var ACTIVITY = {
    w7:  function (t) { var d = t.state === "Done" ? t.closedDays : t.changedDays; return withinDays(d, 7)  || withinDays(t.discDays, 7)  || (t.state !== "Done" && withinDays(t.tcStateChangeDays, 7));  },
    w14: function (t) { var d = t.state === "Done" ? t.closedDays : t.changedDays; return withinDays(d, 14) || withinDays(t.discDays, 14) || (t.state !== "Done" && withinDays(t.tcStateChangeDays, 14)); },
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
  var ACT_LABEL = { w7: "Last 7d", c7: "Closed 7d", range: "Custom" };
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
      if (!opts.ignoreActivity && state.activity.size) {
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

  // ---- weekly status trend chart (SVG, pure DOM, no CDN) --------------------
  function renderStatusChart() {
    var wrap = $("statusChart");
    wrap.textContent = "";
    var rawHistory = META.statusHistory;
    if (!rawHistory || rawHistory.length < 1) {
      var msg = el("p", "muted");
      msg.style.cssText = "font-size:12px;margin:0";
      msg.textContent = "Trend chart appears after the first weekly run.";
      wrap.appendChild(msg);
      return;
    }

    // When exactly one person is selected, slice history through that person's
    // byPerson bucket. Falls back to team totals for any week that predates the
    // per-person data (old history entries won't have byPerson).
    var selPerson = (state.person.size === 1) ? Array.from(state.person)[0] : null;
    var history = rawHistory.map(function (r) {
      if (!selPerson) return r;
      var p = r.byPerson && r.byPerson[selPerson];
      return p ? { date: r.date, toDo: p.toDo || 0, inProgress: p.inProgress || 0, done: p.done || 0, bugs: p.totalBugs || 0, snapToDo: p.snapToDo || 0, snapInProgress: p.snapInProgress || 0, snapDone: p.snapDone || 0 }
               : { date: r.date, toDo: 0, inProgress: 0, done: 0, bugs: 0, snapToDo: 0, snapInProgress: 0, snapDone: 0 };
    });

    var snap = state.chartMode === "snapshot";
    var LINES = [
      { key: snap ? "snapDone"       : "done",        label: "Done",        color: "#22c55e", dash: ""    },
      { key: snap ? "snapInProgress" : "inProgress",  label: "In Progress", color: "#f59e0b", dash: "8,8" },
      { key: snap ? "snapToDo"       : "toDo",        label: "To Do",       color: "#3b82f6", dash: "6,8" },
      { key: "bugs",                                   label: "Bugs",        color: "#ef4444", dash: "3,8" }
    ];

    var W = Math.max(200, wrap.offsetWidth || 560);
    var H = 220;
    var PAD = { top: 22, right: 16, bottom: 30, left: 34 };
    var iW  = W - PAD.left - PAD.right;
    var iH  = H - PAD.top  - PAD.bottom;
    var n   = history.length;

    var allVals = [];
    history.forEach(function (r) {
      LINES.forEach(function (l) { allVals.push(r[l.key] || 0); });
    });
    var maxV = Math.max.apply(null, allVals) || 1;

    function xPos(i) { return PAD.left + (n < 2 ? iW / 2 : i * iW / (n - 1)); }
    function yPos(v) { return PAD.top + iH - (v / maxV) * iH; }

    var NS  = "http://www.w3.org/2000/svg";
    function svgEl(tag, attrs) {
      var e = document.createElementNS(NS, tag);
      Object.keys(attrs).forEach(function (k) { e.setAttribute(k, attrs[k]); });
      return e;
    }

    var svg = svgEl("svg", { viewBox: "0 0 " + W + " " + H, width: "100%", height: H });
    svg.style.overflow = "visible";
    svg.style.display  = "block";

    // Highlight band for the current week (last history entry).
    var curX     = xPos(n - 1);
    var bandHalf = Math.min(n < 2 ? iW / 2 : iW / (n - 1) / 2, 22);
    svg.appendChild(svgEl("rect", {
      x: curX - bandHalf, y: PAD.top - 8,
      width: bandHalf * 2, height: iH + 8,
      fill: "rgba(196,0,255,0.12)", rx: 3
    }));
    // "Now" label above the band.
    var nowLbl = document.createElementNS(NS, "text");
    nowLbl.setAttribute("x", curX); nowLbl.setAttribute("y", PAD.top - 11);
    nowLbl.setAttribute("text-anchor", "middle"); nowLbl.setAttribute("font-size", "9");
    nowLbl.setAttribute("font-weight", "600"); nowLbl.setAttribute("fill", "#c400ff");
    nowLbl.textContent = "Current";
    svg.appendChild(nowLbl);

    // Y gridlines + labels (4 steps)
    var steps = 4;
    for (var s = 0; s <= steps; s++) {
      var v  = Math.round(maxV * s / steps);
      var y  = yPos(v);
      svg.appendChild(svgEl("line", { x1: PAD.left, x2: PAD.left + iW, y1: y, y2: y,
        stroke: "var(--rule)", "stroke-width": 1 }));
      var lbl = document.createElementNS(NS, "text");
      lbl.setAttribute("x", PAD.left - 4); lbl.setAttribute("y", y + 4);
      lbl.setAttribute("text-anchor", "end"); lbl.setAttribute("font-size", "10");
      lbl.setAttribute("fill", "var(--ink-muted)"); lbl.textContent = v;
      svg.appendChild(lbl);
    }

    // X axis date labels — current week (last entry) is bold + blue.
    history.forEach(function (r, i) {
      var isCur = (i === n - 1);
      var lbl = document.createElementNS(NS, "text");
      lbl.setAttribute("x", xPos(i)); lbl.setAttribute("y", PAD.top + iH + 16);
      lbl.setAttribute("text-anchor", "middle"); lbl.setAttribute("font-size", "10");
      lbl.setAttribute("fill", isCur ? "#c400ff" : "var(--ink-muted)");
      if (isCur) lbl.setAttribute("font-weight", "600");
      lbl.textContent = (r.date || "").slice(5); // MM-DD
      svg.appendChild(lbl);
    });

    // Horizontal leader lines: from y-axis to each data point, per series.
    LINES.forEach(function (def) {
      history.forEach(function (r, i) {
        var v = r[def.key] || 0;
        if (v === 0) return;
        var y = yPos(v);
        svg.appendChild(svgEl("line", {
          x1: PAD.left, x2: xPos(i), y1: y, y2: y,
          stroke: def.color, "stroke-width": 0.75, opacity: 0.2
        }));
      });
    });

    // Lines and dots (polyline omitted when only 1 point - dots still render)
    LINES.forEach(function (def) {
      if (history.length > 1) {
        var pts = history.map(function (r, i) {
          return xPos(i) + "," + yPos(r[def.key] || 0);
        }).join(" ");
        var attrs = { points: pts, fill: "none", stroke: def.color, "stroke-width": 2, "stroke-linejoin": "round" };
        if (def.dash) attrs["stroke-dasharray"] = def.dash;
        svg.appendChild(svgEl("polyline", attrs));
      }

      history.forEach(function (r, i) {
        svg.appendChild(svgEl("circle", {
          cx: xPos(i), cy: yPos(r[def.key] || 0), r: 3, fill: def.color
        }));
      });
    });

    wrap.appendChild(svg);
  }

  // ---- weekly status dashboard (filter-responsive) ---------------------------
  function renderStatusDash() {
    var wrap = $("statusDash");
    wrap.textContent = "";

    // Base set: person/product/kind filtered but ignoring state and activity
    // filters so all three state tiles are always meaningful simultaneously.
    var allRows = visible({ ignoreState: true, ignoreActivity: true });

    // Weekly counts (live, filter-aware) using the same ACTIVITY functions
    // as the Activity section - w7 = state-changed/commented in last 7 days,
    // c7 = task closed in last 7 days.
    var toDo = 0, inPrg = 0, done = 0;
    var bugIdSet = new Set();
    allRows.forEach(function (t) {
      if (t.state === "To Do"       && ACTIVITY.w7(t)) toDo++;
      if (t.state === "In Progress" && ACTIVITY.w7(t)) inPrg++;
      if (ACTIVITY.c7(t))                               done++;
      if (t.bugIds) t.bugIds.split(",").filter(Boolean).forEach(function(b) { bugIdSet.add(b); });
    });
    var totalBugs = bugIdSet.size;

    // Total counts for sub-lines (all states, no activity filter).
    var totToDo = allRows.filter(function (t) { return t.state === "To Do"; }).length;
    var totInPrg = allRows.filter(function (t) { return t.state === "In Progress"; }).length;
    var totDone  = allRows.filter(function (t) { return t.state === "Done"; }).length;

    // New bugs this week.
    var selPerson = state.person.size === 1 ? Array.from(state.person)[0] : null;
    var noFilter  = !state.person.size && !state.product.size && !state.kind.size;
    $("statusScope").textContent = selPerson ? selPerson : "Team";
    var newBugs = 0;
    if (noFilter) {
      newBugs = META.statusCounts ? (META.statusCounts.bugs || 0) : 0;
    } else if (selPerson && META.statusHistory && META.statusHistory.length >= 2) {
      var prevH = META.statusHistory[META.statusHistory.length - 2];
      var prevP = prevH.byPerson && prevH.byPerson[selPerson];
      newBugs = prevP ? Math.max(0, totalBugs - (prevP.totalBugs || prevP.bugs || 0)) : 0;
    }

    var snap = state.chartMode === "snapshot";

    // In snapshot mode the card values are totals, not weekly-activity counts.
    var cardToDo = snap ? totToDo : toDo;
    var cardInPrg = snap ? totInPrg : inPrg;
    var cardDone  = snap ? totDone  : done;

    // Week-over-week deltas — use snapshot fields when in snapshot mode.
    var d = {};
    if (noFilter) {
      if (snap && META.statusHistory && META.statusHistory.length >= 2) {
        var prevSnap = META.statusHistory[META.statusHistory.length - 2];
        d = {
          toDo:       totToDo  - (prevSnap.snapToDo       || prevSnap.toDo       || 0),
          inProgress: totInPrg - (prevSnap.snapInProgress || prevSnap.inProgress || 0),
          done:       totDone  - ((prevSnap.snapToDo !== undefined ? (META.statusHistory.reduce(function(s,e){ return s + (e.done||0); }, 0) - (META.statusHistory[META.statusHistory.length-2] ? META.statusHistory.slice(0,-1).reduce(function(s,e){ return s+(e.done||0);},0) : 0)) : null))
        };
        d.done = null; // cumulative Done delta is noise — hide it in snapshot mode
      } else {
        d = META.statusDeltas || {};
      }
    } else if (selPerson && META.statusHistory && META.statusHistory.length >= 2) {
      var ph = META.statusHistory[META.statusHistory.length - 2];
      var pp = ph.byPerson && ph.byPerson[selPerson];
      if (pp) {
        d = snap ? {
          toDo:       totToDo  - (pp.snapToDo       || pp.toDo       || 0),
          inProgress: totInPrg - (pp.snapInProgress || pp.inProgress || 0),
          done:       null
        } : {
          toDo:       toDo - (pp.toDo       || 0),
          inProgress: inPrg - (pp.inProgress || 0),
          done:       done  - (pp.done       || 0)
        };
      }
    }

    function deltaEl(delta, goodDir) {
      if (delta === null || delta === undefined) {
        return el("div", "pct", META.statusHistory && META.statusHistory.length < 2 ? "first run" : GL.mdash);
      }
      if (delta === 0) return el("div", "pct", "same as last week");
      var sign   = delta > 0 ? "+" : "";
      var isGood = goodDir === 0 ? false : (delta > 0 ? goodDir > 0 : goodDir < 0);
      var isBad  = goodDir !== 0 && !isGood;
      return el("div", "pct" + (isGood ? " delta-good" : isBad ? " delta-bad" : ""),
                sign + delta + " from last week");
    }

    // goodDir: +1 = more is good, -1 = fewer is good, 0 = neutral
    var defs = [
      { key: "toDo",       label: "To Do",       color: "#3b82f6", goodDir: -1, val: cardToDo,  tot: snap ? null : totToDo  },
      { key: "inProgress", label: "In Progress",  color: "#f59e0b", goodDir:  0, val: cardInPrg, tot: snap ? null : totInPrg },
      { key: "done",       label: "Done",         color: "#22c55e", goodDir:  1, val: cardDone,  tot: snap ? null : totDone  },
      { key: "bugs",       label: "New Bugs",     color: "#ef4444", goodDir: -1, val: newBugs,   tot: totalBugs, noDelta: true }
    ];

    defs.forEach(function (def) {
      var tile = el("div", "tile");
      tile.style.setProperty("--tile-color", def.color);
      tile.appendChild(el("div", "value", fmt(def.val)));
      tile.appendChild(el("div", "label", def.label));
      if (def.noDelta) {
        tile.appendChild(el("div", "pct", "new this week  " + GL.dot + "  " + fmt(def.tot) + " total"));
      } else {
        var dEl = deltaEl(d[def.key], def.goodDir);
        // Append total to the same line if there is no delta text yet
        if (d[def.key] === null || d[def.key] === undefined) {
          dEl.textContent += "  " + fmt(def.tot) + " total";
        }
        tile.appendChild(dEl);
      }
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

    var win = item.workWindowDays || 0;
    var colour, over = false, days, pct, label;

    if (isDone) {
      // FREEZE at completion. workDaysLeft is measured from today, so leaving a
      // finished item on it would keep the meter ticking down for months and
      // eventually flip a long-since-delivered PBI to "50d over".
      //   target - closed = (target - today) + (today - closed)
      //                   = workDaysLeft + workClosedDays
      // Positive = finished early, negative = finished late.
      var atClose = (item.workClosedDays !== undefined && item.workClosedDays !== null && item.workClosedDays >= 0)
                  ? item.workDaysLeft + item.workClosedDays
                  : null;
      colour = "var(--st-na)";                       // history, never alarming
      pct = atClose !== null && win > 0 ? (win - atClose) / win : 1;
      label = "Done";
    } else {
      days = item.workDaysLeft;
      pct  = win > 0 ? (win - days) / win : (days < 0 ? 1.2 : 1);
      if (days < 0)       { colour = "var(--mtr-late)"; over = true; }
      else if (days <= 14) { colour = "var(--mtr-soon)"; }
      else                 { colour = "var(--mtr-ok)"; }
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
               (win > 0 ? "   (" + win + " working-day window, " + Math.round(pct * 100) + "% elapsed)" : "") +
               (isDone && item.closedOn ? "   closed " + item.closedOn : "");
    td.appendChild(wrap);
    return td;
  }

  function urgencyOf(t) {
    if (t.state === "Done") return "none";
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

  function taskFlags(task) {
    var f = [];
    if (task.state !== "Done") {
      if (!task.targetOn)                                                          f.push("No target date");
      if (task.kind === "Scripting" && task.remWorkNew === null)                   f.push("No remaining work");
      if (task.kind === "Execution" && task.state === "In Progress" && task.exec === 0 && task.cases > 0) f.push("Execution not started");
      if ((task.kind === "Scripting" || task.kind === "Execution") && task.cases === 0) f.push("No cases linked");
    }
    if (task.state === "Done" && task.kind === "Scripting" && task.cases > 0 && task.tcReady < task.cases) f.push("Not all TCs Ready");
    return f;
  }

  function loadGroups(rows) {
    var map = Object.create(null), order = [];
    rows.forEach(function (t) {
      var k = t[state.loadGroup] || "(none)";
      if (!map[k]) {
        map[k] = { name: k, total: 0, pbis: Object.create(null), soonest: null, soonestOn: "",
                   bugSet: Object.create(null),
                   capRate: 0, capRem: 0 };
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
      if (t.state !== "Done" && t.remWork > 0 && t.workDaysLeft < 99999) {
        var effDays = Math.max(t.workDaysLeft, 1);
        g.capRate += t.remWork / effDays;
        g.capRem  += t.remWork;
      }
      var pk = t.pbiId || "(no PBI parent)";
      if (!g.pbis[pk]) {
        g.pbis[pk] = { id: t.pbiId, title: t.pbiTitle || "(no PBI parent)", product: t.product || "", tasks: 0,
                       targetOn: t.targetOn, daysLeft: t.daysLeft, startOn: t.startOn,
                       windowDays: t.windowDays, workDaysLeft: t.workDaysLeft, workWindowDays: t.workWindowDays,
                       closedDays: null, closedOn: "", workClosedDays: null,
                       // Bugs hang off the PBI, so every task of a PBI reports
                       // the same set - take it, don't accumulate it.
                       bugs: t.bugIds ? t.bugIds.split(",").filter(Boolean).length : 0,
                       capRem: 0, capRate: 0 };
        STATES.forEach(function (s) { g.pbis[pk][s.key] = 0; });
        URGENCY.forEach(function (u) { g.pbis[pk]["u_" + u.key] = 0; });
      }
      g.pbis[pk].tasks++;
      if (g.pbis[pk][t.state] === undefined) g.pbis[pk][t.state] = 0;
      g.pbis[pk][t.state]++;
      g.pbis[pk]["u_" + urgencyOf(t)]++;
      if (taskFlags(t).length) g.pbis[pk].hasIncomplete = true;
      if (t.state !== "Done" && t.remWork > 0 && t.workDaysLeft < 99999) {
        var effDays = Math.max(t.workDaysLeft, 1);
        g.pbis[pk].capRem  += t.remWork;
        g.pbis[pk].capRate += t.remWork / effDays;
      }
      // A PBI finishes when its LAST task does, i.e. the most recent close =
      // the SMALLEST closedDays (fewest days ago).
      if (t.closedDays >= 0 && (g.pbis[pk].closedDays === null || t.closedDays < g.pbis[pk].closedDays)) {
        g.pbis[pk].closedDays     = t.closedDays;
        g.pbis[pk].closedOn       = t.closedOn;
        g.pbis[pk].workClosedDays = t.workClosedDays;
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
    // Auto-select first PBI the first time the table is shown (or re-shown).
    // Guard on loadTableView so the hidden initial render doesn't consume the flag.
    if (state.loadTableView && state.selectedPbiKey === null && !state._tableAutoSelected && groups.length > 0) {
      for (var _gi = 0; _gi < groups.length; _gi++) {
        var _gpbis = Object.keys(groups[_gi].pbis)
          .map(function (k) { return groups[_gi].pbis[k]; })
          .sort(function (a, b) {
            var _p = (a.product || "").localeCompare(b.product || "");
            return _p !== 0 ? _p : (a.title || "").localeCompare(b.title || "");
          });
        if (_gpbis.length > 0) {
          state.selectedPbiKey = _gpbis[0].id || "(no PBI parent)";
          state._tableAutoSelected = true;
          break;
        }
      }
    }
    var t = el("table"), thead = el("thead"), hr = el("tr");

    var hasAnyCap = groups.some(function (g) { return g.capRate > 0; });
    var totalCols = 4 + (hasAnyCap ? 1 : 0) + loadDims().length;

    // Target (index 2) is centred to match the meter cells meterCell() emits.
    [state.loadGroup === "assignee" ? "Person" : "Product", "PBI", "Product", "Target"].forEach(function (h, i) {
      hr.appendChild(el("th", i === 3 ? "ctr" : null, h));
    });
    if (hasAnyCap) {
      hr.appendChild(el("th", "cap-col-head", "Burn Rate"));
    }
    // State-breakdown columns (To Do / In Progress / Done) carry the legend
    // swatch so the table shares the same colour encoding as the bars above.
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
                       .sort(function (a, b) {
                         var p = (a.product || "").localeCompare(b.product || "");
                         return p !== 0 ? p : (a.title || "").localeCompare(b.title || "");
                       });

      // --- group (subtotal) row ---
      var sr = el("tr");
      var nameCell = el("td");
      nameCell.style.fontWeight = "600";
      nameCell.textContent = g.name;
      sr.appendChild(nameCell);
      sr.appendChild(el("td", "muted", fmt(pbis.length) + " PBI" + (pbis.length === 1 ? "" : "s")));
      sr.appendChild(el("td"));
      // Target stays BLANK on the group row - it is a PBI-level date.
      sr.appendChild(el("td", "ctr"));
      if (hasAnyCap) {
        if (g.capRate > 0) {
          var cls = g.capRate > 7 ? "over" : g.capRate > 4 ? "warn" : "ok";
          sr.appendChild(el("td", "num cap-rate " + cls, g.capRate.toFixed(1) + " h/day"));
        } else {
          sr.appendChild(el("td", "num muted", GL.mdash));
        }
      }
      loadDims().forEach(function (d) { sr.appendChild(el("td", "num", fmt(g[d.field]))); });
      tb.appendChild(sr);

      // --- PBI rows ---
      pbis.forEach(function (p) {
        var pbiKey = p.id || "(no PBI parent)";
        var isSelected = state.selectedPbiKey === pbiKey;
        var tr = el("tr", "pbi-row" + (isSelected ? " pbi-selected" : ""));
        tr.tabIndex = 0;
        tr.title = isSelected ? "Click to deselect (show all tasks)" : "Click to filter tasks to this PBI";
        tr.setAttribute("aria-selected", isSelected ? "true" : "false");
        (function (key) {
          function toggle() {
            state.selectedPbiKey = state.selectedPbiKey === key ? null : key;
            render();
          }
          tr.addEventListener("click", toggle);
          tr.addEventListener("keydown", function (e) {
            if (e.key === "Enter" || e.key === " ") { e.preventDefault(); toggle(); }
          });
        })(pbiKey);
        tr.appendChild(el("td", null, ""));
        var c = el("td");
        var titleWrap = el("div");
        titleWrap.appendChild(document.createTextNode(p.title));
        if (p.hasIncomplete) {
          var ast = el("span", null, " *");
          ast.style.cssText = "color:#c0392b;font-weight:700;";
          ast.title = "One or more tasks under this PBI have incomplete items";
          titleWrap.appendChild(ast);
        }
        if (p.id && META.tfsBase) {
          var lnk = document.createElement("a");
          lnk.href = META.tfsBase + p.id;
          lnk.target = "_blank";
          lnk.rel = "noopener";
          lnk.className = "tfs-link";
          lnk.title = "Open PBI " + p.id + " in TFS";
          lnk.innerHTML = '<svg viewBox="0 0 12 12" width="11" height="11" fill="currentColor" aria-hidden="true"><path d="M1 1h4v1H2v8h8V7h1v4H1V1zm5 0h4v4h-1V2.7L5.4 6.3l-.7-.7L8.3 2H6V1z"/></svg>';
          lnk.addEventListener("click", function (e) { e.stopPropagation(); });
          titleWrap.appendChild(lnk);
        }
        c.appendChild(titleWrap);
        if (p.id) c.appendChild(el("div", "meta-line", "PBI " + p.id));
        tr.appendChild(c);
        tr.appendChild(el("td", p.product ? null : "muted", p.product || GL.mdash));
        // Target belongs on the PBI row - that is the level the date is set at.
        // A PBI counts as done only when every one of its tasks is.
        tr.appendChild(meterCell(p, p["Done"] === p.tasks));
        if (hasAnyCap) {
          var isDone = p["Done"] === p.tasks;
          if (!isDone && p.capRate > 0) {
            var pcls = p.capRate > 7 ? "over" : p.capRate > 4 ? "warn" : "ok";
            tr.appendChild(el("td", "num cap-rate " + pcls, p.capRate.toFixed(1)));
          } else {
            tr.appendChild(el("td", "num muted", GL.mdash));
          }
        }
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
    { key: "changedDays", label: "Last change", num: true },
    // num:true keeps it sortable and filterable with >/< expressions;
    // align:"center" only changes where it sits in the column.
    { key: "daysLeft",    label: "Target",      num: true, align: "center" },
    { key: "disc",        label: "",             num: false, noSort: true, noFilter: true, align: "center" },
    { key: "bugs",        label: "Bugs",         num: true },
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
    if (key === "bugs")        return task.bugIds ? task.bugIds.split(",").filter(Boolean).length : 0;
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
        // Also search by task ID and PBI so filters on those work even though
        // neither is visible as its own column.
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
      var th = el("th", (c.align === "center" ? "ctr " : (c.num ? "num " : "")) + (c.noSort ? "" : "sortable"));
      var span = el("span", null, c.label);
      labels[c.key] = span;
      th.appendChild(span);
      if (!c.noSort) {
        th.tabIndex = 0;
        function doSort() {
          if (state.sortKey === c.key) state.sortDir = -state.sortDir;
          else { state.sortKey = c.key; state.sortDir = c.num ? -1 : 1; }
          renderTasks(lastRows);
        }
        th.addEventListener("click", doSort);
        th.addEventListener("keydown", function (e) { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); doSort(); } });
      }
      hr.appendChild(th);
    });
    thead.appendChild(hr);

    var fr = el("tr", "filterrow");
    COLS.forEach(function (c) {
      var cell = el("th", c.align === "center" ? "ctr" : (c.num ? "num" : null)), input;
      if (c.noFilter) { fr.appendChild(cell); return; }
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
    var clear = el("button", "icon-btn");
    clear.type = "button";
    clear.title = "Clear column filters";
    clear.innerHTML = '<svg viewBox="0 0 24 24" style="width:18px;height:18px;stroke:currentColor;fill:none;stroke-width:1.75;stroke-linecap:round;stroke-linejoin:round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>';
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

    // When in table view with a PBI selected, scope to that PBI's tasks first.
    var baseRows = (state.loadTableView && state.selectedPbiKey)
      ? rows.filter(function (t) { return (t.pbiId || "(no PBI parent)") === state.selectedPbiKey; })
      : rows;

    var filtered = baseRows.filter(passesColFilters);
    var sorted = filtered.slice().sort(function (a, b) {
      var x = a[state.sortKey], y = b[state.sortKey];
      if (typeof x === "number" && typeof y === "number") return (x - y) * state.sortDir;
      return String(x).localeCompare(String(y)) * state.sortDir;
    });

    taskUI.note.textContent = "";
    if (state.loadTableView && state.selectedPbiKey) {
      var pbiTitle = state.selectedPbiKey === "(no PBI parent)" ? "(no PBI parent)" : "";
      for (var ni = 0; ni < baseRows.length; ni++) {
        if (baseRows[ni].pbiTitle) { pbiTitle = baseRows[ni].pbiTitle; break; }
      }
      var countStr = filtered.length === baseRows.length
        ? fmt(baseRows.length) + " task" + (baseRows.length === 1 ? "" : "s")
        : "Showing " + fmt(filtered.length) + " of " + fmt(baseRows.length) + " tasks";
      taskUI.note.appendChild(document.createTextNode(countStr));
      var clrBtn = document.createElement("button");
      clrBtn.type = "button";
      clrBtn.title = "Show all tasks";
      clrBtn.style.cssText = "font: inherit; margin-left: 6px; padding: 0 5px; line-height: 1.4; vertical-align: middle; cursor: pointer;";
      clrBtn.textContent = GL.times;
      clrBtn.addEventListener("click", function () { state.selectedPbiKey = null; render(); });
      taskUI.note.appendChild(clrBtn);
    } else {
      taskUI.note.textContent = filtered.length === rows.length
        ? fmt(rows.length) + " task" + (rows.length === 1 ? "" : "s")
        : "Showing " + fmt(filtered.length) + " of " + fmt(rows.length) + " tasks (column filters active)";
    }

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

      // --- incomplete flags (active tasks only) ---
      var flags = taskFlags(task);
      if (flags.length) tr.className = "flag-row";

      var td = el("td");
      var titleWrap = el("div", "tasktitle");
      titleWrap.appendChild(document.createTextNode(task.title));
      if (META.tfsBase) {
        var tLnk = document.createElement("a");
        tLnk.href = META.tfsBase + task.id;
        tLnk.target = "_blank";
        tLnk.rel = "noopener";
        tLnk.className = "tfs-link";
        tLnk.title = "Open task " + task.id + " in TFS";
        tLnk.innerHTML = '<svg viewBox="0 0 12 12" width="11" height="11" fill="currentColor" aria-hidden="true"><path d="M1 1h4v1H2v8h8V7h1v4H1V1zm5 0h4v4h-1V2.7L5.4 6.3l-.7-.7L8.3 2H6V1z"/></svg>';
        tLnk.addEventListener("click", function (e) { e.stopPropagation(); });
        titleWrap.appendChild(tLnk);
      }
      td.appendChild(titleWrap);
      var meta = "#" + task.id;
      if (task.testers) meta += SEP + "tested by " + task.testers;
      td.appendChild(el("div", "meta-line", meta));
      if (flags.length) {
        var chips = el("div", "flag-chips");
        flags.forEach(function (f) { chips.appendChild(el("span", "flag-chip", f)); });
        td.appendChild(chips);
      }
      tr.appendChild(td);

      var kindTd = el("td");
      kindTd.appendChild(el("span", "pill", task.kind));
      tr.appendChild(kindTd);

      var stTd = el("td");
      stTd.appendChild(el("span", "pill", task.state));
      tr.appendChild(stTd);

      tr.appendChild(el("td", "nowrap", task.assignee));

      var chg = el("td", "num nowrap");
      var act = lastActivity(task);
      if (act.days < 0) {
        chg.className += " muted";
        chg.textContent = GL.mdash;
      } else {
        chg.appendChild(el("div", null, act.on));
        var age = act.days === 0 ? "today"
                : act.days === 1 ? "1 day ago"
                : act.days + " days ago";
        chg.appendChild(el("div", "meta-line", age));
      }
      tr.appendChild(chg);

      tr.appendChild(meterCell(task, task.state === "Done"));

      var discTd = el("td", "ctr");
      if (task.taskDisc || task.pbiDisc) {
        var dBtn = document.createElement("button");
        dBtn.className = "disc-btn";
        dBtn.title = "View discussion";
        dBtn.setAttribute("aria-label", "Discussion for " + task.title);
        dBtn.innerHTML = '<svg viewBox="0 0 16 16" width="13" height="13" fill="currentColor" aria-hidden="true"><path d="M2 1h12a1 1 0 011 1v8a1 1 0 01-1 1H9l-3 3-3-3H2a1 1 0 01-1-1V2a1 1 0 011-1z"/></svg>';
        (function (t) { dBtn.addEventListener("click", function (e) { e.stopPropagation(); showDisc(t); }); })(task);
        discTd.appendChild(dBtn);
      } else {
        discTd.className += " muted";
        discTd.textContent = GL.mdash;
      }
      tr.appendChild(discTd);

      var bugCount = task.bugIds ? task.bugIds.split(",").filter(Boolean).length : 0;
      var bugTd = el("td", "num" + (bugCount ? "" : " muted"));
      bugTd.appendChild(document.createTextNode(fmt(bugCount)));
      if (task.bugDelta > 0) {
        var dlt = el("span", "bug-delta", "+" + task.bugDelta);
        dlt.title = task.bugDelta + " new bug" + (task.bugDelta > 1 ? "s" : "") + " since last report";
        bugTd.appendChild(dlt);
      }
      tr.appendChild(bugTd);

      tr.appendChild(el("td", "num", fmt(task.cases)));

      // A Scripting task links test cases, not a plan - it has no execution
      // outcomes of its own, and a 0 would read as "all failed to run".
      // For Scripting-kind tasks we show test case ADO state (Ready/Design)
      // instead: a colspan=6 cell spanning the outcome columns. For Other-
      // kind tasks with exec=0, we still show six dashes.
      // NOTE: colspan=6 means this branch emits ONE cell spanning 6 columns,
      // not 6 cells - the total column count is still correct.
      if (task.exec === 0) {
        if (task.kind === "Scripting" && (task.tcReady + task.tcDesign) > 0) {
          var sc = el("td", "scripting-status");
          sc.colSpan = 6;
          var pct = task.cases > 0 ? Math.round(task.tcReady / task.cases * 100) : 0;
          var parts = [
            '<span class="tc-ready">' + fmt(task.tcReady) + "/" + fmt(task.cases) + " Ready (" + pct + "%)</span>",
            '<span class="tc-design">' + fmt(task.tcDesign) + " Design</span>"
          ];
          sc.innerHTML = parts.join('<span class="tc-sep">' + GL.dot + "</span>");
          tr.appendChild(sc);
        } else {
          ["points", "passed", "failed", "blocked", "na", "never"].forEach(function () {
            tr.appendChild(el("td", "num muted", GL.mdash));
          });
        }
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

    renderStatusChart();
    renderStatusDash();
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
      ? "A PBI appears only if it has at least one matching task. Bar length = task count; segments show calendar days remaining to the PBI target date (not effort)."
      : "A PBI appears only if it has at least one matching task. Bar length = task count; segments show workflow state.";
    $("loadWrap").classList.toggle("hidden", state.loadTableView);
    $("loadTable").classList.toggle("hidden", !state.loadTableView);
    $("loadTableBtn").setAttribute("aria-pressed", state.loadTableView ? "true" : "false");

    renderTasks(rows);

    // Fixed title; the "Group by" control beside it already states the
    // grouping, so restating it here only made the heading jitter on change.
    $("chartTitle").textContent = "Test Points";

    var exceptBtn = $("exceptBtn"), exceptBadge = $("exceptBadge");
    if (sum.mistakes > 0) {
      $("mistakeCount").textContent = fmt(sum.mistakes);
      exceptBadge.textContent = sum.mistakes > 99 ? "99+" : sum.mistakes;
      exceptBtn.hidden = false;
    } else {
      exceptBtn.hidden = true;
      $("exceptPopover").classList.remove("open");
    }

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

  // Auto-switch Assigned Work Items to table view when a person is selected,
  // back to chart view when selection is cleared.
  $("fPerson").querySelector(".multi-panel").addEventListener("change", function () {
    var hasPerson = state.person.size > 0;
    if (state.loadTableView !== hasPerson) {
      state.loadTableView = hasPerson;
      if (!state.loadTableView) {
        state.selectedPbiKey = null;
        state._tableAutoSelected = false;
      } else {
        state._tableAutoSelected = false;
      }
      $("loadTableBtn").setAttribute("aria-pressed", hasPerson ? "true" : "false");
      render();
    }
  });

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
    if (state.activity.has("c7")) {
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

  $("chartModeActivity").addEventListener("click", function () {
    state.chartMode = "activity";
    $("chartModeActivity").classList.add("chart-mode-active");
    $("chartModeSnapshot").classList.remove("chart-mode-active");
    renderStatusChart();
    renderStatusDash();
  });
  $("chartModeSnapshot").addEventListener("click", function () {
    state.chartMode = "snapshot";
    $("chartModeSnapshot").classList.add("chart-mode-active");
    $("chartModeActivity").classList.remove("chart-mode-active");
    renderStatusChart();
    renderStatusDash();
  });

  $("fGroup").addEventListener("change",   function (e) { state.group = e.target.value; render(); });
  $("fText").addEventListener("input",     function (e) { state.text = e.target.value; render(); });
  $("fReset").addEventListener("click", function () {
    ["fPerson","fProduct","fKind"].forEach(function (id) {
      $(id).querySelectorAll("input[type=checkbox]").forEach(function (cb) { cb.checked = false; });
    });
    $("fState").querySelectorAll("input[type=checkbox]").forEach(function (cb) {
      cb.checked = (cb.value === "In Progress" || cb.value === "Done");
    });
    $("fActivity").querySelectorAll("input[type=checkbox]").forEach(function (cb) { cb.checked = cb.value === "w7"; });
    $("actRange").hidden = true;
    $("rangeStart").value = ""; $("rangeEnd").value = "";
    state.person = new Set(); state.product = new Set();
    state.state = new Set(["In Progress", "Done"]);
    state.kind = new Set(); state.activity = new Set(["w7"]);
    state.activityRange = { start: "", end: "" }; state.text = "";
    $("fText").value = ""; $("fGroup").value = "assignee"; state.group = "assignee";
    $("fPerson").querySelector(".multi-label").textContent   = "All people";
    $("fProduct").querySelector(".multi-label").textContent  = "All products";
    $("fState").querySelector(".multi-label").textContent    = "2 selected";
    $("fKind").querySelector(".multi-label").textContent     = "All kinds";
    $("fActivity").querySelector(".multi-label").textContent = "Last 7d";
    render();
  });
  function collapseIcon(expanded) {
    return expanded
      ? '<svg viewBox="0 0 24 24" style="width:18px;height:18px;stroke:currentColor;fill:none;stroke-width:1.75;stroke-linecap:round;stroke-linejoin:round"><polyline points="18 15 12 9 6 15"/></svg>'
      : '<svg viewBox="0 0 24 24" style="width:18px;height:18px;stroke:currentColor;fill:none;stroke-width:1.75;stroke-linecap:round;stroke-linejoin:round"><polyline points="6 9 12 15 18 9"/></svg>';
  }
  function wireCollapse(btnId, bodyId) {
    var btn = $(btnId), body = $(bodyId);
    btn.addEventListener("click", function (e) {
      e.stopPropagation();
      var open = body.hidden;
      body.hidden = !open;
      btn.setAttribute("aria-expanded", open ? "true" : "false");
      btn.innerHTML = collapseIcon(open);
    });
  }
  wireCollapse("kpisToggle", "kpisBody");
  wireCollapse("chartToggle", "chartBody");

  // Activity drawer
  function openDrawer()  { $("activityDrawer").classList.add("open");    $("drawerOverlay").classList.add("open"); }
  function closeDrawer() { $("activityDrawer").classList.remove("open"); $("drawerOverlay").classList.remove("open"); }
  $("activityBtn").addEventListener("click", function () {
    $("activityDrawer").classList.contains("open") ? closeDrawer() : openDrawer();
  });
  $("drawerClose").addEventListener("click", closeDrawer);
  $("drawerOverlay").addEventListener("click", closeDrawer);
  document.addEventListener("keydown", function (e) { if (e.key === "Escape") { closeDrawer(); $("exceptPopover").classList.remove("open"); } });

  // Exceptions popover — position it below the button on click
  $("exceptBtn").addEventListener("click", function (e) {
    var pop = $("exceptPopover");
    if (pop.classList.contains("open")) { pop.classList.remove("open"); return; }
    var r = e.currentTarget.getBoundingClientRect();
    pop.style.top  = (r.bottom + 6) + "px";
    pop.style.right = (window.innerWidth - r.right) + "px";
    pop.classList.add("open");
    e.stopPropagation();
  });
  document.addEventListener("click", function () { $("exceptPopover").classList.remove("open"); });

  $("tableBtn").addEventListener("click", function () { state.tableView = !state.tableView; render(); });
  $("fLoadGroup").addEventListener("change", function (e) { state.loadGroup = e.target.value; render(); });
  $("fLoadColor").addEventListener("change", function (e) {
    state.loadColor = e.target.value;
    renderLoadLegend();          // the legend changes with the dimension
    render();
  });
  $("loadTableBtn").addEventListener("click", function () {
    state.loadTableView = !state.loadTableView;
    if (!state.loadTableView) {
      state.selectedPbiKey = null;
      state._tableAutoSelected = false;
    } else {
      state._tableAutoSelected = false;
    }
    render();
  });

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

  // ---- discussion dialog ---------------------------------------------------
  function showDisc(task) {
    $("discTitle").textContent = task.title;
    var content = $("discContent");
    content.textContent = "";

    function section(who, discs) {
      var sec = el("div", "disc-section");
      sec.appendChild(el("div", "disc-who", who));
      var discArr = discs ? (Array.isArray(discs) ? discs : [discs]) : [];
      if (discArr.length) {
        // Newest first
        var reversed = discArr.slice().reverse();
        reversed.forEach(function (disc) {
          if (!disc || !disc.html) return;
          var entry = el("div", "disc-entry");
          var byLine = (disc.author || "") + (disc.date ? "  " + GL.dot + "  " + disc.date : "");
          if (byLine.trim()) entry.appendChild(el("div", "disc-meta", byLine));
          var body = el("div", "disc-body");
          body.innerHTML = disc.html;
          body.querySelectorAll("[style]").forEach(function(n) { n.removeAttribute("style"); });
          body.querySelectorAll("[color]").forEach(function(n) { n.removeAttribute("color"); });
          entry.appendChild(body);
          sec.appendChild(entry);
        });
      } else {
        sec.appendChild(el("div", "disc-none", "No discussion entries."));
      }
      return sec;
    }

    content.appendChild(section("Task #" + task.id, task.taskDisc));

    $("discOverlay").classList.remove("hidden");
    $("discClose").focus();
  }

  function closeDisc() { $("discOverlay").classList.add("hidden"); }

  $("discClose").addEventListener("click", closeDisc);
  $("discOverlay").addEventListener("click", function (e) {
    if (e.target === $("discOverlay")) closeDisc();
  });
  document.addEventListener("keydown", function (e) {
    if (e.key === "Escape" && !$("discOverlay").classList.contains("hidden")) {
      e.preventDefault(); closeDisc();
    }
  });

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

    # Persist current bug counts for next run's delta detection.
    $newCounts = [ordered]@{}
    foreach ($t in $tasks) {
        $newCounts[$t.id] = if ($t.bugIds) { ($t.bugIds -split ',').Count } else { 0 }
    }
    $newCounts | ConvertTo-Json -Compress | Set-Content -LiteralPath $bugCountsPrevPath -Encoding UTF8
    Write-Host ("Bug counts snapshot written ({0} tasks)" -f $newCounts.Count)

    # Persist rolling status history for the trend chart and delta tiles.
    $updatedHistory | ConvertTo-Json -Depth 8 -Compress | Set-Content -LiteralPath $statusHistPath -Encoding UTF8

    if ($Show) { Start-Process $OutputPath }
}
catch {
    Write-Error "Failed: $($_.Exception.Message)"
    exit 1
}
