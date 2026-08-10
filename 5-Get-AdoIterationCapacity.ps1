<#
.SYNOPSIS
    Pulls per-person capacity (hours/day, days off) for an iteration/sprint,
    so you can see who has spare bandwidth this week.

.DESCRIPTION
    This is a different data source than the other 4 scripts in this folder -
    capacity isn't a work item field, it's set per-team, per-iteration under
    Azure Boards > Sprints > Capacity. This script:
      1. Optionally lists the teams in the project (-ListTeams), since
         capacity is scoped to a team, not the whole project.
      2. Optionally lists a team's iterations with their start/finish dates
         (-ListIterations), so you can find the right iteration ID.
      3. Pulls capacity for a specific iteration (by -IterationPath match,
         or the team's current iteration if you don't specify one) and
         computes each person's total available hours for that iteration
         (capacity/day x working days, minus days off).

    Run with -ListTeams first if you don't already know the exact team name
    - "spare bandwidth" only makes sense once we're looking at the right
      team's board. Shares the same encrypted PAT cache as the other
      scripts in this folder.

.PARAMETER Organization
    Base URL of the ADO/TFS server + collection. Default: https://tfs.deltek.com/tfs/Deltek

.PARAMETER Project
    Project name. Default: QEAutomation

.PARAMETER Team
    Team name (capacity is set per-team). If omitted, defaults to
    "<Project> Team" - Azure DevOps' default team name - but run
    -ListTeams first to confirm that's actually right for this project.

.PARAMETER ListTeams
    Lists all teams in the project and exits. Run this first.

.PARAMETER ListIterations
    Lists all iterations for -Team (with ID, path, start/finish dates) and exits.

.PARAMETER IterationPath
    Friendly iteration path to match, e.g. "QEAutomation\Costpoint\Sprint 12".
    If omitted, uses the team's current iteration.

.PARAMETER OutputPath
    Output CSV path. Default: iteration_capacity.csv

.PARAMETER Pat
.PARAMETER ApiVersion
    Optional - if not specified, probes several versions automatically.
.PARAMETER ResetPat

.EXAMPLE
    .\Get-AdoIterationCapacity.ps1 -ListTeams

.EXAMPLE
    .\Get-AdoIterationCapacity.ps1 -Team "QEAutomation Team" -ListIterations

.EXAMPLE
    .\Get-AdoIterationCapacity.ps1 -Team "QEAutomation Team" -IterationPath "QEAutomation\Costpoint\Sprint 12" -OutputPath csv\iteration_capacity.csv
#>

[CmdletBinding()]
param(
    [string]$Organization = "https://tfs.deltek.com/tfs/Deltek",
    [string]$Project = "QEAutomation",
    [string]$Team,
    [switch]$ListTeams,
    [switch]$ListIterations,
    [string]$IterationPath,
    [string]$OutputPath = "iteration_capacity.csv",
    [string]$Pat,
    [string]$ApiVersion,
    [switch]$ResetPat
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = "Stop"
$BaseUrl = $Organization.TrimEnd('/')

$CredentialDir  = Join-Path $env:LOCALAPPDATA "AdoTestPlanExtractor"
$CredentialFile = Join-Path $CredentialDir "pat.dat"

function Get-SavedPat {
    if (-not (Test-Path $CredentialFile)) { return $null }
    try {
        $secure = Get-Content -Path $CredentialFile -ErrorAction Stop | ConvertTo-SecureString
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try { return [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    } catch {
        Write-Warning "Could not read saved PAT. Ignoring it."
        return $null
    }
}

function Save-Pat {
    param([string]$PlainPat)
    if (-not (Test-Path $CredentialDir)) { New-Item -ItemType Directory -Path $CredentialDir -Force | Out-Null }
    $PlainPat | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Set-Content -Path $CredentialFile
    Write-Host "PAT saved (encrypted, this Windows account only) to $CredentialFile"
}

function Get-SecretOrPrompt {
    param([string]$Value, [string]$Label)
    if (-not [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    $saved = Get-SavedPat
    if (-not [string]::IsNullOrWhiteSpace($saved)) {
        Write-Host "Using saved PAT from $CredentialFile (run with -ResetPat to clear it)."
        return $saved
    }
    if (-not [string]::IsNullOrWhiteSpace($env:AZURE_DEVOPS_PAT)) { return $env:AZURE_DEVOPS_PAT }
    $secure = Read-Host -Prompt $Label -AsSecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    if (-not [string]::IsNullOrWhiteSpace($plain)) { Save-Pat -PlainPat $plain }
    return $plain
}

if ($ResetPat -and (Test-Path $CredentialFile)) {
    Remove-Item -Path $CredentialFile -Force
    Write-Host "Cleared saved PAT. You will be prompted for a new one."
}
$Pat = Get-SecretOrPrompt -Value $Pat -Label "Personal Access Token (press Enter to use your Windows login instead)"

$UseWindowsAuth = [string]::IsNullOrWhiteSpace($Pat)
if ($UseWindowsAuth) {
    Write-Host "No PAT provided - using your current Windows login (NTLM/Kerberos)."
    $AuthHeader = @{}
} else {
    $AuthHeader = @{ Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Pat")) }
}

if (-not $Team) {
    $Team = "$Project Team"
    Write-Host "No -Team specified, guessing default team name: '$Team'. Run with -ListTeams to see the real list if this is wrong."
}

$CandidateVersions = if ($ApiVersion) { @($ApiVersion) } else { @("7.1", "7.0", "6.0", "5.1", "5.0", "4.1", "3.2", "3.0") }

function Invoke-AdoGetProbed {
    # Tries each candidate API version against $Uri until one returns HTTP 200
    # with parseable JSON. Returns the parsed object, or $null + prints diagnostics.
    param([string]$Uri, [string]$QueryExtra = "")
    foreach ($v in $CandidateVersions) {
        $fullUri = $Uri + "?api-version=$v" + $QueryExtra
        try {
            if ($UseWindowsAuth) {
                $resp = Invoke-WebRequest -Uri $fullUri -Method Get -UseDefaultCredentials -UseBasicParsing
            } else {
                $resp = Invoke-WebRequest -Uri $fullUri -Method Get -Headers $AuthHeader -UseBasicParsing
            }
            Write-Host "  [$v] GET $fullUri -> HTTP $($resp.StatusCode)"
            if ($resp.StatusCode -eq 200) {
                try {
                    $obj = $resp.Content | ConvertFrom-Json
                    return $obj
                } catch {
                    continue
                }
            }
        } catch {
            $status = $null
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            Write-Host "  [$v] GET $fullUri -> HTTP $status ($($_.Exception.Message))"
        }
    }
    return $null
}

$encodedProject = [Uri]::EscapeDataString($Project)
$encodedTeam    = [Uri]::EscapeDataString($Team)

# ---------------- -ListTeams ----------------
if ($ListTeams) {
    Write-Host "----------------------------------------"
    Write-Host "Teams in project '$Project'"
    Write-Host "----------------------------------------"
    $uri = "$BaseUrl/_apis/projects/$encodedProject/teams"
    $result = Invoke-AdoGetProbed -Uri $uri
    if (-not $result -or -not $result.value) {
        Write-Warning "Could not retrieve team list. See the GET attempts above for diagnostics."
        return
    }
    $result.value | Select-Object name, id, description | Format-Table -AutoSize
    return
}

# ---------------- -ListIterations ----------------
if ($ListIterations) {
    Write-Host "----------------------------------------"
    Write-Host "Iterations for team '$Team' in project '$Project'"
    Write-Host "----------------------------------------"
    $uri = "$BaseUrl/$encodedProject/$encodedTeam/_apis/work/teamsettings/iterations"
    $result = Invoke-AdoGetProbed -Uri $uri
    if (-not $result -or -not $result.value) {
        Write-Warning "Could not retrieve iterations. If the team name is wrong, run -ListTeams first."
        return
    }
    $result.value | ForEach-Object {
        [PSCustomObject]@{
            Id         = $_.id
            Name       = $_.name
            Path       = $_.path
            StartDate  = $_.attributes.startDate
            FinishDate = $_.attributes.finishDate
            TimeFrame  = $_.attributes.timeFrame
        }
    } | Format-Table -AutoSize
    return
}

# ---------------- Resolve target iteration ----------------
Write-Host "----------------------------------------"
Write-Host "Resolving target iteration for team '$Team'"
Write-Host "----------------------------------------"
$iterUri = "$BaseUrl/$encodedProject/$encodedTeam/_apis/work/teamsettings/iterations"

$targetIteration = $null
if ($IterationPath) {
    $result = Invoke-AdoGetProbed -Uri $iterUri
    if ($result -and $result.value) {
        $targetIteration = $result.value | Where-Object { $_.path -eq $IterationPath } | Select-Object -First 1
    }
    if (-not $targetIteration) {
        Write-Warning "No iteration found matching path '$IterationPath'. Run -ListIterations to see valid paths."
        return
    }
} else {
    $result = Invoke-AdoGetProbed -Uri $iterUri -QueryExtra "&`$timeframe=current"
    if ($result -and $result.value -and $result.value.Count -gt 0) {
        $targetIteration = $result.value[0]
    }
    if (-not $targetIteration) {
        Write-Warning "No 'current' iteration found for this team. Pass -IterationPath explicitly (see -ListIterations)."
        return
    }
}

Write-Host "Target iteration: $($targetIteration.name)  [$($targetIteration.path)]"
Write-Host "  Start:  $($targetIteration.attributes.startDate)"
Write-Host "  Finish: $($targetIteration.attributes.finishDate)"

# ---------------- Pull capacities ----------------
Write-Host "----------------------------------------"
Write-Host "Capacities"
Write-Host "----------------------------------------"
$capUri = "$BaseUrl/$encodedProject/$encodedTeam/_apis/work/teamsettings/iterations/$($targetIteration.id)/capacities"
$capResult = Invoke-AdoGetProbed -Uri $capUri
if (-not $capResult -or -not $capResult.value) {
    Write-Warning "Could not retrieve capacities for this iteration. Paste the GET attempts above back for diagnosis."
    return
}

$startDate = [DateTime]$targetIteration.attributes.startDate
$finishDate = [DateTime]$targetIteration.attributes.finishDate

function Get-WorkingDayCount {
    param([DateTime]$Start, [DateTime]$End)
    $count = 0
    $d = $Start.Date
    while ($d -le $End.Date) {
        if ($d.DayOfWeek -ne [DayOfWeek]::Saturday -and $d.DayOfWeek -ne [DayOfWeek]::Sunday) { $count++ }
        $d = $d.AddDays(1)
    }
    return $count
}

$rows = @()
foreach ($member in $capResult.value) {
    $name = $member.teamMember.displayName
    $capPerDay = 0
    if ($member.activities) {
        $capPerDay = ($member.activities | Measure-Object -Property capacityPerDay -Sum).Sum
    }

    $daysOffCount = 0
    if ($member.daysOff) {
        foreach ($off in $member.daysOff) {
            $offStart = [DateTime]$off.start
            $offEnd = [DateTime]$off.end
            $clampedStart = if ($offStart -gt $startDate) { $offStart } else { $startDate }
            $clampedEnd = if ($offEnd -lt $finishDate) { $offEnd } else { $finishDate }
            if ($clampedStart -le $clampedEnd) {
                $daysOffCount += (Get-WorkingDayCount -Start $clampedStart -End $clampedEnd)
            }
        }
    }

    $totalWorkingDays = Get-WorkingDayCount -Start $startDate -End $finishDate
    $availableDays = [Math]::Max(0, $totalWorkingDays - $daysOffCount)
    $totalHours = $availableDays * $capPerDay

    $rows += [PSCustomObject]@{
        TeamMember          = $name
        Iteration           = $targetIteration.name
        IterationStart      = $targetIteration.attributes.startDate
        IterationFinish     = $targetIteration.attributes.finishDate
        CapacityPerDayHours = $capPerDay
        TotalWorkingDays    = $totalWorkingDays
        DaysOff             = $daysOffCount
        AvailableDays       = $availableDays
        TotalAvailableHours = $totalHours
    }
}

$rows | Sort-Object TotalAvailableHours -Descending | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "Wrote $($rows.Count) row(s) to $OutputPath"
$rows | Sort-Object TotalAvailableHours -Descending | Format-Table -AutoSize
