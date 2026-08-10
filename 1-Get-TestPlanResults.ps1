<#
.SYNOPSIS
    Pulls every Test Suite under an Azure DevOps Test Plan, with the current
    outcome of every Test Case in each suite, into one CSV (or Excel) file.

.DESCRIPTION
    Nothing about the org, project, plan, or credentials is hardcoded. Pass
    them as parameters, set AZURE_DEVOPS_PAT as an environment variable, or
    just run the script with no parameters and answer the prompts.

.PARAMETER Organization
    Either a cloud ADO organization name (the part after dev.azure.com/), or
    the full base URL of an on-prem Azure DevOps Server / TFS collection,
    e.g. https://tfs.deltek.com/tfs/Deltek

.PARAMETER Project
    Azure DevOps project name (e.g. QEAutomation)

.PARAMETER Plan
    Test Plan ID (number) or name (exact or partial match)

.PARAMETER RootSuite
    Optional. Only pull this suite and its children instead of the whole plan
    (e.g. "Costpoint").

.PARAMETER OutputPath
    Output file path. .csv or .xlsx decided by extension.
    Default: test_plan_results.csv

.PARAMETER Detailed
    Switch. Also fetch last-run date/duration per test case (slower - one
    extra API call per test case that has a result).

.PARAMETER Pat
    Personal Access Token. If omitted, checks $env:AZURE_DEVOPS_PAT; if still
    empty, you'll be prompted but can just press Enter to skip it - in that
    case the script falls back to your current Windows login (NTLM/Kerberos),
    which is normally how on-prem/internal Azure DevOps Server is accessed.

.PARAMETER ApiVersion
    REST API version to call. Default 7.1 (cloud). On-prem Azure DevOps
    Server / TFS instances usually cap out lower (e.g. 6.0, 5.1, or 5.0) -
    if you get version errors, try passing an older one here.

.EXAMPLE
    .\Get-TestPlanResults.ps1 -Organization deltek -Project QEAutomation -Plan "Performance Testing" -OutputPath results.xlsx

.EXAMPLE
    .\Get-TestPlanResults.ps1 -Organization "https://tfs.deltek.com/tfs/Deltek" -Project QEAutomation -Plan 2535838 -ApiVersion 6.0

.EXAMPLE
    .\Get-TestPlanResults.ps1
    Prompts for everything interactively.

.NOTES
    Required PAT scopes (if using a PAT): Test Management (Read), Work Items (Read)
    Excel output requires the ImportExcel module:
        Install-Module ImportExcel -Scope CurrentUser
#>

[CmdletBinding()]
param(
    # Hardcoded for now per current setup. Override with -Organization / -Project / -Plan if needed.
    [string]$Organization = "https://tfs.deltek.com/tfs/Deltek",
    [string]$Project = "QEAutomation",
    [string]$Plan = "2535838",
    [string]$RootSuite,
    [string]$OutputPath = "test_plan_results.csv",
    [switch]$Detailed,
    [string]$Pat,
    [string]$ApiVersion = "7.1",
    [switch]$ResetPat
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = "Stop"
$ApiVersionWasSpecified = $PSBoundParameters.ContainsKey('ApiVersion')

function Get-ValueOrPrompt {
    param([string]$Value, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return Read-Host -Prompt $Label
    }
    return $Value
}

# PAT is cached encrypted (Windows DPAPI, tied to this user + this machine) outside
# the project folder so it never ends up in source control and never touches disk
# in plain text. Use -ResetPat to clear it (e.g. after the token rotates/expires).
$CredentialDir  = Join-Path $env:LOCALAPPDATA "AdoTestPlanExtractor"
$CredentialFile = Join-Path $CredentialDir "pat.dat"

function Get-SavedPat {
    if (-not (Test-Path $CredentialFile)) { return $null }
    try {
        $secure = Get-Content -Path $CredentialFile -ErrorAction Stop | ConvertTo-SecureString
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try {
            return [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        } finally {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    } catch {
        Write-Warning "Could not read saved PAT (it may have been saved by a different Windows account/machine). Ignoring it."
        return $null
    }
}

function Save-Pat {
    param([string]$PlainPat)
    if (-not (Test-Path $CredentialDir)) {
        New-Item -ItemType Directory -Path $CredentialDir -Force | Out-Null
    }
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
    try {
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }

    if (-not [string]::IsNullOrWhiteSpace($plain)) {
        Save-Pat -PlainPat $plain
    }
    return $plain
}

function Get-Prop {
    param($Obj, [string[]]$Path)
    $cur = $Obj
    foreach ($p in $Path) {
        if ($null -eq $cur) { return $null }
        $cur = $cur.$p
    }
    return $cur
}

function Get-BaseUrl {
    param([string]$OrgOrUrl)
    $trimmed = $OrgOrUrl.TrimEnd('/')
    if ($trimmed -match '^https?://') {
        return $trimmed
    }
    return "https://dev.azure.com/$trimmed"
}

$Organization = Get-ValueOrPrompt -Value $Organization -Label "ADO organization name, or full on-prem collection URL (e.g. https://tfs.deltek.com/tfs/Deltek)"
$Project      = Get-ValueOrPrompt -Value $Project -Label "ADO project name"
$Plan         = Get-ValueOrPrompt -Value $Plan -Label "Test Plan ID or name"
if ($ResetPat -and (Test-Path $CredentialFile)) {
    Remove-Item -Path $CredentialFile -Force
    Write-Host "Cleared saved PAT. You will be prompted for a new one."
}

$Pat          = Get-SecretOrPrompt -Value $Pat -Label "Personal Access Token (press Enter to use your Windows login instead)"

Write-Host "----------------------------------------"
Write-Host "Organization : $Organization"
Write-Host "Project      : $Project"
Write-Host "Test Plan ID : $Plan"
Write-Host "----------------------------------------"

$BaseUrl = Get-BaseUrl -OrgOrUrl $Organization
Write-Host "Base URL     : $BaseUrl"
$UseWindowsAuth = [string]::IsNullOrWhiteSpace($Pat)

if ($UseWindowsAuth) {
    Write-Host "No PAT provided - using your current Windows login (NTLM/Kerberos)."
    $AuthHeader = @{}
} else {
    $AuthHeader = @{
        Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Pat"))
    }
}

function Test-Endpoint {
    param([string]$UriWithoutVersion, [string]$Version)
    $sep = "?"
    if ($UriWithoutVersion -match '\?') { $sep = "&" }
    $testUri = $UriWithoutVersion + $sep + "api-version=$Version"
    try {
        if ($UseWindowsAuth) {
            $null = Invoke-RestMethod -Uri $testUri -UseDefaultCredentials -Method Get -ErrorAction Stop
        } else {
            $null = Invoke-RestMethod -Uri $testUri -Headers $AuthHeader -Method Get -ErrorAction Stop
        }
        return $true
    } catch {
        return $false
    }
}

function Resolve-WorkingApiVersion {
    param([string]$UriWithoutVersion, [string]$PreferredVersion)
    $candidates = @($PreferredVersion) + @("7.1","7.0","6.0","5.1","5.0","4.1","3.2","3.0","2.3","2.2","2.0","1.0") |
        Select-Object -Unique
    foreach ($v in $candidates) {
        Write-Host "  Probing api-version=$v against $UriWithoutVersion ..."
        if (Test-Endpoint -UriWithoutVersion $UriWithoutVersion -Version $v) {
            return $v
        }
    }
    return $null
}

function Invoke-AdoGet {
    param([string]$Uri, [hashtable]$QueryParams = @{})
    $qp = $QueryParams.Clone()
    $qp["api-version"] = $ApiVersion
    $pairs = foreach ($key in $qp.Keys) { "$key=$([Uri]::EscapeDataString([string]$qp[$key]))" }
    $fullUri = $Uri + "?" + ($pairs -join '&')
    Write-Host "  GET [$fullUri]"
    try {
        if ($UseWindowsAuth) {
            return Invoke-RestMethod -Uri $fullUri -UseDefaultCredentials -Method Get
        }
        return Invoke-RestMethod -Uri $fullUri -Headers $AuthHeader -Method Get
    } catch {
        Write-Host "  FAILED calling: $fullUri" -ForegroundColor Red
        throw
    }
}

function Get-AdoPaged {
    param([string]$Uri, [hashtable]$QueryParams = @{})
    $results = @()
    $skip = 0
    $top = 200
    while ($true) {
        $qp = $QueryParams.Clone()
        $qp['$top'] = $top
        $qp['$skip'] = $skip
        $data = Invoke-AdoGet -Uri $Uri -QueryParams $qp
        $batch = @($data.value)
        if ($batch.Count -eq 0) { break }
        $results += $batch
        if ($batch.Count -lt $top) { break }
        $skip += $top
    }
    return $results
}

function Resolve-PlanId {
    param([string]$PlanIdentifier)
    if ($PlanIdentifier -match '^\d+$') {
        return @{ Id = [int]$PlanIdentifier; Name = $PlanIdentifier }
    }
    $uri = "$BaseUrl/$Project/_apis/testplan/plans"
    $data = Invoke-AdoGet -Uri $uri
    $plans = @($data.value)
    $matched = $plans | Where-Object { $_.name -like "*$PlanIdentifier*" }
    if (-not $matched -or $matched.Count -eq 0) {
        throw "No test plan found matching '$PlanIdentifier'. Available: $($plans.name -join ', ')"
    }
    if ($matched.Count -gt 1) {
        $exact = $matched | Where-Object { $_.name -eq $PlanIdentifier }
        if (@($exact).Count -eq 1) { return @{ Id = $exact[0].id; Name = $exact[0].name } }
        throw "Multiple plans match '$PlanIdentifier': $($matched.name -join ', '). Use the exact name or ID."
    }
    return @{ Id = $matched[0].id; Name = $matched[0].name }
}

function Get-AllSuites {
    param([int]$PlanId)
    $uri = "$BaseUrl/$Project/_apis/test/Plans/$PlanId/suites"
    $suites = Get-AdoPaged -Uri $uri
    $map = @{}
    foreach ($s in $suites) { $map[[int]$s.id] = $s }
    return $map
}

function Build-SuitePaths {
    param([hashtable]$Suites)
    $paths = @{}

    function Walk {
        param([int]$SuiteId)
        if ($paths.ContainsKey($SuiteId)) { return $paths[$SuiteId] }
        $suite = $Suites[$SuiteId]
        if ($null -eq $suite) { return "?" }
        $parentId = Get-Prop $suite @('parentSuite', 'id')
        if ($parentId -and $Suites.ContainsKey([int]$parentId)) {
            $parentPath = Walk -SuiteId ([int]$parentId)
            $path = "$parentPath > $($suite.name)"
        } else {
            $path = $suite.name
        }
        $paths[$SuiteId] = $path
        return $path
    }

    foreach ($sid in $Suites.Keys) { [void](Walk -SuiteId $sid) }
    return $paths
}

function Filter-ToSubtree {
    param([hashtable]$Suites, [string]$RootName)
    $rootIds = @($Suites.Keys | Where-Object { $Suites[$_].name -ieq $RootName })
    if ($rootIds.Count -eq 0) {
        $topLevel = $Suites.Values | Where-Object { -not (Get-Prop $_ @('parentSuite', 'id')) } | Select-Object -ExpandProperty name
        throw "-RootSuite '$RootName' not found. Available top-level suites: $($topLevel -join ', ')"
    }

    $keep = [System.Collections.Generic.HashSet[int]]::new()
    foreach ($id in $rootIds) { [void]$keep.Add($id) }

    $changed = $true
    while ($changed) {
        $changed = $false
        foreach ($sid in $Suites.Keys) {
            $parentId = Get-Prop $Suites[$sid] @('parentSuite', 'id')
            if ($parentId -and $keep.Contains([int]$parentId) -and -not $keep.Contains($sid)) {
                [void]$keep.Add($sid)
                $changed = $true
            }
        }
    }

    $filtered = @{}
    foreach ($sid in $Suites.Keys) { if ($keep.Contains($sid)) { $filtered[$sid] = $Suites[$sid] } }
    return $filtered
}

function Get-PointsForSuite {
    param([int]$PlanId, [int]$SuiteId)
    $uri = "$BaseUrl/$Project/_apis/test/Plans/$PlanId/Suites/$SuiteId/points"
    return Get-AdoPaged -Uri $uri
}

function Get-TestCaseTitles {
    param([int[]]$TestCaseIds)
    $titles = @{}
    if ($TestCaseIds.Count -eq 0) { return $titles }
    $uri = "$BaseUrl/$Project/_apis/wit/workitemsbatch?api-version=$ApiVersion"
    for ($i = 0; $i -lt $TestCaseIds.Count; $i += 200) {
        $endIndex = [Math]::Min($i + 199, $TestCaseIds.Count - 1)
        $chunk = @($TestCaseIds[$i..$endIndex])
        $body = @{ ids = $chunk; fields = @("System.Title") } | ConvertTo-Json
        if ($UseWindowsAuth) {
            $resp = Invoke-RestMethod -Uri $uri -UseDefaultCredentials -Method Post -Body $body -ContentType "application/json"
        } else {
            $resp = Invoke-RestMethod -Uri $uri -Headers $AuthHeader -Method Post -Body $body -ContentType "application/json"
        }
        foreach ($item in @($resp.value)) {
            $titles[[int]$item.id] = $item.fields.'System.Title'
        }
    }
    return $titles
}

function Get-ResultDetail {
    param([string]$RunId, [string]$ResultId)
    try {
        $uri = "$BaseUrl/$Project/_apis/test/Runs/$RunId/results/$ResultId"
        $data = Invoke-AdoGet -Uri $uri
        return @{
            CompletedDate = $data.completedDate
            DurationInMs  = $data.durationInMs
            ErrorMessage  = $data.errorMessage
        }
    } catch {
        return @{ CompletedDate = ""; DurationInMs = ""; ErrorMessage = "" }
    }
}

function Write-Results {
    param([array]$Rows, [string]$Path)
    if ($Path -like "*.xlsx") {
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module ImportExcel
            $Rows | Export-Excel -Path $Path -AutoSize
            return
        }
        Write-Warning "ImportExcel module not found. Install with: Install-Module ImportExcel -Scope CurrentUser"
        Write-Warning "Falling back to CSV."
        $Path = [IO.Path]::ChangeExtension($Path, ".csv")
    }
    $Rows | Export-Csv -Path $Path -NoTypeInformation
    Write-Host "Wrote $($Rows.Count) rows to $Path"
}

# ---- Main ----

try {
    Write-Host "Resolving plan '$Plan'..."
    $planInfo = Resolve-PlanId -PlanIdentifier $Plan
    Write-Host "Using plan: $($planInfo.Name) (ID $($planInfo.Id))"

    if (-not $ApiVersionWasSpecified) {
        $suitesProbeUri = "$BaseUrl/$Project/_apis/test/Plans/$($planInfo.Id)/suites"
        Write-Host "No -ApiVersion specified. Auto-detecting a working REST API version against the Test Plan suites endpoint..."
        $working = Resolve-WorkingApiVersion -UriWithoutVersion $suitesProbeUri -PreferredVersion $ApiVersion
        if ($working) {
            $ApiVersion = $working
            Write-Host "Using api-version=$ApiVersion"
        } else {
            Write-Warning "No api-version could reach $suitesProbeUri. This usually means: (a) plan $($planInfo.Id) does not belong to project '$Project', or (b) the Test Management REST API isn't available on this server. Double-check the project name and plan ID in the ADO web UI."
        }
    }

    Write-Host "Fetching suite tree..."
    $suites = Get-AllSuites -PlanId $planInfo.Id
    if ($RootSuite) { $suites = Filter-ToSubtree -Suites $suites -RootName $RootSuite }
    $paths = Build-SuitePaths -Suites $suites
    Write-Host "Found $($suites.Count) suite(s)."

    $rows = @()
    $allTestCaseIds = [System.Collections.Generic.HashSet[int]]::new()

    foreach ($suiteId in $suites.Keys) {
        Write-Host "  Pulling points for suite: $($paths[$suiteId])"
        $points = Get-PointsForSuite -PlanId $planInfo.Id -SuiteId $suiteId
        foreach ($p in $points) {
            $tcIdRaw = Get-Prop $p @('testCase', 'id')
            if ($null -eq $tcIdRaw) { continue }
            $tcId = [int]$tcIdRaw
            [void]$allTestCaseIds.Add($tcId)

            $outcome = "Not Run"
            if ($p.outcome) { $outcome = $p.outcome }

            $row = [ordered]@{
                "Suite Path"      = $paths[$suiteId]
                "Suite ID"        = $suiteId
                "Test Case ID"    = $tcId
                "Test Case Title" = ""
                "Outcome"         = $outcome
                "Configuration"   = Get-Prop $p @('configuration', 'name')
                "Assigned To"     = Get-Prop $p @('assignedTo', 'displayName')
                "Last Run ID"     = Get-Prop $p @('lastTestRun', 'id')
                "Last Result ID"  = Get-Prop $p @('lastResult', 'id')
            }

            # Always add these columns when -Detailed is set (even if blank)
            # so every row has the same schema for a clean CSV/Excel export.
            if ($Detailed) {
                $completedDate = ""
                $durationMs = ""
                $errorMessage = ""
                if ($row["Last Run ID"] -and $row["Last Result ID"]) {
                    $detail = Get-ResultDetail -RunId $row["Last Run ID"] -ResultId $row["Last Result ID"]
                    $completedDate = $detail.CompletedDate
                    $durationMs = $detail.DurationInMs
                    $errorMessage = $detail.ErrorMessage
                }
                $row["Completed Date"] = $completedDate
                $row["Duration (ms)"]  = $durationMs
                $row["Error Message"]  = $errorMessage
            }

            $rows += [PSCustomObject]$row
        }
    }

    if ($rows.Count -eq 0) {
        Write-Warning "No rows found."
        return
    }

    Write-Host "Resolving test case titles..."
    $titles = Get-TestCaseTitles -TestCaseIds @($allTestCaseIds)
    foreach ($row in $rows) {
        $row."Test Case Title" = $titles[[int]$row."Test Case ID"]
    }

    Write-Results -Rows $rows -Path $OutputPath
    Write-Host "Done."
}
catch {
    Write-Error "Failed: $($_.Exception.Message)"
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
        Write-Error "Details: $($_.ErrorDetails.Message)"
    }
    exit 1
}