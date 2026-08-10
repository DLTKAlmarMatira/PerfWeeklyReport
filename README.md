# ADO Extract & Report Pipeline

Pulls performance-testing data out of Deltek's **on-prem TFS** server and turns it into
the weekly CSVs and the HTML meeting dashboard.

One command does everything:

```
Run-AdoExtracts.bat
```

Double-click it, or run it from a terminal in this folder. It pauses at the end so you
can read the result.

---

## Quick start

| I want to... | Run this |
|---|---|
| Refresh everything from ADO | `Run-AdoExtracts.bat` |
| Rebuild reports from data I already pulled | `powershell -File 6-Build-WeeklyReports.ps1` |
| Rebuild just the HTML dashboard | `powershell -File 7-Build-MeetingReport.ps1 -Show` |

The last two need **no network and no credentials** — they only re-read the CSVs in
`csv\`. Re-run them as often as you like.

---

## Prerequisites

- **Windows PowerShell 5.1** (the built-in one; nothing to install)
- **Network access to `tfs.deltek.com`** — on the corporate network or VPN
- Optional: the `ImportExcel` module, only if you want `.xlsx` output instead of `.csv`

No Python is needed for this pipeline. (The two `build_*pptx.py` deck scripts in this
folder are a separate, unrelated toolchain.)

---

## What the batch actually runs

Six steps, in this order. Any failure stops the run — it will not continue with
missing or half-written data.

| # | Script | Network? | Produces |
|---|---|---|---|
| 1 | `1-Get-TestPlanResults.ps1` | yes | `csv\test_plan_results.csv` |
| 2 | `2-Get-AdoQueryResults.ps1` | yes | `csv\pbi_task_links.csv` |
| 3 | `3-Get-TaskTestsLinkResults.ps1` | yes | `csv\task_tests_link_results.csv` |
| 4 | `8-Get-PbiBugLinks.ps1` | yes | `csv\pbi_bug_links.csv` |
| 5 | `6-Build-WeeklyReports.ps1` | **no** | 4 derived CSVs (below) |
| 6 | `7-Build-MeetingReport.ps1` | **no** | `weekly_meeting_report.html` |

Steps 1–4 are the **extraction** — they talk to ADO and may prompt for credentials.
Steps 5–6 are **pure local transformation** of what steps 1–4 wrote.

> The script numbers are not the step order: the batch runs 1, 2, 3, **8**, 6, 7.
> Script 8 was added after 6 and 7 were already numbered, and it has to run before
> them because 7 reads its output.

### What each step does

**1 — Test plan results.** Walks every Test Suite under Test Plan `2535838` and records
the current outcome of every Test Case in each suite. One row per *test point*.

**2 — PBI → Task links.** Runs the saved "PBI - Active" query and exports its results,
plus extra fields the query itself doesn't return. This is the **task roster** — the
authoritative list of who is working on what.

**3 — Task → Test links.** Runs the saved "Tests" link query, capturing which tasks link
to which Test Cases and Test Plans, and by what link type. This is where an incorrectly
used `Child` link (instead of `Tests`) becomes visible.

**4 — PBI → Bug links.** Bugs hang off a PBI with a plain `Related` link, which the saved
query in step 2 does not return — it asks only for Parent/Child. This step walks the
relations through the REST API instead, so nobody has to edit the ADO query.

**5 — Weekly reports.** Joins the three raw extracts into the derived reporting CSVs.
No network, no credentials.

**6 — HTML dashboard.** Renders one self-contained `weekly_meeting_report.html` — data,
styling and scripting all embedded, no server and no internet needed. Email it, copy it
to a share, open it from disk; it works anywhere.

---

## Output

Everything lands in `csv\` next to the batch file, except the HTML report which lands
beside the batch file itself.

**Raw extracts** — point-in-time snapshots of ADO. These are committed to git, because
once ADO moves on they cannot be recovered:

```
csv\test_plan_results.csv
csv\pbi_task_links.csv
csv\task_tests_link_results.csv
csv\pbi_bug_links.csv
```

**Derived reports** — fully recomputable from the above, so they are *not* committed.
If one looks stale, regenerate rather than trusting it:

```
csv\connected_pbi_task_test_results.csv   the full joined dataset
csv\pbi_summary_weekly.csv                one row per PBI
csv\tester_workload_weekly.csv            one row per tester
csv\exceptions_weekly.csv                 link mistakes + failed/blocked outcomes
weekly_meeting_report.html                the dashboard
```

---

## Credentials

All four network steps share one sign-in. They look for a token in this order:

1. A `-Pat` argument, if you passed one
2. A cached token at `%LOCALAPPDATA%\AdoTestPlanExtractor\pat.dat`
   (encrypted with DPAPI — readable only by your Windows account on this machine)
3. The `AZURE_DEVOPS_PAT` environment variable
4. An interactive prompt

**For the on-prem server, just press Enter at the prompt.** An empty token makes the
scripts fall back to Windows authentication, which is the normal path here — you almost
certainly do not need a PAT at all.

If you do use a PAT and it later rotates, clear the cache with `-ResetPat`:

```
powershell -File 1-Get-TestPlanResults.ps1 -ResetPat
```

The cache deliberately lives outside this repo. **Never put a PAT in a file in this
folder.**

---

## Running the pieces individually

Each script works standalone and defaults to the on-prem server and the `QEAutomation`
project, so you usually need no arguments beyond an output path.

```powershell
# Re-pull one extract
powershell -NoProfile -ExecutionPolicy Bypass -File 1-Get-TestPlanResults.ps1 -OutputPath csv\test_plan_results.csv

# Rebuild the derived CSVs, then the dashboard
powershell -File 6-Build-WeeklyReports.ps1
powershell -File 7-Build-MeetingReport.ps1 -Show
```

Script 6 also takes `-PassThru`, which sends the joined rows down the pipeline so you can
answer a one-off question without opening a CSV:

```powershell
.\6-Build-WeeklyReports.ps1 -PassThru | Where-Object Outcome -eq 'Failed' | Group-Object Product
```

### Parameters

Common to the network scripts (1, 2, 3, 8): `-Organization`, `-Project`, `-Pat`,
`-ApiVersion`, and `-ResetPat` (all except script 8). None are mandatory.

| Script | Notable parameters | Default output |
|---|---|---|
| `1-Get-TestPlanResults` | `-Plan` (default `2535838`), `-RootSuite`, `-Detailed` | `test_plan_results.csv` |
| `2-Get-AdoQueryResults` | `-QueryUrl` (prompts if omitted), `-ExtraFields` | `query_results.csv` |
| `3-Get-TaskTestsLinkResults` | `-QueryUrl` (hardcoded default), `-ExtraFields` | `task_tests_link_results.csv` |
| `8-Get-PbiBugLinks` | `-CsvDir` | `csv\pbi_bug_links.csv` |
| `6-Build-WeeklyReports` | `-CsvDir`, `-PassThru` | 4 CSVs into `csv\` |
| `7-Build-MeetingReport` | `-CsvDir`, `-OutputPath`, `-Show` | `weekly_meeting_report.html` |

Scripts 1, 2 and 3 pick their output format from the file extension — pass a `.xlsx` path
and they use Excel, if the `ImportExcel` module is installed. Without it they warn and
write `.csv` instead. Scripts 4, 5 and 8 always write CSV.

---

## Configuring it

Two things in `Run-AdoExtracts.bat` are worth knowing about:

**The saved query URL** (`PBI_QUERY_URL`, near the top). If the "PBI - Active" query is
ever moved or recreated, paste its new URL here.

**The extra field list** (`EXTRA_FIELDS`). Fields the saved queries don't return but the
reports need. Two rules if you edit it:

- It must stay **one comma-separated string in one set of quotes.** Space-separating the
  names does not work across the `cmd.exe → powershell.exe` boundary — only the first
  name binds, and the rest spill into `-Pat` and `-ApiVersion`, producing confusing
  "positional parameter cannot be found" errors. The batch file's own comments explain
  this at length.
- Only add fields that actually exist on this org's work item types. Confirm with
  `4-Get-AdoWorkItemTypeFields.ps1` first. In particular
  `Microsoft.VSTS.Common.ActivatedDate` does **not** exist here — use
  `Microsoft.VSTS.Common.StateChangeDate` for "worked on this week".

---

## Two scripts not in the batch

Diagnostic tools, run by hand when you need them.

**`4-Get-AdoWorkItemTypeFields.ps1`** — lists every field that actually exists on Task and
Product Backlog Item in this project, flagging anything that looks like a "Blocked" or
"Due Date" field. Run this before adding a field to `EXTRA_FIELDS`; guessing at field
names is how you get an extract that succeeds and returns nothing.

```powershell
powershell -File 4-Get-AdoWorkItemTypeFields.ps1 -OutputPath csv\workitemtype_fields.csv
```

**`5-Get-AdoIterationCapacity.ps1`** — per-person capacity for an iteration. Discover the
names first, then pull:

```powershell
powershell -File 5-Get-AdoIterationCapacity.ps1 -ListTeams
powershell -File 5-Get-AdoIterationCapacity.ps1 -ListIterations
powershell -File 5-Get-AdoIterationCapacity.ps1 -IterationPath "QEAutomation\Costpoint"
```

---

## Troubleshooting

**"Missing raw extract(s)" from step 5.** The three files from steps 1–3 aren't in
`csv\`. Run the full batch, or check whether they've been moved into `csv\old\`.

**"Not found: connected_pbi_task_test_results.csv" from step 6.** Step 5 hasn't run
since the extracts changed. Run `powershell -File 6-Build-WeeklyReports.ps1`.

**An extract succeeds but the file is empty or nearly so.** Most likely an API version
mismatch, not an absence of data. The on-prem server does not support
`api-version=7.1`, so the scripts probe downward (7.1 → 7.0 → 6.0 → 5.1 → …) and use the
first that answers. Don't pin `-ApiVersion` unless you have a reason to.

**"Positional parameter cannot be found."** Almost always a space-separated
`-ExtraFields` list. See *Configuring it* above.

**Credential prompt loops or 401s.** Clear the cached token with `-ResetPat`, then press
Enter at the prompt to use Windows auth.

**The report opens but numbers look stale.** The HTML is generated, not live. Its "as of"
date is stamped at build time — check it against today before quoting any figure in a
meeting.

---

## Reading the report

Two things in the dashboard are commonly misread:

**Pass rate is not run rate.** The headline figure is **pass rate** — passed ÷ actually
run. Run rate (run ÷ planned) is a *coverage* measure; it is dragged down by tests nobody
has executed yet and says nothing about quality. Every pass rate is printed directly
above its denominator (`246 of 640 run`) so it can't be quoted without its coverage.

**Passed can legitimately exceed Cases.** `Cases` counts distinct test case IDs; `Points`
counts test points. One case sitting in several suites is several points. The `Points`
column exists so that reconciles visibly instead of looking like a bug.

For the process rules behind all of this — what links to what, and why a `Tests` link is
not a `Child` link — see **`ADO_Performance_Testing_Structure.md`**.
