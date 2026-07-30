# Runbook

Generic operating procedure. The original was specific to one directory; adapt freely.

## Prerequisites

Domain-joined workstation, RSAT ActiveDirectory module, delegated write rights on the target
OUs.

```powershell
Get-Module ActiveDirectory -ListAvailable
Get-ADOrganizationalUnit -Identity 'OU=Staff,OU=Accounts,OU=ORG,DC=example,DC=local'
Unblock-File .\*.ps1
```

## 1. Export account state

```powershell
.\Export-ADDisabledAccounts.ps1
```

Read-only. An Entra portal export has no `accountEnabled` column, and `onPremisesSyncEnabled`
describes directory sync rather than account state — so this is the only way to tell a disabled
account from an active one.

## 2. Build the work files

```bash
python3 Build-ADReconciliation.py
```

Produces `AD-Updates.csv` (approved matches), a review CSV, an exceptions list and an enriched
workbook.

## 3. Approve the uncertain matches

Open the review CSV. Rows already marked `Approved = Y` matched with high confidence. Anything
blank needs a human — usually a spelling variant, an initial collision, or a name change.

Record the decisions in the tables at the top of `Build-ADReconciliation.py` rather than only
in the CSV, so they survive the next rebuild.

## 4. Dry run

```powershell
.\Sync-ADFromRoster.ps1
```

Writes nothing. The change log it produces is the complete plan.

## 5. Review before committing

```powershell
Import-Csv .\RosterSync-*-changes.csv | Group-Object Action | Sort-Object Count -Descending
Import-Csv .\RosterSync-*-changes.csv | Where-Object Status -eq 'Skipped' | Group-Object Action
Import-Csv .\RosterSync-*-issues.csv  | Where-Object Severity -eq 'ERROR'
```

Stop and investigate if you see:

- **`disabled pass ABANDONED`** — the mass-disable circuit breaker fired. Find out what
  disabled those accounts before re-running.
- **An OU move count near your whole staff** — plausible on a first run, suspicious afterwards.
- **`DisabledOnRoster` rows** — an account is disabled yet on the roster. Resolve which is right.

## 6. Pilot, then commit

Prove the write path on a handful of people first. Filter the updates CSV to one site and run
with the scanning passes disabled:

```powershell
.\Sync-ADFromRoster.ps1 -UpdatesCsv .\AD-Updates-PILOT.csv -SkipVolunteerPass -SkipDisabledPass -SkipGalPass -Commit
```

Verify, roll it back to prove the reverse path, then commit for real:

```powershell
.\Sync-ADFromRoster.ps1 -RollbackFrom .\RosterSync-<stamp>-changes.csv -Commit
.\Sync-ADFromRoster.ps1 -Commit
```

## 7. Sync outward

```powershell
Start-ADSyncSyncCycle -PolicyType Delta
```

Title, department and company all surface in the address book, so a large commit becomes
visible organisation-wide as soon as this completes. Worth timing deliberately.

## If something goes wrong

```powershell
.\Sync-ADFromRoster.ps1 -RollbackFrom .\RosterSync-<stamp>-changes.csv -Commit
```

Only rows whose `Status` is `Applied` are reversed. Moves reverse before attributes so restored
manager DNs resolve. An attribute that changed since the run is skipped, not clobbered, unless
`-Force`.

**Archive the change logs.** They are the only record of prior values and the only input
rollback can consume. They also contain before-and-after state for every person — keep them
somewhere durable and out of source control.

### What rollback does not undo

- **Time already elapsed in the address book.** AD is restored; what people already saw is not.
- **Downstream reactions.** If a dynamic group or licence rule keys off `department` or OU,
  those systems should recompute, but any action they already took is theirs to reverse.
- Nothing here deletes, and nothing enables or disables an account, so there is nothing else to
  restore.
