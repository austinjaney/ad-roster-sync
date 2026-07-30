<#
.SYNOPSIS
    Audits and normalises the accountExpires attribute across example.local.

.DESCRIPTION
    Addresses the first root cause of the a past incident offboarding incident, in which an
    the identity-automation platform workflow meant to target one expired test account disabled 14 live staff
    accounts.

    accountExpires is an Integer8 holding a Windows FILETIME. Active Directory accepts TWO
    different values for "this account never expires":

        0                     often written by older tools and by some ADUC code paths
        9223372036854775807   Int64.MaxValue, 0x7FFFFFFFFFFFFFFF

    Both are equivalent to authentication, and Get-ADUser reports AccountExpirationDate as
    empty for either. They are NOT equivalent to a numeric filter. A workflow asking for
    "accounts whose expiry has passed" with a comparison like

        (accountExpires<=132000000000000000)

    matches every account storing 0, because 0 is less than any date. That is precisely how
    a workflow scoped to expired accounts swept up accounts that had no expiry at all.

    This script reports the split, then optionally rewrites one representation to the other
    so a single convention holds domain-wide.

    NORMALISATION REDUCES THE EXPOSURE. IT DOES NOT REMOVE IT. Any query that filters on
    accountExpires must still exclude both sentinels explicitly - see the LDAP filter this
    script prints at the end. Fix the queries as well as the data.

.PARAMETER NormalizeTo
    Which sentinel becomes the convention: 'MaxValue' (default) or 'Zero'. MaxValue is what
    current ADUC writes when you tick "Never", and being the largest possible value it
    cannot satisfy a "less than some date" test, which is the failure mode that caused the
    incident. Zero is only preferable if existing tooling depends on it.

.PARAMETER SearchBase
    Restrict the audit to one subtree. Defaults to the whole domain, which is what you want
    for an audit of this kind.

.PARAMETER Commit
    Write the normalisation. Omit for an audit only.

.PARAMETER MaxChanges
    Refuse to write if more than this many objects would change. Default 2000.

.PARAMETER RollbackFrom
    Path to an AccountExpires-<timestamp>.csv from an earlier -Commit run. Restores the
    RawValue recorded for every row it normalised, then exits. Add -Commit to write. A row
    whose value no longer matches what that run wrote is skipped, not clobbered.

.EXAMPLE
    .\Repair-ADAccountExpires.ps1
    Audit only. Reports the split and writes a CSV.

.EXAMPLE
    .\Repair-ADAccountExpires.ps1 -Commit

.EXAMPLE
    .\Repair-ADAccountExpires.ps1 -RollbackFrom .\AccountExpires-20260730-091500.csv -Commit

.NOTES
    Changing accountExpires between the two sentinels does not alter whether an account can
    log on: both mean "never expires" before and after. Accounts with a real expiry date are
    never touched.
#>

[CmdletBinding()]
param(
    [ValidateSet('MaxValue','Zero')]
    [string] $NormalizeTo   = 'MaxValue',
    [string] $SearchBase,
    [string] $LogDirectory  = $PSScriptRoot,
    [switch] $Commit,
    [int]    $MaxChanges    = 2000,
    [string] $RollbackFrom
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module ActiveDirectory -ErrorAction Stop

$NEVER_MAX  = [int64]9223372036854775807
$NEVER_ZERO = [int64]0
$target     = if ($NormalizeTo -eq 'MaxValue') { $NEVER_MAX } else { $NEVER_ZERO }
$source     = if ($NormalizeTo -eq 'MaxValue') { $NEVER_ZERO } else { $NEVER_MAX }

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$log   = Join-Path $LogDirectory "AccountExpires-$stamp.csv"
$mode  = if ($Commit) { 'COMMIT' } else { 'AUDIT' }

Write-Host ''
Write-Host '=== accountExpires audit ===' -ForegroundColor Cyan
Write-Host ("Mode          : {0}" -f $mode) -ForegroundColor $(if ($Commit) { 'Red' } else { 'Green' })
Write-Host ("Convention    : {0} ({1})" -f $NormalizeTo, $target)
Write-Host ''

# ---------------------------------------------------------------- rollback ----
if ($RollbackFrom) {
    if (-not (Test-Path -LiteralPath $RollbackFrom)) { throw "Report not found: $RollbackFrom" }
    $prior = @(Import-Csv -LiteralPath $RollbackFrom | Where-Object { $_.WillNormalize -eq 'True' })
    Write-Host ("{0} rows were normalised in that run" -f $prior.Count)
    $done = 0; $skip = 0; $fail = 0
    foreach ($r in $prior) {
        try {
            $u = Get-ADUser -Identity $r.SamAccountName -Properties accountExpires -ErrorAction Stop
            $live = [int64]$u.accountExpires
            # Only revert if the value is still the one this script wrote.
            if ($live -ne [int64]$r.NewValue) {
                Write-Host ("  [skip] {0}: now {1}, not the {2} this run wrote" -f $r.SamAccountName, $live, $r.NewValue) -ForegroundColor Yellow
                $skip++; continue
            }
            if ($Commit) {
                Set-ADUser -Identity $u.DistinguishedName -Replace @{ accountExpires = [int64]$r.RawValue } -ErrorAction Stop
            }
            $done++
        } catch {
            Write-Host ("  [ERROR] {0}: {1}" -f $r.SamAccountName, $_.Exception.Message) -ForegroundColor Red
            $fail++
        }
    }
    Write-Host ''
    Write-Host ("Rollback {0}: {1} reverted, {2} skipped as changed since, {3} failed" -f `
        $(if ($Commit) { 'applied' } else { 'planned - add -Commit to write' }), $done, $skip, $fail) -ForegroundColor Green
    return
}

$params = @{
    LDAPFilter = '(&(objectCategory=person)(objectClass=user))'
    Properties = @('accountExpires','SamAccountName','DisplayName','Enabled',
                   'DistinguishedName','LastLogonDate')
}
if ($SearchBase) { $params['SearchBase'] = $SearchBase }
$users = @(Get-ADUser @params)
Write-Host ("{0} user objects read" -f $users.Count)

$rows   = [System.Collections.Generic.List[object]]::new()
$counts = [ordered]@{ 'never (0)' = 0; 'never (MaxValue)' = 0; 'real expiry, future' = 0
                      'real expiry, past' = 0; 'unreadable' = 0 }
$now = Get-Date

foreach ($u in $users) {
    $raw = if ($u.PSObject.Properties.Name -contains 'accountExpires') { $u.accountExpires } else { $null }
    $val = $null
    if ($null -ne $raw) { try { $val = [int64]$raw } catch { $val = $null } }

    if ($null -eq $val)          { $bucket = 'unreadable';       $expiry = '' }
    elseif ($val -eq $NEVER_ZERO){ $bucket = 'never (0)';        $expiry = 'never' }
    elseif ($val -eq $NEVER_MAX) { $bucket = 'never (MaxValue)'; $expiry = 'never' }
    else {
        try   { $d = [datetime]::FromFileTimeUtc($val) }
        catch { $d = $null }
        if ($null -eq $d) { $bucket = 'unreadable'; $expiry = "uninterpretable: $val" }
        else {
            $bucket = if ($d -gt $now) { 'real expiry, future' } else { 'real expiry, past' }
            $expiry = $d.ToString('yyyy-MM-dd')
        }
    }
    $counts[$bucket]++

    # Only the two sentinels are ever rewritten. A real date, past or future, is data.
    $needsFix = ($val -eq $source)
    $rows.Add([pscustomobject]@{
        SamAccountName = $u.SamAccountName
        DisplayName    = $u.DisplayName
        Enabled        = $u.Enabled
        Bucket         = $bucket
        RawValue       = if ($null -ne $val) { $val } else { '' }
        ExpiresOn      = $expiry
        WillNormalize  = $needsFix
        NewValue       = if ($needsFix) { $target } else { '' }
        DistinguishedName = $u.DistinguishedName
    })
}

Write-Host ''
Write-Host '--- accountExpires distribution ---' -ForegroundColor Cyan
foreach ($k in $counts.Keys) { Write-Host ("  {0,-22} {1,6}" -f $k, $counts[$k]) }

$toFix = @($rows | Where-Object WillNormalize)
Write-Host ''
Write-Host ("{0} objects use the non-standard sentinel ({1}) and would be set to {2}" -f `
    $toFix.Count, $source, $target)

# The value that matters more than the cleanup: a filter that cannot repeat the incident.
$safeFilter = "(&(accountExpires>=1)(!(accountExpires=$NEVER_MAX)))"
Write-Host ''
Write-Host '--- use this filter for "has a real expiry date" ---' -ForegroundColor Cyan
Write-Host "  $safeFilter"
Write-Host '  accountExpires>=1 excludes the 0 sentinel; the negation excludes MaxValue.'
Write-Host '  Combine with an explicit -SearchBase. Scope was the second root cause on a past incident.'

if ($Commit) {
    if ($toFix.Count -gt $MaxChanges) {
        $rows | Export-Csv -LiteralPath $log -NoTypeInformation -Encoding UTF8
        throw "Would change $($toFix.Count) objects, over -MaxChanges $MaxChanges. Nothing written. See $log"
    }
    Write-Host ''
    Write-Host '--- Applying ---' -ForegroundColor Red
    $ok = 0; $fail = 0
    foreach ($r in $toFix) {
        try {
            Set-ADUser -Identity $r.DistinguishedName -Replace @{ accountExpires = $target } -ErrorAction Stop
            $ok++
        } catch {
            Write-Host ("  [ERROR] {0}: {1}" -f $r.SamAccountName, $_.Exception.Message) -ForegroundColor Red
            $fail++
        }
    }
    Write-Host ("  {0} normalised, {1} failed" -f $ok, $fail)
}

$rows | Export-Csv -LiteralPath $log -NoTypeInformation -Encoding UTF8
Write-Host ''
Write-Host ("Report: {0}" -f $log) -ForegroundColor Green
if (-not $Commit) {
    Write-Host 'Audit only - no changes written. Re-run with -Commit to normalise.' -ForegroundColor Yellow
}
