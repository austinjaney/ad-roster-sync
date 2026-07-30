<#
.SYNOPSIS
    Reconciles example.local Active Directory against the Staff Roster source of truth.

.DESCRIPTION
    Three passes, in order:

      1. ATTRIBUTES - for every approved roster match, set title, department, description,
         company (the campus), employeeType (FT/PT) and manager.
      2. OU PLACEMENT - move each matched user into OU=<campus>,OU=Staff,OU=Accounts,
         OU=ORG,DC=example,DC=local.
      3. DISABLED CLEANUP - a disabled account inside Staff\* is moved to
         OU=Disabled Accounts. An account that is disabled but still on the roster is a
         contradiction, so it is reported instead of moved.
      5. ADDRESS BOOK VISIBILITY - disabled accounts get msExchHideFromAddressLists set,
         enabled staff and volunteers get it cleared. Hiding is only ever used here to
         retire a departed account, so an enabled person should never carry the flag.
         Shared mailboxes and room accounts are left alone: they are neither staff nor
         volunteers, and theirs is a deliberate setting.
      4. VOLUNTEER FLAG - any enabled user already inside Staff\* whose UPN ends in
         @example.org and who is NOT on the roster gets title and description
         set to "Volunteer". Staff who move into volunteer roles keep their account and
         their OU, so the label is what distinguishes them from current staff. Dormant
         accounts are labelled too, and additionally flagged in the change log.

    The domain is Entra-hybrid, so Entra Connect propagates every change written here
    to Entra ID / Microsoft 365 on its next delta sync.

    Nothing is written unless -Commit is supplied. Without it the script performs a
    full read-only pass and produces the same logs, so the CSV log IS the change plan.

.PARAMETER UpdatesCsv
    AD-Updates.csv - one row per approved roster-to-AD match.

.PARAMETER MappingCsv
    AD-Name-Mapping-REVIEW.csv - the human-reviewed mapping. Only rows with
    Approved = Y are trusted. Rows approved here that are absent from UpdatesCsv are
    picked up automatically, so approving a REVIEW row does not require regenerating
    the updates file.

.PARAMETER Commit
    Perform writes. Omit for a dry run.

.PARAMETER SkipMove
    Update attributes only; never move objects between OUs.

.PARAMETER SkipVolunteerPass
    Skip the volunteer pass entirely.

.PARAMETER SkipDisabledPass
    Skip the disabled-account cleanup.

.PARAMETER SkipDisabledAccounts
    Leave every disabled account untouched, in all five passes. Use this when the goal is
    accurate staff and volunteer data and disabled accounts are somebody else's problem for
    now. Specifically it:
      - skips a roster member whose account is disabled, rather than updating their
        attributes and moving their OU
      - implies -SkipDisabledPass, so no disabled account is relocated
      - stops pass 5 hiding disabled accounts from the address book, while still unhiding
        enabled staff and volunteers
    Note the consequence: departed accounts that are currently visible in the address book
    stay visible. The dry run reports how many that is.

.PARAMETER SkipGalPass
    Skip the address-book visibility pass.

.PARAMETER AlsoScanOus
    Sibling OUs of Staff, by name, to include in the volunteer and disabled passes.
    Defaults to 'New Accounts', a staging OU whose occupants are otherwise never seen by
    either pass. Pass an empty array to confine both passes to Staff\*.

.PARAMETER MaxDisabledMoves
    Abandon the disabled pass, without moving anything, if more than this many accounts
    qualify. Default 25. A spike in disabled accounts is far more likely to be a runaway
    automation than a genuine wave of departures: on a past incident an identity-automation workflow disabled
    a number of live staff accounts in one pass. Relocating those would have added OU restoration to that
    recovery.

.PARAMETER RecentLogonHoldDays
    Hold back a disabled account that logged on within this many days. Default 7. Someone
    who genuinely left had usually stopped signing in before the account was disabled; a
    disabled account with a logon from yesterday looks like it was disabled by mistake.

.PARAMETER NewAccountGraceDays
    An account younger than this that is not on the roster is reported as a pending new
    hire rather than labelled a volunteer. Default 30. A hire who starts next week is not
    on this month's roster, and calling them a volunteer in the GAL is worse than leaving
    them alone for a cycle.

.PARAMETER StaleDays
    Dormancy threshold in days, default 180. Dormant non-roster accounts are still
    labelled Volunteer; they additionally get a DormancyAdvisory row in the change log
    so a long-abandoned account is visible rather than hidden behind a plausible title.
    The Entra export carries 130 accounts past this line, 72 of them over a year.

.PARAMETER HoldStaleForReview
    Opt out of labelling dormant accounts. They are logged as OffboardReview and left
    untouched instead. Off by default: staff who move to volunteering must be labelled
    so they can be told apart from staff.

.PARAMETER AlsoWriteOffice
    Additionally write the campus into physicalDeliveryOfficeName. Off by default: NCC
    records campus in the Company field, and holding the same fact in two attributes invites
    them to disagree. Existing Office values are left alone either way.

.PARAMETER RollbackFrom
    Path to a change log from an earlier -Commit run. Reverses every row whose Status is
    Applied, restoring the recorded OldValue, and exits without doing anything else.
    Combine with -Commit to actually write; without it you get the rollback plan only.

.PARAMETER Force
    During rollback, revert an attribute even if it no longer holds the value this run wrote.
    Off by default so a rollback cannot silently discard someone else's later edit.

.PARAMETER PreserveStaffDescription
    Leave the description attribute alone for roster staff. The volunteer pass still
    writes description on non-roster users.

.PARAMETER MaxChanges
    Abort before writing anything if the plan exceeds this many object modifications.
    Guards against a malformed CSV rewriting the directory. Default 1000, which sits
    above a realistic first run: roughly 300 roster users plus however many non-roster
    accounts sit in Staff\* and need the volunteer label. Later runs should be far
    smaller, since only drift gets written.

.EXAMPLE
    .\Sync-NCStaffFromRoster.ps1
    Dry run. Review RosterSync-<timestamp>-changes.csv before committing.

.EXAMPLE
    .\Sync-NCStaffFromRoster.ps1 -Commit

.EXAMPLE
    .\Sync-NCStaffFromRoster.ps1 -RollbackFrom .\RosterSync-20260730-091500-changes.csv
    Rollback plan for that run. Add -Commit to reverse it for real.

.NOTES
    Requires the ActiveDirectory module (RSAT) and delegated write rights on
    OU=Accounts,OU=ORG,DC=example,DC=local.
#>

[CmdletBinding()]
param(
    [string]  $UpdatesCsv               = (Join-Path $PSScriptRoot 'AD-Updates.csv'),
    [string]  $MappingCsv               = (Join-Path $PSScriptRoot 'AD-Name-Mapping-REVIEW.csv'),
    [string]  $LogDirectory             = $PSScriptRoot,
    [string]  $StaffOuDn                = 'OU=Staff,OU=Accounts,OU=ORG,DC=example,DC=local',
    [string]  $UpnSuffix                = 'example.org',
    [string]  $VolunteerValue           = 'Volunteer',
    [switch]  $Commit,
    [switch]  $SkipMove,
    [switch]  $SkipVolunteerPass,
    [switch]  $SkipDisabledPass,
    [switch]  $SkipDisabledAccounts,
    [switch]  $SkipGalPass,
    [string[]]$AlsoScanOus              = @('New Accounts'),
    [int]     $NewAccountGraceDays      = 30,
    [int]     $MaxDisabledMoves         = 25,
    [int]     $RecentLogonHoldDays      = 7,
    [int]     $StaleDays                = 180,
    [switch]  $HoldStaleForReview,
    [switch]  $PreserveStaffDescription,
    [switch]  $AlsoWriteOffice,
    [string]  $RollbackFrom,
    [switch]  $Force,
    [int]     $MaxChanges               = 1000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------------ setup ----
Import-Module ActiveDirectory -ErrorAction Stop

# One switch, applied consistently. Relocating a disabled account is still a change to a
# disabled account, so -SkipDisabledAccounts has to imply -SkipDisabledPass or the promise
# the parameter makes would be false.
if ($SkipDisabledAccounts) { $SkipDisabledPass = $true }

$stamp     = Get-Date -Format 'yyyyMMdd-HHmmss'
$changeLog = Join-Path $LogDirectory "RosterSync-$stamp-changes.csv"
$issueLog  = Join-Path $LogDirectory "RosterSync-$stamp-issues.csv"
$mode      = if ($Commit) { 'COMMIT' } else { 'DRY RUN' }

$changes = [System.Collections.Generic.List[object]]::new()
$issues  = [System.Collections.Generic.List[object]]::new()

function Add-Change {
    param($Sam, $Action, $Attribute, $Old, $New, $Status, $Detail = '')
    $changes.Add([pscustomobject]@{
        Timestamp = (Get-Date -Format 's'); Mode = $mode; SamAccountName = $Sam
        Action = $Action; Attribute = $Attribute
        OldValue = $Old; NewValue = $New; Status = $Status; Detail = $Detail
    })
}
# Set-StrictMode makes a missing property on an ADUser object throw. Attributes that
# have never been populated are absent from the object, not null, so read them safely.
function Get-Attr {
    param($AdObject, [string]$Name)
    if ($AdObject.PSObject.Properties.Name -contains $Name) { return [string]$AdObject.$Name }
    return ''
}
# Same idea as Get-Attr, but hands back the live object. Casting a DateTime to string and
# back reformats it through the current culture, which can throw on a non-US locale.
function Get-AttrRaw {
    param($AdObject, [string]$Name)
    if ($AdObject.PSObject.Properties.Name -contains $Name) { return $AdObject.$Name }
    return $null
}
# Import-Csv rows, and the pscustomobjects folded in beside them, do not always carry the
# same columns. Set-StrictMode turns a missing column into a terminating PropertyNotFoundStrict
# error, and inside a cmdlet scriptblock PowerShell reports it against line 1 of the script,
# which makes it near-impossible to locate. Read every CSV field through here instead.
function Get-Field {
    param($Row, [string]$Name, [string]$Default = '')
    if ($null -eq $Row) { return $Default }
    if ($Row.PSObject.Properties.Name -contains $Name) {
        $v = $Row.$Name
        if ($null -eq $v) { return $Default }
        return [string]$v
    }
    return $Default
}

# RFC 4515 escaping for values placed inside an LDAP filter.
function Escape-LdapValue {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    $out = $Value -replace '\\', '\5c'
    $out = $out -replace '\*', '\2a'
    $out = $out -replace '\(', '\28'
    $out = $out -replace '\)', '\29'
    return ($out -replace "`0", '\00')
}

# Get-ADUser -Identity raises a TERMINATING ADIdentityNotFoundException when the object does
# not exist, and -ErrorAction SilentlyContinue does not suppress it. -LDAPFilter simply
# returns nothing, which is what a lookup that is allowed to miss needs. Escaping the values
# keeps this safe for names like O'Brien that broke the old -Filter form.
#
# This builds the filter and nothing else. An earlier version of this helper returned the
# Get-ADUser results, which cost two runs: PowerShell unrolls a single-element array on the
# way out of a function, so one hit arrived as a bare ADUser, and indexing that with [0]
# quietly returns an ADPropertyValueCollection instead of failing. Guarding it with the comma
# operator then produced a nested array. Returning a string cannot unroll, so the callers
# invoke Get-ADUser themselves and wrap the result in @() where the shape is unambiguous.
#
# Both samAccountName and userPrincipalName are matched. They are usually the same string,
# but not always: the roster CSV carries the UPN prefix from the Entra export, and an account
# whose UPN changed can have a samAccountName that no longer matches it.
function Get-ADUserLdapFilter {
    param([string]$Sam, [string]$Upn)
    $clauses = @()
    if ($Sam) { $clauses += "(sAMAccountName=$(Escape-LdapValue $Sam))" }
    if ($Upn) { $clauses += "(userPrincipalName=$(Escape-LdapValue $Upn))" }
    if ($clauses.Count -eq 0) { return '' }
    if ($clauses.Count -eq 1) { return $clauses[0] }
    return ('(|' + ($clauses -join '') + ')')
}


function Add-Issue {
    param($Sam, $Severity, $Message, $Detail = '')
    $issues.Add([pscustomobject]@{
        Timestamp = (Get-Date -Format 's'); SamAccountName = $Sam
        Severity = $Severity; Message = $Message; Detail = $Detail
    })
    $color = if ($Severity -eq 'ERROR') { 'Red' } else { 'Yellow' }
    Write-Host ("  [{0}] {1} {2} {3}" -f $Severity, $Sam, $Message, $Detail) -ForegroundColor $color
}

Write-Host ''
Write-Host "=== Staff Roster -> Active Directory sync ===" -ForegroundColor Cyan
Write-Host ("Mode          : {0}" -f $mode) -ForegroundColor $(if ($Commit) { 'Red' } else { 'Green' })
Write-Host ("Domain        : {0}" -f (Get-ADDomain).DNSRoot)
Write-Host ("Staff root    : {0}" -f $StaffOuDn)
Write-Host ("Change log    : {0}" -f $changeLog)
if ($SkipDisabledAccounts) {
    Write-Host "Disabled accts: LEFT UNTOUCHED in every pass (-SkipDisabledAccounts)" -ForegroundColor Yellow
}
Write-Host ''

# Rollback reads only the change log, so do not make it depend on the roster inputs. During
# an incident the log may well be the only file to hand.
if (-not $RollbackFrom) {
    foreach ($f in @($UpdatesCsv, $MappingCsv)) {
        if (-not (Test-Path -LiteralPath $f)) { throw "Required input not found: $f" }
    }
}
try { $null = Get-ADOrganizationalUnit -Identity $StaffOuDn }
catch { throw "Staff OU not found or not readable: $StaffOuDn" }

# does the forest schema carry employeeType?
$hasEmployeeType = [bool](Get-ADObject -SearchBase (Get-ADRootDSE).schemaNamingContext `
    -LDAPFilter '(lDAPDisplayName=employeeType)' -ErrorAction SilentlyContinue)
if (-not $hasEmployeeType) {
    Add-Issue '-' 'WARN' 'employeeType absent from schema' 'FT/PT will not be written'
}

# ---------------------------------------------------------------- rollback ----
if ($RollbackFrom) {
    if (-not (Test-Path -LiteralPath $RollbackFrom)) { throw "Change log not found: $RollbackFrom" }
    Write-Host ("Rollback source: {0}" -f $RollbackFrom) -ForegroundColor Yellow
    $applied = @(Import-Csv -LiteralPath $RollbackFrom | Where-Object { $_.Status -eq 'Applied' })
    Write-Host ("{0} applied rows to reverse" -f $applied.Count)

    # Moves are reversed first, then attributes, mirroring the forward order. Restoring a
    # manager DN needs the objects back where they were before it can resolve.
    foreach ($c in ($applied | Where-Object { $_.Action -in @('MoveObject','MoveDisabled') })) {
        try {
            $dn = (Get-ADUser -Identity $c.SamAccountName -ErrorAction Stop).DistinguishedName
            if ($Commit) { Move-ADObject -Identity $dn -TargetPath $c.OldValue -ErrorAction Stop }
            Add-Change $c.SamAccountName 'RollbackMove' 'DistinguishedName' $c.NewValue $c.OldValue `
                $(if ($Commit) { 'Applied' } else { 'Planned' })
        } catch { Add-Issue $c.SamAccountName 'ERROR' 'rollback move failed' $_.Exception.Message }
    }
    $drift = 0
    foreach ($c in ($applied | Where-Object { $_.Action -in @('SetAttribute','SetVolunteer','GalHide','GalUnhide') })) {
        try {
            # Only revert if the attribute still holds what this run wrote. If someone edited
            # it in between, a blind restore would silently destroy their change. Manager is
            # exempt: the logged NewValue is a samAccountName while AD stores a DN.
            if (-not $Force -and $c.Attribute -ne 'Manager') {
                $live = Get-Attr (Get-ADUser -Identity $c.SamAccountName -Properties $c.Attribute -ErrorAction Stop) $c.Attribute
                if ($live -ne [string]$c.NewValue) {
                    Add-Change $c.SamAccountName 'RollbackSkipped' $c.Attribute $live $c.OldValue 'Skipped' `
                        "changed since this run wrote '$($c.NewValue)' - left alone; use -Force to override"
                    $drift++
                    continue
                }
            }
            if ($Commit) {
                if ([string]::IsNullOrWhiteSpace($c.OldValue)) {
                    Set-ADUser -Identity $c.SamAccountName -Clear $c.Attribute -ErrorAction Stop
                } else {
                    # msExchHideFromAddressLists is a Boolean; the log holds it as text, so
                    # restore it typed or the write is rejected.
                    $val = if ($c.Attribute -eq 'msExchHideFromAddressLists') {
                        [bool]::Parse($c.OldValue)
                    } else { $c.OldValue }
                    # Build the hashtable in two steps. A member-access expression used
                    # directly as a hashtable key is not reliably parsed.
                    $repl = @{}
                    $repl[[string]$c.Attribute] = $val
                    Set-ADUser -Identity $c.SamAccountName -Replace $repl -ErrorAction Stop
                }
            }
            Add-Change $c.SamAccountName 'RollbackAttribute' $c.Attribute $c.NewValue `
                $(if ($c.OldValue) { $c.OldValue } else { '(cleared)' }) `
                $(if ($Commit) { 'Applied' } else { 'Planned' })
        } catch { Add-Issue $c.SamAccountName 'ERROR' 'rollback set failed' $_.Exception.Message }
    }
    if ($drift) {
        Write-Host ("  {0} attributes changed since the run and were left alone (-Force overrides)" -f $drift) -ForegroundColor Yellow
    }
    $changes | Export-Csv -LiteralPath $changeLog -NoTypeInformation -Encoding UTF8
    $issues  | Export-Csv -LiteralPath $issueLog  -NoTypeInformation -Encoding UTF8
    Write-Host ''
    Write-Host ("Rollback {0}. Log: {1}" -f $(if ($Commit) { 'applied' } else { 'planned only - add -Commit to write' }), $changeLog) -ForegroundColor Green
    return
}

# ------------------------------------------------------- load roster plan ----
$rows = @(Import-Csv -LiteralPath $UpdatesCsv)
Write-Host ("Loaded {0} approved rows from AD-Updates.csv" -f $rows.Count)

# Every account the mapping file names is protected from the volunteer pass, whether or
# not it was approved. An unapproved row means "we are not sure this is the right account",
# which is emphatically not the same as "this person is not staff". Without this, the 16
# roster members whose match is still pending review would be relabelled Volunteer.
# Only the candidate on an UNAPPROVED row is protected. Entries in AltCandidates are
# rejected alternatives, and protecting those would be actively wrong: a relative may be
# listed as an alternative on his son's row, and he is precisely the person who should be
# labelled a volunteer. Same for 'jgarcia' (Jimmy, an alternative on Jason Garcia's row)
# and 'eweese' (Ethan, on Emma Weese's row). Approved rows need no protection here - they
# arrive via rosterSams once their account is confirmed to exist.
$protected = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$mapRows   = @(Import-Csv -LiteralPath $MappingCsv)
foreach ($m in $mapRows) {
    if ((Get-Field $m 'Approved') -match '^\s*[Yy]') { continue }
    $ms = Get-Field $m 'SamAccountName'
    if (-not [string]::IsNullOrWhiteSpace($ms)) { $null = $protected.Add($ms) }
}
Write-Host ("{0} unresolved accounts are protected from the volunteer pass" -f $protected.Count)

# fold in rows a human approved in the review file but that predate AD-Updates.csv
$known   = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($r in $rows) {
    $rs = Get-Field $r 'SamAccountName'
    if ($rs) { $null = $known.Add($rs) }
}
$extra = 0
foreach ($m in $mapRows) {
    if ((Get-Field $m 'Approved') -notmatch '^\s*[Yy]') { continue }
    $ms = Get-Field $m 'SamAccountName'
    if ([string]::IsNullOrWhiteSpace($ms)) {
        Add-Issue '-' 'WARN' 'approved review row has no SamAccountName' (Get-Field $m 'RosterName')
        continue
    }
    if ($known.Contains($ms)) { continue }
    # Same column set as AD-Updates.csv, OfficeName included, so downstream reads never hit a
    # missing property on these folded-in rows.
    $rows += [pscustomobject]@{
        SamAccountName = $ms
        UserPrincipalName = Get-Field $m 'MatchedUPN'
        RosterName    = Get-Field $m 'RosterName'
        ADDisplayName = Get-Field $m 'MatchedDisplayName'
        Title         = Get-Field $m 'Title'
        Department    = ''
        Campus        = Get-Field $m 'Campus'
        OfficeName    = ''
        TargetOU      = Get-Field $m 'TargetOU'
        EmployeeType  = ''
        ManagerSam    = ''
        ReportsToName = ''
        Tier          = Get-Field $m 'Tier'
        Confidence    = Get-Field $m 'Confidence'
        RosterRow     = Get-Field $m 'RosterRow'
    }
    $null = $known.Add($ms); $extra++
}
if ($extra) { Write-Host ("Added {0} extra rows approved in the review CSV" -f $extra) }

# one account claimed by two roster rows is a data error, not something to guess at
$bySam = @{}
foreach ($r in $rows) {
    $rs = Get-Field $r 'SamAccountName'
    if (-not $rs) { continue }
    $k = $rs.ToLowerInvariant()
    if (-not $bySam.ContainsKey($k)) { $bySam[$k] = [System.Collections.Generic.List[object]]::new() }
    $bySam[$k].Add($r)
}
$dupeSams = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($k in $bySam.Keys) {
    if ($bySam[$k].Count -lt 2) { continue }
    $detail = @()
    foreach ($g in $bySam[$k]) { $detail += "row $(Get-Field $g 'RosterRow') $(Get-Field $g 'RosterName')" }
    Add-Issue $k 'ERROR' 'account claimed by multiple roster rows - skipped' ($detail -join ' / ')
    $null = $dupeSams.Add($k); $null = $protected.Add($k)
}
if ($dupeSams.Count) {
    $keep = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $rows) {
        if (-not $dupeSams.Contains((Get-Field $r 'SamAccountName'))) { $keep.Add($r) }
    }
    $rows = @($keep)
}

# ----------------------------------------------------- resolve target OUs ----
$ouCache = @{}
function Resolve-CampusOu {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    if ($ouCache.ContainsKey($Name)) { return $ouCache[$Name] }
    $ou = Get-ADOrganizationalUnit -SearchBase $StaffOuDn -SearchScope OneLevel `
            -LDAPFilter "(ou=$($Name -replace '([\\()*])','\$1'))" -ErrorAction SilentlyContinue |
          Select-Object -First 1
    $ouCache[$Name] = if ($ou) { $ou.DistinguishedName } else { $null }
    return $ouCache[$Name]
}
$targetOus = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($r in $rows) {
    $t = Get-Field $r 'TargetOU'
    if ($t) { $null = $targetOus.Add($t) }
}
foreach ($campusOu in $targetOus) {
    if (-not (Resolve-CampusOu $campusOu)) {
        Add-Issue '-' 'ERROR' "target OU not found under Staff" $campusOu
    }
}

# ------------------------------------------- resolve sibling OUs of Staff ----
# Staff sits under Accounts alongside Disabled Accounts, New Accounts and the rest, so the
# parent of the Staff DN is the container to search for them.
$accountsOuDn = ($StaffOuDn -split '(?<!\\),', 2)[1]
function Resolve-SiblingOu {
    param([string]$Name)
    $esc = $Name -replace '([\\()*])', '\$1'
    $ou = Get-ADOrganizationalUnit -SearchBase $accountsOuDn -SearchScope OneLevel `
            -LDAPFilter "(ou=$esc)" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($ou) { return $ou.DistinguishedName }
    return $null
}
$disabledOuDn = Resolve-SiblingOu 'Disabled Accounts'
if (-not $disabledOuDn -and -not $SkipDisabledPass) {
    Add-Issue '-' 'ERROR' "OU 'Disabled Accounts' not found under $accountsOuDn" 'disabled pass will be skipped'
}
$scanRoots = @($StaffOuDn)
foreach ($name in $AlsoScanOus) {
    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    $dn = Resolve-SiblingOu $name
    if ($dn) { $scanRoots += $dn }
    else { Add-Issue '-' 'WARN' "OU '$name' not found under $accountsOuDn" 'not scanned' }
}
Write-Host ("Scanning {0} OU root(s): {1}" -f $scanRoots.Count, (($scanRoots | ForEach-Object { ($_ -split ',')[0] }) -join ', '))

# ------------------------------------------------------- resolve managers ----
$mgrCache = @{}
function Resolve-Manager {
    param([string]$Sam, [string]$DisplayName)
    $key = "$Sam|$DisplayName"
    if ($mgrCache.ContainsKey($key)) { return $mgrCache[$key] }
    $u = $null
    if ($Sam) {
        $f = Get-ADUserLdapFilter -Sam $Sam
        if ($f) {
            $hit = @(Get-ADUser -LDAPFilter $f -ErrorAction SilentlyContinue)
            if ($hit.Count -eq 1) { $u = $hit[0] }
        }
    }
    if (-not $u -and $DisplayName) {
        $esc = $DisplayName -replace '\\', '\5c' -replace '\*', '\2a' -replace '\(', '\28' `
                            -replace '\)', '\29' -replace "`0", '\00'
        $u = @(Get-ADUser -LDAPFilter "(|(displayName=$esc)(cn=$esc))" -ErrorAction SilentlyContinue)
        if ($u.Count -ne 1) { $u = $null } else { $u = $u[0] }
    }
    $mgrCache[$key] = $u
    return $u
}

# -------------------------------------------------- pass 1 + 2 : per user ----
Write-Host ''
Write-Host "--- Pass 1+2: attributes and campus OU placement ---" -ForegroundColor Cyan

$plan       = [System.Collections.Generic.List[object]]::new()
$rosterSams = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

$props = @('SamAccountName','UserPrincipalName','DisplayName','Title','Department',
           'Description','company','physicalDeliveryOfficeName','Manager',
           'DistinguishedName','Enabled')
if ($hasEmployeeType) { $props += 'employeeType' }

$rowNum = 0
$disabledRosterSkipped = 0
$companyPreserved = 0
$script:firstFailureDumped = $false
foreach ($r in $rows) {
  $rowNum++
  # One malformed row must not end the run. On failure, name the row and list the columns it
  # actually has - that is the information needed to fix the input.
  try {
    $sam = (Get-Field $r 'SamAccountName').Trim()
    if (-not $sam) { continue }

    $filter = Get-ADUserLdapFilter -Sam $sam -Upn (Get-Field $r 'UserPrincipalName')
    $found = if ($filter) {
        @(Get-ADUser -LDAPFilter $filter -Properties $props -ErrorAction SilentlyContinue)
    } else { @() }
    if ($found.Count -eq 0) {
        Add-Issue $sam 'ERROR' 'no AD account matches this samAccountName or UPN' `
            "roster row $(Get-Field $r 'RosterRow') $(Get-Field $r 'RosterName') - UPN tried: $(Get-Field $r 'UserPrincipalName')"
        continue
    }
    if ($found.Count -gt 1) {
        Add-Issue $sam 'ERROR' 'samAccountName and UPN resolve to different accounts - skipped' `
            (($found | ForEach-Object { Get-Attr $_ 'SamAccountName' }) -join ' / ')
        continue
    }
    # Select-Object rather than [0]: no indexing, so nothing to confuse with ADUser's own
    # property indexer.
    $user = $found | Select-Object -First 1
    $adSam = Get-Attr $user 'SamAccountName'
    if (-not $adSam) {
        # Should be impossible. If it happens, say exactly what came back rather than
        # failing with a bare property error.
        Add-Issue $sam 'ERROR' 'AD returned an object with no SamAccountName' `
            ("type=" + $user.GetType().FullName + " properties=" + `
             (($user.PSObject.Properties.Name | Sort-Object) -join ','))
        continue
    }
    if ($adSam -ne $sam) {
        Add-Issue $sam 'WARN' 'matched by UPN, not samAccountName' `
            "CSV says '$sam', AD account is '$adSam' - using the AD value"
    }
    $null = $rosterSams.Add($adSam)

    if (-not $user.Enabled) {
        if ($SkipDisabledAccounts) {
            # Still counted as roster staff so the volunteer pass never treats them as
            # non-staff, and so the pass-1 coverage guard is not skewed by the skip.
            Add-Change $sam 'DisabledSkipped' 'Title' (Get-Attr $user 'Title') `
                '(no change written)' 'Skipped' `
                'on the roster but disabled - left untouched by -SkipDisabledAccounts'
            $disabledRosterSkipped++
            continue
        }
        Add-Issue $sam 'WARN' 'account disabled but present on roster' (Get-Field $r 'RosterName')
    }
    $adUpn = Get-Attr $user 'UserPrincipalName'
    $csvUpn = Get-Field $r 'UserPrincipalName'
    if ($adUpn -and $csvUpn -and $adUpn -ne $csvUpn) {
        Add-Issue $sam 'WARN' 'UPN differs from export' "AD=$adUpn csv=$csvUpn"
    }

    # ---- attribute deltas
    $set = @{}
    # OfficeName is the expanded, address-book form of the site code. Only used
    # when -AlsoWriteOffice is set; campus normally lives in Company alone.
    $rCampus = Get-Field $r 'Campus'
    $rOffice = Get-Field $r 'OfficeName'
    $officeValue = if ($rOffice) { $rOffice } else { $rCampus }

    # NCC records campus in the Company field. The roster shorthand is expanded to the
    # convention already in AD - a short site code becomes its full Company value.
    $companyValue = Get-Field $r 'CompanyName'
    $currentCompany = Get-Attr $user 'company'

    # Central staff whose company carries a team as well as the region hold more detail than
    # the roster does: 'Global Video Team' says something 'Global Team' does not. Keep it.
    if ($companyValue -eq 'Global Team' -and
        $currentCompany -match '^\s*Global\s+\S.*\s+Team\s*$' -and
        $currentCompany.Trim() -ne 'Global Team') {
        Add-Change $sam 'CompanyPreserved' 'company' $currentCompany '(no change written)' 'Skipped' `
            "roster says Global; existing value is more specific, so it was kept"
        $companyPreserved++
        $companyValue = ''
    }
    $rTitle = Get-Field $r 'Title'
    $rEmpType = Get-Field $r 'EmployeeType'
    $desired = [ordered]@{
        Title      = $rTitle
        Department = Get-Field $r 'Department'
        company    = $companyValue
    }
    # Office duplicated the same fact in a second field, so it is no longer written by
    # default. -AlsoWriteOffice puts it back.
    if ($AlsoWriteOffice) { $desired['physicalDeliveryOfficeName'] = $officeValue }
    if (-not $PreserveStaffDescription) { $desired['Description'] = $rTitle }
    if ($hasEmployeeType -and $rEmpType) { $desired['employeeType'] = $rEmpType }

    foreach ($attr in $desired.Keys) {
        $want = [string]$desired[$attr]
        if ([string]::IsNullOrWhiteSpace($want)) { continue }
        $have = Get-Attr $user $attr
        if ($have -ne $want) {
            $set[$attr] = $want
            Add-Change $sam 'SetAttribute' $attr $have $want 'Planned'
        }
    }

    # ---- manager
    $rMgrSam = Get-Field $r 'ManagerSam'
    $rReports = Get-Field $r 'ReportsToName'
    if ($rMgrSam -or $rReports) {
        $mgr = Resolve-Manager $rMgrSam $rReports
        if (-not $mgr) {
            Add-Issue $sam 'WARN' 'manager not resolvable' "ReportsTo='$rReports'"
        }
        elseif ($mgr.DistinguishedName -eq $user.DistinguishedName) {
            Add-Issue $sam 'WARN' 'roster lists user as their own manager - skipped' $rReports
        }
        elseif ((Get-Attr $user 'Manager') -ne $mgr.DistinguishedName) {
            $set['Manager'] = $mgr.DistinguishedName
            Add-Change $sam 'SetAttribute' 'Manager' (Get-Attr $user 'Manager') $mgr.SamAccountName 'Planned'
        }
    }

    # ---- OU move
    $moveTo = $null
    $rTargetOu = Get-Field $r 'TargetOU'
    if (-not $SkipMove -and $rTargetOu) {
        $ouDn = Resolve-CampusOu $rTargetOu
        if ($ouDn) {
            $parent = ($user.DistinguishedName -split '(?<!\\),', 2)[1]
            if ($parent -ne $ouDn) {
                $moveTo = $ouDn
                Add-Change $sam 'MoveObject' 'DistinguishedName' $parent $ouDn 'Planned' `
                    "campus $rCampus"
            }
        }
    }

    if ($set.Count -or $moveTo) {
        $plan.Add([pscustomobject]@{ User = $user; Set = $set; MoveTo = $moveTo; Row = $r })
    }
  }
  catch {
    $who = if ($sam) { $sam } else { "row#$rowNum" }
    Add-Issue $who 'ERROR' 'row failed and was skipped' $_.Exception.Message
    if (-not $script:firstFailureDumped) {
        $script:firstFailureDumped = $true
        Write-Host ''
        Write-Host "--- first failure, full detail (row $rowNum, $who) ---" -ForegroundColor Magenta
        Write-Host ("  exception : " + $_.Exception.GetType().FullName)
        Write-Host ("  message   : " + $_.Exception.Message)
        Write-Host ("  at        : " + $_.InvocationInfo.PositionMessage.Trim())
        Write-Host ("  statement : " + $_.InvocationInfo.Line.Trim())
        Write-Host ''
    }
  }
}

# Pass 4 decides who is NOT staff by absence from $rosterSams. If pass 1 resolved nothing
# while the CSV had rows, that set is empty for a reason unrelated to anybody's employment,
# and carrying on would stamp Volunteer across the entire staff. Stop instead.
if ($rows.Count -gt 0 -and $rosterSams.Count -eq 0) {
    $changes | Export-Csv -LiteralPath $changeLog -NoTypeInformation -Encoding UTF8
    $issues  | Export-Csv -LiteralPath $issueLog  -NoTypeInformation -Encoding UTF8
    throw ("Pass 1 resolved 0 of $($rows.Count) roster rows, so no later pass can tell staff " +
           "from non-staff. Stopping before the volunteer pass, which would otherwise have " +
           "relabelled every account in Staff\*. See $issueLog.")
}
if ($rows.Count -gt 0 -and $rosterSams.Count -lt ($rows.Count * 0.9)) {
    Add-Issue '-' 'WARN' 'pass 1 resolved fewer than 90% of roster rows' `
        "$($rosterSams.Count) of $($rows.Count) - the volunteer pass will treat the rest as non-staff"
}

# --------------------------------------------- pass 3 : disabled cleanup ----
$disabledPlan = [System.Collections.Generic.List[object]]::new()
if (-not $SkipDisabledPass -and $disabledOuDn) {
    Write-Host ''
    Write-Host "--- Pass 3: disabled accounts out of Staff ---" -ForegroundColor Cyan

    $dProps = @('SamAccountName','DisplayName','DistinguishedName','Enabled','LastLogonDate','Title','UserPrincipalName')
    $onRosterButDisabled = 0
    $recentLogonHeld     = 0
    $systemSkipped       = 0
    $logonHoldCutoff     = (Get-Date).AddDays(-$RecentLogonHoldDays)
    foreach ($root in $scanRoots) {
        foreach ($u in (Get-ADUser -SearchBase $root -SearchScope Subtree `
                          -LDAPFilter '(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))' `
                          -Properties $dProps)) {
            if ($u.Enabled) { continue }   # belt and braces against the UAC filter

            # Disabled and on the roster is a contradiction: either the account was disabled
            # in error or the person has left and the roster is stale. Moving it would hide
            # whichever it is, so report and leave it alone.
            if ($rosterSams.Contains($u.SamAccountName)) {
                Add-Issue $u.SamAccountName 'WARN' 'disabled but on the roster - not moved' $u.DisplayName
                Add-Change $u.SamAccountName 'DisabledOnRoster' 'DistinguishedName' `
                    $u.DistinguishedName '(no change written)' 'Skipped' `
                    'account is disabled yet appears on the roster - resolve which is right'
                $onRosterButDisabled++
                continue
            }
            # Relocate only real mail-domain user accounts. Directory infrastructure is
            # disabled by design and must stay where it is: the Azure AD
            # Kerberos Server account (krbtgt_*) is disabled by design and backs Windows Hello
            # for Business cloud trust and FIDO2 sign-in to on-premises resources. Moving it
            # risks breaking passwordless authentication. It has no userPrincipalName, which
            # is the cleanest thing to test for - system objects generally do not have one.
            $dUpn = Get-Attr $u 'UserPrincipalName'
            if (-not $dUpn -or -not $dUpn.ToLowerInvariant().EndsWith("@$($UpnSuffix.ToLowerInvariant())")) {
                Add-Change $u.SamAccountName 'SystemAccountSkipped' 'DistinguishedName' `
                    $u.DistinguishedName '(no change written)' 'Skipped' `
                    ("disabled, but UPN is '" + $(if ($dUpn) { $dUpn } else { 'absent' }) +
                     "' rather than @$UpnSuffix - treated as directory infrastructure, not a departed person")
                $systemSkipped++
                continue
            }

            $parent = ($u.DistinguishedName -split '(?<!\\),', 2)[1]
            if ($parent -eq $disabledOuDn) { continue }

            # A disabled account that was in use days ago reads as disabled by accident, not
            # as a completed departure. Leave it where it is and say so.
            $ll = Get-AttrRaw $u 'LastLogonDate'
            if (($ll -is [datetime]) -and ($ll -gt $logonHoldCutoff)) {
                Add-Change $u.SamAccountName 'RecentLogonHold' 'DistinguishedName' `
                    $u.DistinguishedName '(no change written)' 'Skipped' `
                    ("disabled, but logged on " + $ll.ToString('yyyy-MM-dd') +
                     " - within $RecentLogonHoldDays days, so not relocated")
                $recentLogonHeld++
                continue
            }

            Add-Change $u.SamAccountName 'MoveDisabled' 'DistinguishedName' $parent $disabledOuDn `
                'Planned' ("last logon " + $(if ($ll -is [datetime]) { $ll.ToString('yyyy-MM-dd') } else { 'never' }))
            $disabledPlan.Add([pscustomobject]@{ User = $u; MoveTo = $disabledOuDn })
        }
    }
    # Circuit breaker. Discard the whole pass rather than a subset: a partial relocation
    # during a mass-disable event is harder to unpick than none at all.
    if ($disabledPlan.Count -gt $MaxDisabledMoves) {
        Add-Issue '-' 'ERROR' ("disabled pass abandoned: {0} accounts qualify, over the -MaxDisabledMoves limit of {1}" -f `
            $disabledPlan.Count, $MaxDisabledMoves) 'check for a runaway offboarding automation before re-running'
        foreach ($c in ($changes | Where-Object Action -eq 'MoveDisabled')) {
            $c.Status = 'Skipped'; $c.Detail = "$($c.Detail) | pass abandoned, over -MaxDisabledMoves"
        }
        $disabledPlan.Clear()
        Write-Host "  disabled pass ABANDONED - too many accounts qualified, nothing will be moved" -ForegroundColor Red
    }

    Write-Host ("  {0} disabled accounts to move into Disabled Accounts" -f $disabledPlan.Count)
    if ($systemSkipped) {
        Write-Host ("  {0} disabled objects skipped as directory infrastructure (no @{1} UPN)" -f $systemSkipped, $UpnSuffix) -ForegroundColor Yellow
    }
    if ($recentLogonHeld) {
        Write-Host ("  {0} held back: disabled but logged on within {1} days" -f $recentLogonHeld, $RecentLogonHoldDays) -ForegroundColor Yellow
    }
    if ($onRosterButDisabled) {
        Write-Host ("  {0} disabled accounts are still on the roster - reported, not moved" -f $onRosterButDisabled) -ForegroundColor Yellow
    }
}

# ------------------------------------------------ pass 4 : volunteer flag ----
$volPlan = [System.Collections.Generic.List[object]]::new()
if (-not $SkipVolunteerPass) {
    Write-Host ''
    Write-Host "--- Pass 4: volunteer flag for non-roster accounts ---" -ForegroundColor Cyan

    $vProps = @('SamAccountName','UserPrincipalName','DisplayName','Title','Description',
                'DistinguishedName','Enabled','LastLogonDate','whenCreated','GivenName','Surname')
    $staleBefore        = (Get-Date).AddDays(-$StaleDays)
    $staleCount         = 0
    $alreadyVolunteer   = 0
    $unresolvedSkipped  = 0
    $nonPersonSkipped   = 0
    $newHireSkipped     = 0
    $candidates = @()
    foreach ($root in $scanRoots) {
        $candidates += Get-ADUser -SearchBase $root -SearchScope Subtree `
                        -LDAPFilter "(&(objectCategory=person)(objectClass=user)(userPrincipalName=*@$UpnSuffix))" `
                        -Properties $vProps
    }
    $newHireCutoff = (Get-Date).AddDays(-$NewAccountGraceDays)

    foreach ($u in $candidates) {
        if (-not $u.Enabled)                        { continue }
        if ($rosterSams.Contains($u.SamAccountName)) { continue }

        # An account whose match is still unresolved is not evidence the person left staff.
        if ($protected.Contains($u.SamAccountName)) {
            Add-Change $u.SamAccountName 'UnresolvedSkipped' 'Title' (Get-Attr $u 'Title') `
                '(no change written)' 'Skipped' `
                'named in the mapping file but not approved - resolve the match, do not label'
            $unresolvedSkipped++
            continue
        }

        # These OU names sit beside Staff under Accounts rather than inside it, so this test
        # only fires on a nested exception. Kept because one misfiled service account is
        # cheaper to skip than to explain.
        if ($u.DistinguishedName -match 'OU=(Service Accounts|Administrative Accounts|Shared Local Accounts|Disabled Accounts|Disable Pending|SharedMailBoxes),') { continue }

        # A fresh account that is not on the roster is far more likely to be a hire whose
        # roster row has not landed yet than someone who moved to volunteering. Hannah
        # A hire starting next week is absent from this month's roster for that reason.
        $createdDate = Get-AttrRaw $u 'whenCreated'
        if ($createdDate -is [datetime]) {
            if ($createdDate -gt $newHireCutoff) {
                Add-Change $u.SamAccountName 'NewHirePending' 'Title' (Get-Attr $u 'Title') `
                    '(no change written)' 'Skipped' `
                    ("account created " + $createdDate.ToString('yyyy-MM-dd') +
                     ", within the $NewAccountGraceDays-day grace window - not labelled")
                $newHireSkipped++
                continue
            }
        }

        # Shared mailboxes, team aliases and room accounts are not volunteers. Labelling
        # "Info Mailbox" a Volunteer would publish that to the GAL for the whole church.
        # A real person has both a given name and a surname; these almost never do.
        $gn = Get-Attr $u 'GivenName'; $sn = Get-Attr $u 'Surname'
        $looksShared = $u.DisplayName -match '(?i)\b(team|mailbox|helpdesk|room|rooms|calendar|workflow|kiosk|shared|scan|printer|signage|display|referrals|onboarding|maintenance|support|admin|automation|robots)\b'
        if ((-not $gn) -or (-not $sn) -or $looksShared) {
            Add-Change $u.SamAccountName 'NonPersonSkipped' 'Title' (Get-Attr $u 'Title') `
                '(no change written)' 'Skipped' `
                "looks like a shared or service object (givenName='$gn' surname='$sn') - not labelled"
            $nonPersonSkipped++
            continue
        }

        # Not on the roster means not staff, so the volunteer label is written either way -
        # that distinction is the point of the pass. Dormancy is recorded alongside it as
        # an advisory, because a volunteer who has not signed in for two years is a
        # separate question from whether they are staff. lastLogonTimestamp replicates with
        # up to ~14 days of lag, immaterial against a 180-day threshold.
        $lastLogon = $u.LastLogonDate
        $lastText  = if ($lastLogon) { $lastLogon.ToString('yyyy-MM-dd') } else { 'never' }
        $isStale   = (-not $lastLogon) -or ($lastLogon -lt $staleBefore)
        if ($isStale) {
            $staleCount++
            $age = if ($lastLogon) { "$([int]((Get-Date) - $lastLogon).TotalDays) days ago" } else { 'never signed in' }
            if ($HoldStaleForReview) {
                Add-Change $u.SamAccountName 'OffboardReview' 'LastLogonDate' $lastText `
                    '(held back by -HoldStaleForReview)' 'Skipped' `
                    "dormant - last logon $age - not labelled Volunteer"
                continue
            }
            Add-Change $u.SamAccountName 'DormancyAdvisory' 'LastLogonDate' $lastText `
                '(advisory only, no change)' 'Info' `
                "labelled Volunteer, but last logon $age - check whether this account should be offboarded"
        }

        $set = @{}
        foreach ($attr in @('Title','Description')) {
            $cur = Get-Attr $u $attr
            if ($cur -ne $VolunteerValue) {
                $set[$attr] = $VolunteerValue
                Add-Change $u.SamAccountName 'SetVolunteer' $attr $cur $VolunteerValue 'Planned' `
                    $(if ($cur) { "was '$cur'" } else { 'previously blank' })
            }
        }
        if ($set.Count) { $volPlan.Add([pscustomobject]@{ User = $u; Set = $set }) }
        else { $alreadyVolunteer++ }
    }
    Write-Host ("  {0} of {1} scanned accounts are not on the roster" -f `
        ($volPlan.Count + $alreadyVolunteer), $candidates.Count)
    Write-Host ("  {0} need the volunteer flag written, {1} already carry it" -f `
        $volPlan.Count, $alreadyVolunteer)
    if ($unresolvedSkipped) {
        Write-Host ("  {0} skipped: named in the mapping file but not yet approved" -f $unresolvedSkipped) -ForegroundColor Yellow
    }
    if ($nonPersonSkipped) {
        Write-Host ("  {0} skipped: shared mailbox / team / room objects" -f $nonPersonSkipped) -ForegroundColor Yellow
    }
    if ($newHireSkipped) {
        Write-Host ("  {0} skipped: accounts newer than {1} days - likely pending new hires" -f `
            $newHireSkipped, $NewAccountGraceDays) -ForegroundColor Yellow
    }
    if ($staleCount) {
        $verb = if ($HoldStaleForReview) { 'held back from labelling' } else { 'labelled Volunteer but flagged dormant' }
        Write-Host ("  {0} of them are dormant (>{1}d) - {2}; see the change log" -f `
            $staleCount, $StaleDays, $verb) -ForegroundColor Yellow
    }
}

# --------------------------------------- pass 5 : address book visibility ----
# msExchHideFromAddressLists is the on-premises attribute. In a hybrid tenant this is the
# right place to write it: Entra Connect owns the flow into Exchange Online, and setting
# HiddenFromAddressListsEnabled directly in EXO on a synced object gets overwritten.
$galPlan = [System.Collections.Generic.List[object]]::new()
$GAL_ATTR = 'msExchHideFromAddressLists'
if (-not $SkipGalPass) {
    Write-Host ''
    Write-Host "--- Pass 5: address book visibility ---" -ForegroundColor Cyan

    $hasGalAttr = [bool](Get-ADObject -SearchBase (Get-ADRootDSE).schemaNamingContext `
        -LDAPFilter "(lDAPDisplayName=$GAL_ATTR)" -ErrorAction SilentlyContinue)
    if (-not $hasGalAttr) {
        Add-Issue '-' 'ERROR' "$GAL_ATTR is not in the schema" 'Exchange schema extensions absent - pass skipped'
    }
    else {
        $gProps = @('SamAccountName','DisplayName','DistinguishedName','Enabled',
                    'GivenName','Surname','mail','proxyAddresses',$GAL_ATTR)
        # Also sweep Disabled Accounts: an account retired before this script existed may
        # still be visible in the address book.
        $galRoots = @($scanRoots)
        if ($disabledOuDn -and $galRoots -notcontains $disabledOuDn) { $galRoots += $disabledOuDn }

        $seenGal = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $toHide = 0; $toShow = 0; $galNoMail = 0; $galShared = 0; $galDisabledSkipped = 0
        foreach ($root in $galRoots) {
            foreach ($u in (Get-ADUser -SearchBase $root -SearchScope Subtree -Filter * -Properties $gProps)) {
                if (-not $seenGal.Add($u.SamAccountName)) { continue }

                # Nothing that has no mail identity appears in the address book, so the flag
                # would be inert. Skip rather than write a meaningless value.
                $mail = Get-Attr $u 'mail'
                $prox = Get-AttrRaw $u 'proxyAddresses'
                if (-not $mail -and -not $prox) { $galNoMail++; continue }

                $curRaw = Get-AttrRaw $u $GAL_ATTR
                $isHidden = ($curRaw -eq $true)
                $curText = if ($null -eq $curRaw) { '' } else { [string]$curRaw }

                if (-not $u.Enabled) {
                    if ($isHidden) { continue }
                    if ($SkipDisabledAccounts) {
                        Add-Change $u.SamAccountName 'GalDisabledSkipped' $GAL_ATTR $curText `
                            '(no change written)' 'Skipped' `
                            'disabled and visible in the address book, but left alone by -SkipDisabledAccounts'
                        $galDisabledSkipped++
                        continue
                    }
                    Add-Change $u.SamAccountName 'GalHide' $GAL_ATTR $curText 'True' 'Planned' `
                        'account is disabled - hiding it from the address book'
                    $galPlan.Add([pscustomobject]@{ User = $u; Hide = $true }); $toHide++
                    continue
                }

                if (-not $isHidden) { continue }

                # Enabled and hidden. Staff and volunteers get unhidden; shared mailboxes,
                # team aliases and room accounts are neither, and theirs is intentional.
                $gn = Get-Attr $u 'GivenName'; $sn = Get-Attr $u 'Surname'
                $looksShared = $u.DisplayName -match '(?i)\b(team|mailbox|helpdesk|room|rooms|calendar|workflow|kiosk|shared|scan|printer|signage|display|referrals|onboarding|maintenance|support|admin|automation|robots)\b'
                if ((-not $gn) -or (-not $sn) -or $looksShared) {
                    Add-Change $u.SamAccountName 'GalLeaveHidden' $GAL_ATTR $curText '(no change written)' 'Skipped' `
                        "hidden, but looks like a shared or room object (givenName='$gn' surname='$sn') - left as is"
                    $galShared++
                    continue
                }

                Add-Change $u.SamAccountName 'GalUnhide' $GAL_ATTR $curText '' 'Planned' `
                    ('enabled ' + $(if ($rosterSams.Contains($u.SamAccountName)) { 'staff member' } else { 'volunteer' }) +
                     ' - clearing the hide flag')
                $galPlan.Add([pscustomobject]@{ User = $u; Hide = $false }); $toShow++
            }
        }
        Write-Host ("  {0} disabled accounts to hide" -f $toHide)
        Write-Host ("  {0} enabled staff/volunteers to unhide" -f $toShow)
        if ($galShared) { Write-Host ("  {0} hidden shared/room objects left alone" -f $galShared) -ForegroundColor Yellow }
        if ($galDisabledSkipped) {
            Write-Host ("  {0} disabled accounts are visible in the address book and were LEFT VISIBLE by -SkipDisabledAccounts" -f `
                $galDisabledSkipped) -ForegroundColor Yellow
        }
        if ($galNoMail) { Write-Host ("  {0} objects skipped: no mail identity, so not in the address book" -f $galNoMail) }
    }
}

# ------------------------------------------------------------------ apply ----
$totalObjects = $plan.Count + $volPlan.Count + $disabledPlan.Count + $galPlan.Count
Write-Host ''
Write-Host "--- Plan summary ---" -ForegroundColor Cyan
Write-Host ("  roster users to modify : {0}" -f $plan.Count)
Write-Host ("  volunteer flags to set : {0}" -f $volPlan.Count)
Write-Host ("  disabled accts to move : {0}" -f $disabledPlan.Count)
Write-Host ("  GAL hides / unhides    : {0} / {1}" -f `
    (@($galPlan | Where-Object Hide).Count), (@($galPlan | Where-Object { -not $_.Hide }).Count))
Write-Host ("  new hires held back    : {0}" -f (@($changes | Where-Object Action -eq 'NewHirePending').Count))
if ($SkipDisabledAccounts) {
    Write-Host ("  disabled, left alone   : {0} roster / {1} still in the address book" -f `
        $disabledRosterSkipped, (@($changes | Where-Object Action -eq 'GalDisabledSkipped').Count)) -ForegroundColor Yellow
}
Write-Host ("  dormant, flagged       : {0}" -f (@($changes | Where-Object Action -eq 'DormancyAdvisory').Count))
Write-Host ("  dormant, held back     : {0}" -f (@($changes | Where-Object Action -eq 'OffboardReview').Count))
Write-Host ("  OU moves               : {0}" -f (@($plan | Where-Object MoveTo).Count))
if ($companyPreserved) {
    Write-Host ("  company kept as-is     : {0} (more specific than 'Global Team')" -f $companyPreserved)
}
Write-Host ("  individual changes     : {0}" -f $changes.Count)
Write-Host ("  issues                 : {0} error / {1} warn" -f `
    (@($issues | Where-Object Severity -eq 'ERROR').Count),
    (@($issues | Where-Object Severity -eq 'WARN').Count))

if ($Commit -and $totalObjects -gt $MaxChanges) {
    $changes | Export-Csv -LiteralPath $changeLog -NoTypeInformation -Encoding UTF8
    $issues  | Export-Csv -LiteralPath $issueLog  -NoTypeInformation -Encoding UTF8
    throw ("Plan touches $totalObjects objects ($($plan.Count) roster, $($volPlan.Count) volunteer, " +
           "$($disabledPlan.Count) disabled moves, $($galPlan.Count) address book), " +
           "over the -MaxChanges limit of $MaxChanges. Nothing was written. Review $changeLog - " +
           "if the plan looks right, re-run with -MaxChanges $($totalObjects + 50).")
}

if ($Commit) {
    Write-Host ''
    Write-Host "--- Applying ---" -ForegroundColor Red

    # Attributes for every user first, then every move. 62 of 65 managers are themselves in
    # the update set, and the manager DN was read before any move happened. Interleaving the
    # two would hand Set-ADUser a manager DN that an earlier move had already invalidated.
    # Moves come second because 'manager' is a linked attribute: AD retargets the reference
    # by itself when the object moves.
    Write-Host "  phase 1 of 2: attributes" -ForegroundColor DarkGray
    foreach ($item in $plan) {
        if (-not $item.Set.Count) { continue }
        $sam = $item.User.SamAccountName
        try {
            Set-ADUser -Identity $item.User.DistinguishedName -Replace $item.Set -ErrorAction Stop
            foreach ($k in $item.Set.Keys) {
                ($changes | Where-Object { $_.SamAccountName -eq $sam -and $_.Attribute -eq $k -and
                                           $_.Status -eq 'Planned' }) | ForEach-Object { $_.Status = 'Applied' }
            }
        } catch {
            Add-Issue $sam 'ERROR' 'Set-ADUser failed' $_.Exception.Message
            ($changes | Where-Object { $_.SamAccountName -eq $sam -and $_.Status -eq 'Planned' -and
                                       $_.Action -eq 'SetAttribute' }) | ForEach-Object { $_.Status = 'Failed' }
        }
    }

    if ($galPlan.Count) {
        Write-Host "  phase 1b: address book visibility" -ForegroundColor DarkGray
        foreach ($item in $galPlan) {
            $sam = $item.User.SamAccountName
            $act = if ($item.Hide) { 'GalHide' } else { 'GalUnhide' }
            try {
                if ($item.Hide) {
                    $galSet = @{}; $galSet[$GAL_ATTR] = $true
                    Set-ADUser -Identity $item.User.DistinguishedName -Replace $galSet -ErrorAction Stop
                } else {
                    # Clearing is the normal state for a visible mailbox; Exchange treats an
                    # absent attribute and FALSE identically.
                    Set-ADUser -Identity $item.User.DistinguishedName -Clear $GAL_ATTR -ErrorAction Stop
                }
                ($changes | Where-Object { $_.SamAccountName -eq $sam -and $_.Action -eq $act }) |
                    ForEach-Object { $_.Status = 'Applied' }
            } catch {
                Add-Issue $sam 'ERROR' "$act failed" $_.Exception.Message
                ($changes | Where-Object { $_.SamAccountName -eq $sam -and $_.Action -eq $act }) |
                    ForEach-Object { $_.Status = 'Failed' }
            }
        }
    }

    Write-Host "  phase 2 of 2: OU moves" -ForegroundColor DarkGray
    foreach ($item in @($plan) + @($disabledPlan)) {
        if (-not $item.MoveTo) { continue }
        $sam = $item.User.SamAccountName
        try {
            Move-ADObject -Identity $item.User.DistinguishedName -TargetPath $item.MoveTo -ErrorAction Stop
            ($changes | Where-Object { $_.SamAccountName -eq $sam -and
                                       $_.Action -in @('MoveObject','MoveDisabled') }) |
                ForEach-Object { $_.Status = 'Applied' }
        } catch {
            Add-Issue $sam 'ERROR' 'Move-ADObject failed' $_.Exception.Message
            ($changes | Where-Object { $_.SamAccountName -eq $sam -and
                                       $_.Action -in @('MoveObject','MoveDisabled') }) |
                ForEach-Object { $_.Status = 'Failed' }
        }
    }

    foreach ($item in $volPlan) {
        $sam = $item.User.SamAccountName
        try {
            Set-ADUser -Identity $item.User.DistinguishedName -Replace $item.Set -ErrorAction Stop
            ($changes | Where-Object { $_.SamAccountName -eq $sam -and $_.Action -eq 'SetVolunteer' }) |
                ForEach-Object { $_.Status = 'Applied' }
        } catch {
            Add-Issue $sam 'ERROR' 'volunteer Set-ADUser failed' $_.Exception.Message
            ($changes | Where-Object { $_.SamAccountName -eq $sam -and $_.Action -eq 'SetVolunteer' }) |
                ForEach-Object { $_.Status = 'Failed' }
        }
    }
}

# -------------------------------------------------------------------- logs ----
$changes | Export-Csv -LiteralPath $changeLog -NoTypeInformation -Encoding UTF8
$issues  | Export-Csv -LiteralPath $issueLog  -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Host ("Change log : {0}" -f $changeLog) -ForegroundColor Green
Write-Host ("Issue log  : {0}" -f $issueLog)  -ForegroundColor Green
if (-not $Commit) {
    Write-Host ''
    Write-Host "Dry run only - no changes written. Re-run with -Commit to apply." -ForegroundColor Yellow
} else {
    Write-Host ''
    Write-Host "Done. Entra Connect will propagate these to Microsoft 365 on its next delta sync." -ForegroundColor Green
    Write-Host "Force it from the Connect server with: Start-ADSyncSyncCycle -PolicyType Delta"
}
