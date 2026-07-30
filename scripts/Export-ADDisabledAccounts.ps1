<#
.SYNOPSIS
    Exports every disabled user account from example.local so the roster build can mark them.

.DESCRIPTION
    The Entra portal export carries no accountEnabled column. Its onPremisesSyncEnabled
    field describes directory synchronisation, not account state, so nothing in that export
    can tell a disabled account from an active one.

    This produces the missing input. Run it on a domain-joined workstation, put the CSV next
    to Build-ADReconciliation.py, and re-run that script. Accounts listed here are marked
    "Disabled - pending offboarding" in the roster instead of "Volunteer (dormant)", and any
    roster member who turns up disabled is flagged as a contradiction to resolve.

    Read-only. Writes one CSV and touches nothing in the directory.

.PARAMETER OutFile
    Destination CSV. Defaults to AD-Disabled-Accounts.csv beside this script. The build
    script looks for exactly that name.

.PARAMETER SearchBase
    Restrict to one subtree. Defaults to the whole domain, which is what the roster build
    expects: a disabled account already filed under Disabled Accounts still needs marking.

.EXAMPLE
    .\Export-ADDisabledAccounts.ps1

.EXAMPLE
    .\Export-ADDisabledAccounts.ps1 -OutFile 'C:\Temp\AD-Disabled-Accounts.csv'
#>

[CmdletBinding()]
param(
    [string] $OutFile    = (Join-Path $PSScriptRoot 'AD-Disabled-Accounts.csv'),
    [string] $SearchBase
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module ActiveDirectory -ErrorAction Stop

# The UAC bit test is authoritative and cheaper than filtering on the Enabled property.
$params = @{
    LDAPFilter = '(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))'
    Properties = @('SamAccountName','UserPrincipalName','DisplayName','DistinguishedName',
                   'LastLogonDate','whenChanged','Description','Title')
}
if ($SearchBase) { $params['SearchBase'] = $SearchBase }

Write-Host ''
Write-Host '=== Exporting disabled accounts ===' -ForegroundColor Cyan
Write-Host ("Domain : {0}" -f (Get-ADDomain).DNSRoot)
Write-Host ("Scope  : {0}" -f $(if ($SearchBase) { $SearchBase } else { 'entire domain' }))

$users = @(Get-ADUser @params)

$rows = foreach ($u in $users) {
    $parent = ($u.DistinguishedName -split '(?<!\\),', 2)[1]
    [pscustomobject]@{
        SamAccountName    = $u.SamAccountName
        UserPrincipalName = $u.UserPrincipalName
        DisplayName       = $u.DisplayName
        Title             = $u.Title
        Description       = $u.Description
        CurrentOU         = $parent
        LastLogonDate     = if ($u.LastLogonDate) { $u.LastLogonDate.ToString('yyyy-MM-dd') } else { 'never' }
        LastChanged       = if ($u.whenChanged)   { $u.whenChanged.ToString('yyyy-MM-dd') }   else { '' }
    }
}

$rows | Sort-Object DisplayName | Export-Csv -LiteralPath $OutFile -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Host ("{0} disabled accounts written to {1}" -f $rows.Count, $OutFile) -ForegroundColor Green

if ($rows.Count) {
    Write-Host ''
    Write-Host 'By current OU:' -ForegroundColor Cyan
    $rows | Group-Object { ($_.CurrentOU -split ',')[0] } | Sort-Object Count -Descending |
        ForEach-Object { Write-Host ("  {0,5}  {1}" -f $_.Count, $_.Name) }
}

Write-Host ''
Write-Host 'Next: copy this CSV beside Build-ADReconciliation.py and re-run it.'
