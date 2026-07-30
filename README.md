# ad-roster-sync

Reconcile Active Directory against an HR-owned staff roster, in a hybrid (Entra Connect)
environment, without a shared identifier between the two.

MIT licensed. Extracted from a working implementation that ran against a ~440-row roster and
~670 internal accounts in July 2026. All site-specific configuration and every human decision
has been replaced with commented examples — see [Adapting this](#adapting-this).

---

## The problem this solves

Your roster is the source of truth for who works here and in what role. AD has drifted from
it. You want to reconcile them, and you discover the roster carries `Last, First` and nothing
else — no UPN, no `samAccountName`, no employee number.

So before you can sync a single attribute, you have to solve a name-matching problem. And name
matching against a real directory is worse than it looks:

- Two people share a display name (a parent and child, both on staff)
- Surnames are spelled differently on each side, in both directions
- Hyphenated surnames appear truncated in AD
- Someone's surname changed and only the UPN was updated, not the `samAccountName`
- Preferred names are not misspellings: `Chris` for `Christopher` must not overwrite the legal
  name, but `Smyth` → `Smythe` should

This tooling treats that as the central problem rather than an afterthought, and — the part
that matters — **refuses to guess** when it cannot be sure.

> **If you can put a stable identifier in your roster, do that instead.** Roughly two thirds
> of `Build-ADReconciliation.py` exists to infer an identifier that could simply be written
> down. This code is what you need when you cannot.

---

## What it does

```
scripts/
  Build-ADReconciliation.py      Reads the roster + directory exports, matches people to
                                 accounts, produces an approval CSV and an enriched workbook.
                                 Never touches AD.
  Sync-ADFromRoster.ps1          Five passes against live AD. Dry run by default,
                                 -Commit to write, -RollbackFrom to reverse.
  Export-ADDisabledAccounts.ps1  Read-only. Supplies account state, which an Entra portal
                                 export cannot: it has no accountEnabled column.
  Repair-ADAccountExpires.ps1    Audits and normalises accountExpires. See below - this one
                                 is worth reading even if you use nothing else.
```

### The five passes

1. **Attributes** — `title`, `description`, `department`, `company`, `employeeType`, `manager`
2. **OU placement** — file each user under the right site OU
3. **Disabled cleanup** — disabled accounts out of the staff OUs
4. **Non-staff flag** — enabled accounts absent from the roster get a marker
5. **Address book visibility** — `msExchHideFromAddressLists`

All attributes are written before any move happens. That order is load-bearing: if managers
are themselves being relocated, interleaving hands `Set-ADUser` a manager DN that an earlier
move already invalidated.

---

## Why `accountExpires` deserves its own script

`accountExpires` is an Integer8 FILETIME, and AD accepts **two** values meaning "never
expires":

| Value | Written by |
|---|---|
| `0` | older tooling, some ADUC code paths |
| `9223372036854775807` | `Int64.MaxValue` — current ADUC "Never" |

Both are equivalent to authentication. `Get-ADUser` reports `AccountExpirationDate` as empty
for either. **They are not equivalent to a numeric filter.**

A workflow asking for "accounts whose expiry has passed" with a comparison like
`(accountExpires<=132000000000000000)` matches **every account storing `0`**, because zero is
less than any date. An identity-automation workflow scoped to one expired test account can
therefore reach every account that has no expiry at all — which is exactly how a real incident
disabled a batch of live staff accounts in a single pass.

Normalising the data reduces the exposure. It does not remove it — one non-compliant tool
writes the other sentinel back. **The durable fix is the filter:**

```
(&(accountExpires>=1)(!(accountExpires=9223372036854775807)))
```

`>=1` excludes the zero sentinel; the negation excludes MaxValue. The script prints this on
every run. Scope your workflows to an OU while you are in there — missing scope was the second
root cause.

---

## Safety design

The directory is production and the roster is imperfect, so this is built to refuse rather
than guess.

- **Dry run by default.** Nothing is written without `-Commit`. The dry-run change log is the
  complete plan — a commit performs exactly those rows.
- **Full rollback.** Every change is logged with its prior value. `-RollbackFrom <log>`
  reverses a run: moves before attributes, so restored manager DNs resolve. It refuses to
  revert an attribute that changed after the run unless `-Force`, so a rollback cannot
  silently discard someone else's later edit.
- **Circuit breakers.** `-MaxChanges` aborts the whole plan. `-MaxDisabledMoves` abandons the
  disabled pass entirely — all of it, not a subset — on the reasoning that a spike in disabled
  accounts is more likely a runaway automation than a wave of departures. A partial relocation
  mid-incident is harder to unpick than none.
- **Dependent passes refuse to run on collapsed input.** If pass 1 resolves zero roster rows,
  the script throws before pass 4 rather than concluding that every account is non-staff. This
  fired for real during development and prevented a mass mislabelling.
- **Only mail-domain accounts are relocated.** Directory infrastructure is disabled by design
  and must stay put — the Azure AD Kerberos Server account (`krbtgt_*`) backs Windows Hello for
  Business cloud trust and FIDO2 sign-in. Moving it can break passwordless auth. It has no
  UPN, which is the cleanest thing to test for.
- **Every skip is logged with a reason.** A skip is a decision, so it is recorded as one.

---

## Adapting this

Four maps at the top of `Build-ADReconciliation.py` are site configuration:

| Map | Purpose |
|---|---|
| `CAMPUS_COMPANY` | roster site code → the value written to `company` |
| `CAMPUS_DISPLAY` | roster site code → address-book short form |
| `CAMPUS_OU` | roster site code → OU name under the staff root |
| `REMAPPED` | sites with no OU of their own, so the fallback is reported |

`CAMPUS_COMPANY` and `CAMPUS_OU` can legitimately differ. A site with no OU of its own may be
filed with central staff without *being* the central site, and writing that into the address
book would be wrong.

Four more tables hold human decisions, shipped empty with commented examples:

`PINS`, `NO_ACCOUNT_ROWS`, `ROSTER_NAME_OVERRIDES`, `SECONDARY_ROLE_ROWS`.

These exist so that a decision made once survives a rebuild. Recording *why*, next to the
decision, is the point — the matcher will otherwise re-offer the same wrong near-miss every
cycle.

The PowerShell scripts default to `example.local` / `example.org` and
`OU=Staff,OU=Accounts,OU=ORG`. Override via parameters rather than editing:
`-StaffOuDn`, `-UpnSuffix`, `-AlsoScanOus`.

### Expected inputs

| File | Source |
|---|---|
| `Staff-Roster.xlsx` | your HR roster |
| `AzureAD-exportUsers.csv` | Entra portal user export |
| `All-Devices.csv` | MDM export (Kandji in the original; any tool with a user column) |
| `Passwordless-Pending.csv` | optional; group membership export |
| `AD-Disabled-Accounts.csv` | produced by `Export-ADDisabledAccounts.ps1` |

`samples/` shows the column shapes with synthetic values. **`.gitignore` blocks all real
inputs and outputs** — they carry names, UPNs, device serials and sign-in history. Keep them
beside the scripts at runtime; do not commit them.

---

## Requirements

- Python 3 with `openpyxl`
- Windows PowerShell 5.1+ with the RSAT ActiveDirectory module
- Delegated write rights on the target OUs
- Exchange schema extensions for pass 5 (`msExchHideFromAddressLists`)
- Hybrid domain — changes are written on-premises and flow outward via Entra Connect. Setting
  `HiddenFromAddressListsEnabled` directly in Exchange Online on a synced object is overwritten
  on the next sync.

---

## Lessons worth reading

`docs/LESSONS.md` covers the PowerShell and AD-module behaviours that cost the most time. The
short version, because these are not obvious and will bite anyone doing similar work:

- `Get-ADUser -Identity` raises a **terminating** error on a miss, and `-ErrorAction
  SilentlyContinue` does **not** suppress it. `-LDAPFilter` returns nothing instead.
- PowerShell **unrolls a single-element array on function return**. If a helper returns
  `@(Get-ADUser ...)`, one hit arrives as a bare `ADUser` — and indexing that with `[0]` does
  not fail, because `ADUser` supports indexing into its own *property* collection. You get an
  `ADPropertyValueCollection` and no error.
- `Set-StrictMode -Version Latest` turns a missing CSV column into a terminating error that
  PowerShell reports against **line 1** of the script. Read CSV fields through a helper.
- `msExchHideFromAddressLists` is a Boolean. A change log holds it as text, so a rollback must
  re-type it before writing.
- A hashtable key built from a member-access expression (`@{ $x.Prop = $v }`) is not reliably
  parsed. Build it in two steps.

---

## License

MIT. See `LICENSE`.
