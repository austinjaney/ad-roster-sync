# Lessons

Behaviours in PowerShell, the ActiveDirectory module and AD itself that cost real time. Kept
because none are obvious and a future change is likely to reintroduce one.

The original implementation was written on a machine with no PowerShell available, so it
reached its first dry run unexecuted. Six dry runs were needed to reach a clean first pass.
**No write ever reached the directory during any of it** — which is the whole argument for
dry-run-by-default. Six wrong attempts cost nothing but time.

---

## Things that cannot be caught by reading the code

### `Get-ADUser -Identity` throws on a miss, and `-ErrorAction SilentlyContinue` does not suppress it

`-Filter` returns `$null` quietly when nothing matches. `-Identity` raises a **terminating**
`ADIdentityNotFoundException`, and the usual suppression does not help.

Switching from `-Filter` to `-Identity` to close a quoting hole traded a quiet miss for a hard
crash. Use `-LDAPFilter` with RFC 4515 escaping: injection-safe *and* non-throwing.

### PowerShell unrolls a single-element array on function return

The most expensive defect in the project. Three separate attempts.

A helper doing `return @(Get-ADUser ...)` looks safe. It is not: on the way out of the function
a one-element array is unrolled, so a single hit arrives at the caller as a bare `ADUser`.

Then `$found[0]` **does not fail**. `ADUser` supports indexing into its own *property*
collection, so it silently returns an `ADPropertyValueCollection`. No error, wrong object, and
nothing in the source looks wrong.

Guarding with the comma operator *and* wrapping at the call site then produced a **nested**
array — `Object[]` containing `Object[]`. Belt and braces, where the braces undid the belt.

**The fix is to remove the category, not patch the instance.** Have the helper return a filter
*string* — strings cannot unroll — and call `Get-ADUser` directly at the call site where the
shape is visible. No helper should return AD objects.

Diagnosing it required making the script report the .NET type and property names of what it
actually received. Three attempts at reasoning from the source failed; one diagnostic settled
it immediately.

### `samAccountName` is not always the UPN prefix

They usually match, which makes the assumption easy to bake in. They diverge after a name
change where the UPN was updated and `sam` was not — and the roster-side identifier is
typically derived from the UPN.

Look up on **both**, report which one matched, and write to the AD value.

### Directory infrastructure looks like a departed user

A pass that relocates disabled accounts will happily pick up the Azure AD Kerberos Server
account (`krbtgt_*`): disabled by design, sitting in a staging OU, no recent logon. It backs
Windows Hello for Business cloud trust and FIDO2 sign-in to on-premises resources, so moving it
can break passwordless authentication.

No name-pattern guard catches this reliably. Require a UPN in the mail domain before relocating
anything — system objects do not have one.

---

## Things a careful review does catch

### Deriving "not staff" from an incomplete set

If the set of known staff is built only from *approved* matches, then everyone whose match is
still pending review is invisible to a pass that treats absence as "not staff". That mislabels
current employees.

*"We are unsure which account this is"* is not *"this person is not staff."*

The first fix over-corrected by also protecting rejected alternative candidates — which is
actively wrong. A relative listed as an alternative on someone else's row is precisely the
account that *should* be labelled.

### Interleaving attribute writes and object moves

If managers are themselves in the update set, manager DNs read before the run go stale as soon
as the first move lands. Write all attributes first, then all moves. Moves go second because
`manager` is a linked attribute — AD retargets the reference itself.

### OU exclusion lists that are dead code

Excluding `OU=Service Accounts` is pointless when the search base is `OU=Staff` and Service
Accounts is its *sibling*. The test never fires, and it reads like protection.

Filter on object shape instead: a real person has both `givenName` and `sn`. Shared mailboxes,
team aliases and room accounts usually have neither.

Then check the **account name** too, not just `displayName` — service accounts often have
perfectly human-looking display names.

### Rollback that overwrites blind

Commit at 09:00, someone corrects a value at 11:00, roll back at 12:00 — their edit is gone
silently. Verify the attribute still holds what the run wrote; skip and log otherwise.

### Locale-dependent date handling

Reading a `DateTime` through a stringifying helper and casting back reformats it through the
current culture, and can throw. Return the live object and test with `-is [datetime]`.

### Hashtable keys from member-access expressions

`@{ $c.Attribute = $val }` is not reliably parsed. Build it in two steps.

### `Set-StrictMode -Version Latest` plus a missing CSV column

StrictMode turns a missing property into a **terminating** error, and inside a cmdlet
scriptblock PowerShell reports it against **line 1** of the script. Unlocatable.

Keep StrictMode — it catches real bugs — but read every CSV field through a helper that returns
empty for a missing column. Also avoid `Group-Object <Property>` on heterogeneous rows for the
same reason; group by hand.

---

## Design principles that paid for themselves

- **Dry run by default, and the dry-run log *is* the change plan.** A commit performs exactly
  those rows. This is what made six failed attempts free.
- **Per-row error isolation.** A 300-row job that dies on row 1 with no context is worse than
  one that processes 299 and names the failure. Log the row number, the exception type, and the
  columns that row actually has.
- **Dependent passes must refuse to run on collapsed input.** If the matching pass resolves
  nothing, a later pass that infers "not staff" from absence will mislabel everyone. Throw
  instead. This fired for real.
- **Every skip is a logged decision.** If the tool declined to act, the reason belongs in the
  output, not in someone's memory.
- **Circuit-break the whole pass, not a subset.** A partial relocation during a mass-disable
  incident is harder to unpick than none at all.
- **Record human decisions in code, with the reasoning.** Otherwise the matcher re-offers the
  same wrong near-miss every cycle, and the reviewer re-derives the same conclusion.
