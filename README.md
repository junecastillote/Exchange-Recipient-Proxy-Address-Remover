# Remove-ExRecipientProxyAddress.ps1

## 🧭 Purpose

This PowerShell script removes **proxy email addresses** from one or more Exchange recipients. You can target these proxy addresses either by:

- Domain name(s) (e.g., remove all addresses from `@old-domain.com`), or
- Specific proxy address string(s) (e.g., remove `alias@domain.com`)

It supports multiple recipient types (mailboxes, mail users, contacts, groups, etc.) and writes a report of changes to a CSV file.

---

## ⚙️ Parameters

### `-Identity` *(Required)*

- **Type:** `string[]`
- One or more recipient identities to process (e.g., email addresses, GUIDs, or other identifiers accepted by `Get-Recipient`).

### `-DomainName`

- **Type:** `string[]`
- **Parameter Set:** `byDomain` *(default)*
- One or more domain names (e.g., `oldcompany.com`) used to match proxy addresses for removal.
- Must be accepted domains in your Exchange organization.

### `-ProxyAddress`

- **Type:** `string[]`
- **Parameter Set:** `byProxyAddress`
- List of proxy addresses (e.g., `smtp:alias@domain.com`) to match exactly for removal.

### `-TargetDomainController`

- **Type:** `string`
- Optional. Specifies a domain controller to use.

### `-OutputCsv`

- **Type:** `string`
- Optional. Full file path to write results to a CSV file.
- If not provided, a timestamped file will be saved in the script directory.

### `-ReturnResult`

- **Type:** `switch`
- If specified, the result object will be returned after execution.

### `-Quiet`

- **Type:** `switch`
- If specified, suppresses `Write-Information` messages.

---

## 🔍 Script Behavior

1. **Loads prerequisites** (`generic_functions.ps1`, recipient types list)
2. **Validates domain names** via `Get-AcceptedDomain`
3. **Iterates through each identity** using `Get-Recipient`
4. **Matches and removes proxy addresses** based on domain or exact match
5. **Applies changes** using the appropriate `Set-*` command
6. **Captures results** and writes to CSV if applicable

---

## 📤 Output CSV

If changes were made, results are written to CSV with columns:

- `DisplayName`
- `PrimarySMTPAddress`
- `RemovedProxyAddress`
- `RemainingProxyAddress`

---

## ✅ Supported Recipient Types

Loaded from `supported_recipient_types.list`, including:

- `UserMailbox`
- `SharedMailbox`
- `MailUser`
- `MailContact`
- `RoomMailbox`
- `RemoteUserMailbox`
- `MailUniversalDistributionGroup`
- `MailUniversalSecurityGroup`
- `DynamicDistributionGroup`
- *(and others)*

Unsupported types are logged and skipped.

---

## 🧪 Examples

### Remove proxies by domain

```powershell
.\Remove-ExRecipientProxyAddress.ps1 -Identity user1@domain.com,user2@domain.com -DomainName olddomain.com
```

### Remove proxies by exact address

```powershell
.\Remove-ExRecipientProxyAddress.ps1 -Identity user1@domain.com -ProxyAddress smtp:alias@olddomain.com
```

### Specify domain controller and return results

```powershell
.\Remove-ExRecipientProxyAddress.ps1 -Identity user1@domain.com -DomainName olddomain.com -TargetDomainController dc01.contoso.com -ReturnResult
```

---

## 📎 Notes

- **Primary SMTP addresses are never removed**, even if they match.
- Must be run by an account with proper permissions.
- Assumes `SayInfo` helper function is defined in `generic_functions.ps1`.
