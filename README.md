# M365 Admin Scripts

A collection of PowerShell scripts for Microsoft 365 and Azure AD
administration, reporting, and automation.

Built for IT admins who manage M365 tenants and want to automate
the repetitive stuff.

---

## 📁 Structure

| Folder | Purpose |
|---|---|
| `Users/` | User management and offboarding tasks |
| `Licensing/` | License reporting and assignment |
| `Reporting/` | Tenant and usage reporting |
| `Security/` | MFA, sign-in, and security reporting |

---

## ⚙️ Requirements

- Windows PowerShell 5.1+ or PowerShell 7+
- Microsoft Graph PowerShell SDK
- Run as Administrator where noted

Install the Graph module if you haven't already:

    Install-Module Microsoft.Graph -Scope CurrentUser

---

## 📌 Scripts

### Users
- **Offboard-M365User** — Revoke sessions, remove licenses, set auto-reply, and disable account
- **Get-InactiveUsers** — Find users who haven't signed in within X days

### Licensing
- **Get-LicenseReport** — Report on all assigned licenses across the tenant

### Reporting
- **Get-MailboxSizeReport** — Report on mailbox sizes across the tenant

### Security
- **Get-MFAStatusReport** — Report on MFA status for all users
- **Get-RiskySignIns** — Pull recent risky sign-in events from Azure AD

---

## 🤝 Contributing

Feel free to fork and submit PRs. Scripts should be well-commented and
include a usage example at the top.