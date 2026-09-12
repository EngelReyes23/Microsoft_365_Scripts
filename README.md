# Legacy Microsoft 365 Administration Scripts

A collection of PowerShell scripts I wrote in 2023 during the early stage of my Microsoft 365 support experience. This repository is kept primarily as a historical reference and as evidence of the progression from software development into Microsoft 365 administration and automation.

> **Status:** Legacy / historical. These scripts are not actively maintained and may rely on cmdlets, modules, or administrative workflows that have changed since they were written.

## Included Scripts

| Script | Area / purpose |
| --- | --- |
| `Archive.ps1` | Exchange / mailbox archiving workflows. |
| `ChangeOfficeProductkey.ps1` | Office product key management. |
| `CreateUsers.ps1` | User creation and administrative automation. |
| `DynamicDistributionList.ps1` | Exchange Online dynamic distribution list workflows. |
| `Holds.ps1` | Hold-related administrative checks or operations. |
| `InactiveMailbox.ps1` | Inactive mailbox investigation and administration. |
| `MFA.ps1` | MFA-related administrative workflows. |
| `RetentionPolicy(MRM).ps1` | Exchange MRM retention policy workflows. |
| `SMTP.ps1` | SMTP-related configuration and troubleshooting. |

## Why This Repository Is Kept Public

This repository documents an earlier stage of my Microsoft 365 automation work. My current SharePoint Online and OneDrive for Business tooling is maintained separately in [`m365-admin-scripts`](https://github.com/EngelReyes23/m365-admin-scripts).

## Important

Do not assume these scripts are production-ready in a current Microsoft 365 tenant. Microsoft 365 services, PowerShell modules, authentication requirements, and administrative cmdlets evolve over time.

Before using any script from this repository:

1. Review every command and dependency.
2. Validate current Microsoft documentation and module requirements.
3. Test in a non-production environment.
4. Confirm that the account executing the script has only the permissions required for the intended operation.

## Disclaimer

These scripts are personal tooling and are not official Microsoft products. They are provided as historical and educational reference material without warranty.
