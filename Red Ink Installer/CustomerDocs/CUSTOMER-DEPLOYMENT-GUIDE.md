# Red Ink Office Add-ins — Enterprise Deployment Guide

## Packages

Red Ink is supplied as separate Windows Installer packages for Word, Excel and Outlook. Install only the applications your organization wants.

For 32-bit Microsoft Office use the x86 MSI. For 64-bit Microsoft Office use the x64 MSI. The Windows operating-system bitness is not the deciding factor; Microsoft Office bitness is.

Examples:

- Word only: install Red Ink for Word MSI
- Excel only: install Red Ink for Excel MSI
- Word + Excel: install both Word and Excel MSIs
- Word + Excel + Outlook: install all three MSIs

No separate SharedLibrary MSI is required.

## Preview versus GA

Preview products appear in Windows Installed Apps as:

- Red Ink for Word (Preview)
- Red Ink for Excel (Preview)
- Red Ink for Outlook (Preview)

GA products appear as:

- Red Ink for Word
- Red Ink for Excel
- Red Ink for Outlook

Preview and GA for the same Office application are mutually exclusive. Uninstall the Preview edition before deploying GA, or uninstall GA before deploying Preview.

For managed deployment, configure supersedence in your software distribution system so the older/opposite-channel package is uninstalled first.

## Silent installation

Run from an elevated process:

`msiexec.exe /i "<package>.msi" /qn /norestart`

## Silent uninstall

Preferred enterprise method: use the MSI ProductCode recorded by your deployment system.

`msiexec.exe /x {PRODUCT-CODE-GUID} /qn /norestart`

Alternatively, invoke the original MSI package with `/x` if your deployment platform supports that workflow.

## Logging

For troubleshooting:

`msiexec.exe /i "<package>.msi" /qn /norestart /L*v "%TEMP%\RedInk-install.log"`

## Required Microsoft components

Before installation, ensure the prerequisites described in `CUSTOMER-PREREQUISITES.md` are installed.

A PowerShell detection helper `Detect-RedInk-Prerequisites.ps1` is included with the release package. It does not install or modify anything.

## Updates

For managed environments we recommend that IT deploy newer MSI versions through Intune, Microsoft Configuration Manager/SCCM, Group Policy, RMM, or equivalent software distribution.

Do not allow an application self-updater to bypass the organization's software-management policy. Red Ink can expose a policy that disables automatic MSI updating on managed machines.

## Rollout recommendation

1. Pilot on a small device/user group.
2. Confirm Office bitness.
3. Deploy/check prerequisites.
4. Uninstall the opposite Red Ink channel if present.
5. Install the required Word/Excel/Outlook MSI packages.
6. Start each selected Office application once and confirm the add-in is loaded.
7. Expand deployment to broader groups.
