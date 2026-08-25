# Red Ink MSI Installer

This repository contains the MSI build and deployment tooling for the Red Ink Microsoft Office add-ins for Word, Excel, and Outlook.

## What this installer produces

Red Ink is packaged separately by:

- release channel: Preview or GA
- Office application: Word, Excel, or Outlook
- Microsoft Office architecture: x86 or x64

This produces six MSI packages per channel:

- Word x86
- Word x64
- Excel x86
- Excel x64
- Outlook x86
- Outlook x64

> **Important:** x86/x64 refers to the installed Microsoft Office architecture, not the Windows architecture.

## Public release ZIP files

The publishing workflow creates one stable public ZIP per Office application. Preview uses:

- `D:\Clickonce\redink\apps\preview\redink-word-preview-msi.zip`
- `D:\Clickonce\redink\apps\preview\redink-excel-preview-msi.zip`
- `D:\Clickonce\redink\apps\preview\redink-outlook-preview-msi.zip`

GA uses:

- `D:\Clickonce\redink\apps\ga\redink-word-msi.zip`
- `D:\Clickonce\redink\apps\ga\redink-excel-msi.zip`
- `D:\Clickonce\redink\apps\ga\redink-outlook-msi.zip`

These filenames remain unchanged from release to release, so download links can stay stable.

The build still produces and validates all six x86/x64 MSI packages per channel. Public ZIP files contain the x64 package only, and each application is packaged separately to keep the download size small.

Each application ZIP contains:

- exactly one signed x64 MSI package for Word, Excel, or Outlook
- `VERSION.txt`
- `RELEASE-INFO.txt`
- `INSTALL.txt`
- `LICENSE.txt`
- `CUSTOMER-DEPLOYMENT-GUIDE.md`
- `CUSTOMER-PREREQUISITES.md`
- `CUSTOMER-SILENT-COMMANDS.md`
- `Detect-RedInk-Prerequisites.ps1`
- `SHA256SUMS.txt`

Each ZIP also has its own `.sha256` file beside it.

## Digital signatures

The MSI packages are Authenticode-signed with the installed code-signing certificate for **VISCHER AG**.

The MSI product manufacturer/provider remains **LawDigital Ltd.**

These are intentionally separate:

- **LawDigital Ltd.** identifies the product manufacturer/provider.
- **VISCHER AG** is the cryptographic publisher shown by Windows because the signing certificate is issued to VISCHER AG.

The publishing process refuses to publish MSI files that do not pass SignTool signature verification.

## Version numbers

MSI release versions are derived automatically from the current VSTO project `ApplicationVersion`.

For example:

```text
ApplicationVersion 1.6.14.200
MSI version         1.6.14
```

The MSI version is not intended to be maintained separately in the installer scripts.

`VERSION.txt` inside each published ZIP contains the resulting MSI version.

## Normal release workflow

Run:

```text
Menu.cmd
```

The menu provides the normal build, signing, verification, and publishing operations.

For a Preview release:

1. Check out the source branch/worktree containing the intended Preview release.
2. Make sure the Preview `ApplicationVersion` in the VSTO projects is correct.
3. Run `Menu.cmd`.
4. Choose **Build + sign + publish PREVIEW**.
5. Test the resulting MSI release before making it available to customers.

For a GA release:

1. Check out the source branch/worktree containing the intended GA release.
2. Make sure the GA `ApplicationVersion` in the VSTO projects is correct.
3. Run `Menu.cmd`.
4. Choose **Build + sign + publish GA**.
5. Test the resulting MSI release before making it available to customers.

The build tools do **not** switch Git branches automatically. Release commands verify the current branch before building or publishing: Preview requires `preview`, and GA requires `main`. If the wrong branch is active, the release stops before the build/publish operation proceeds.

## Updated libraries and dependencies

Normally, no installer maintenance is required when Red Ink libraries or NuGet dependencies change.

The installer build:

1. builds the current SharedLibrary and VSTO projects;
2. collects the actual files produced in the current build output;
3. preserves nested output directories;
4. packages those files into the MSI.

Therefore, when a dependency is added, removed, or updated and the project build correctly places the required runtime files in its output directory, the MSI payload is updated automatically.

Do not manually add ordinary dependency DLLs to the `.vdproj` projects.

### Manual installer changes may be required when

Installer maintenance is normally required only if a release changes something outside the ordinary project output, for example:

- a new external prerequisite must be installed separately;
- an Office add-in registration identity changes;
- installation paths or product names change;
- Preview/GA channel behavior changes;
- x86/x64 packaging requirements change;
- a required file is no longer copied to project output;
- a new native dependency has architecture-specific deployment requirements;
- MSI product/upgrade behavior itself needs to change.

If a new library is simply part of the built application output, no manual MSI dependency list should normally be necessary.

## Prerequisites

The MSI packages do not bootstrap enterprise prerequisites automatically.

Customer environments may require:

- Microsoft .NET Framework 4.8
- Microsoft Visual Studio 2010 Tools for Office Runtime
- Microsoft Office
- other prerequisites documented for the relevant Red Ink release

See:

- `CUSTOMER-PREREQUISITES.md`
- `Detect-RedInk-Prerequisites.ps1`

Enterprise administrators can deploy prerequisites using their normal management platform, such as Intune, SCCM, Group Policy, RMM tooling, or equivalent software deployment systems.

## Installation

Example silent installation:

```cmd
msiexec.exe /i "RedInk-Excel-Preview-<version>-x64.msi" /qn /norestart /L*v "%TEMP%\RedInk-Excel.log"
```

Use the x86 MSI for 32-bit Office and the x64 MSI for 64-bit Office.

See `CUSTOMER-DEPLOYMENT-GUIDE.md` and `CUSTOMER-SILENT-COMMANDS.md` for deployment details.

## Release archive

When a new application ZIP differs from its current live ZIP, the previous application ZIP is archived under:

```text
D:\Clickonce\redink\msi-backup\preview
D:\Clickonce\redink\msi-backup\ga
```

Identical historical ZIP files for the same application are not archived more than once.

## Further information

For details about producing and publishing releases, see:

```text
PUBLISHING-README.md
```
