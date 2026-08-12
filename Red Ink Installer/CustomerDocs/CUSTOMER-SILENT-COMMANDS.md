# Red Ink silent deployment examples

Use the filenames contained in the release ZIP. `<version>` means the version in `VERSION.txt`.

## Preview example — Excel x64

`msiexec.exe /i "RedInk-Excel-Preview-<version>-x64.msi" /qn /norestart /L*v "%TEMP%\RedInk-Excel.log"`

## GA example — Excel x64

`msiexec.exe /i "RedInk-Excel-<version>-x64.msi" /qn /norestart /L*v "%TEMP%\RedInk-Excel.log"`

## Install Word + Excel

Run the appropriate Word and Excel MSI commands sequentially. Use x86 for 32-bit Microsoft Office and x64 for 64-bit Microsoft Office.

## Install all three

Install the Word, Excel and Outlook MSI files sequentially.

## Preview-to-GA promotion

For each selected application:

1. Uninstall `Red Ink for <Application> (Preview)` using its ProductCode/MSI deployment record.
2. Install `Red Ink for <Application>` GA MSI.

Do not deploy Preview and GA for the same application simultaneously.

## Exit codes

Use standard Windows Installer exit-code handling in the deployment platform. In particular, treat reboot-required MSI codes according to your organization's reboot policy rather than forcing an immediate reboot.
