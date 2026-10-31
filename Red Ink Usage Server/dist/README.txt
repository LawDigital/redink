WINDOWS DISTRIBUTION OUTPUT

This folder is the destination for real Windows builds. No Windows executable
is represented as having been compiled or signed in the delivered source bundle.

On a Windows build machine, run:

    .\build-windows.ps1

Output:
    dist\windows-onedir\RedInkUsageServer\RedInkUsageServer.exe
    dist\windows-onedir\RedInkUsageServer-1.1.0-windows-onedir-unsigned.zip

Optional:
    .\build-windows.ps1 -Mode onefile
    .\build-windows.ps1 -CertificateThumbprint YOUR_CERTIFICATE_THUMBPRINT

The build validates the source and frozen CLI, includes local assets and docs,
records dependency versions, and produces SHA256 hashes after optional signing.
See docs\WINDOWS.md for signing and unattended startup.

Keep production data outside the build-output folder. Build scripts refuse to
replace output containing a data directory, server configuration/credentials,
lock files, or SQLite state. Otherwise they replace only their own reproducible
windows-onedir or windows-onefile output directory.
