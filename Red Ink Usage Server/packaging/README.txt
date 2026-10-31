RED INK USAGE SERVER - WINDOWS APPLICATION

This archive contains the Windows application and files required to operate it.
It is not a source-code distribution and does not include development materials.

FIRST INSTALLATION
1. Extract this entire folder to a protected location, e.g. C:\Program Files\RedInkUsageServer.
   Keep RedInkUsageServer.exe and its _internal folder together (if supplied).
2. Open PowerShell in this folder and initialize a separate writable data directory ONCE:
     .\RedInkUsageServer.exe init --data-dir "C:\ProgramData\RedInkUsageServer"
   Save the generated administrator password and both tokens securely.
3. Check settings and start the server:
     .\RedInkUsageServer.exe check --data-dir "C:\ProgramData\RedInkUsageServer"
     .\RedInkUsageServer.exe serve --data-dir "C:\ProgramData\RedInkUsageServer"
4. Open http://localhost:8765 (or the configured HTTPS URL) and sign in.

UPDATING AN EXISTING INSTALLATION
Stop the server, back up the database and configuration/secrets, replace application
files, and start it using the SAME data directory. Never run init again.
Supported database changes are applied automatically on startup.

UNATTENDED USE
  .\RedInkUsageServer.exe serve --no-browser --data-dir "C:\ProgramData\RedInkUsageServer"
Windows login autostart and orderly shutdown can also be managed in Settings.

DOCUMENTATION
  docs\ADMINISTRATOR_GUIDE.md  - installation, security, LicenseCounter, LogPath,
                                  publishing, maintenance and troubleshooting
  docs\USER_GUIDE.md           - daily use of the reporting dashboard
  third-party-licenses\       - notices and license terms of redistributed components

The data directory, credentials, usage history and private keys must NOT be stored
in this application directory or published with the distribution.
