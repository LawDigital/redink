# Red Ink Usage Server 1.4.3

Red Ink Usage Server collects monthly licence-use evidence, provides daily load counts, summarizes Red Ink `LogPath` activity, serves approved Red Ink configuration files, and exports usage reports. It does **not** grant, validate or enforce Red Ink licences.

## Information for administrators and users

- [Administrator Guide](docs/ADMINISTRATOR_GUIDE.md) — full English instructions covering installation, configuration, LicenseCounter, `LogPath`, publishing, statistics, security, backups, retention and troubleshooting.
- [User Guide](docs/USER_GUIDE.md) — interpreting monthly and daily usage, agent runs, tools, user averages and reports.
- [Third-Party License Notices](docs/THIRD_PARTY_LICENSES.md) — license and attribution information. Windows distributions also include the required original notices in `third-party-licenses/`.

## Windows installation

Download the signed **Windows distribution ZIP**, extract the entire `RedInkUsageServer` folder (including `_internal` when present), and use a separate writable data directory. For example:

```powershell
.\RedInkUsageServer.exe init --data-dir "C:\ProgramData\RedInkUsageServer"
.\RedInkUsageServer.exe check --data-dir "C:\ProgramData\RedInkUsageServer"
.\RedInkUsageServer.exe serve --data-dir "C:\ProgramData\RedInkUsageServer"
```

Initialise **only once** for a new installation. Store the displayed administrator password and the distinct reporting and configuration tokens securely. On an update, keep the existing data directory; supported database migrations run automatically.

The local dashboard is normally at <http://localhost:8765>. The web interface includes Settings, download URLs with copy controls, optional Windows-login autostart, token rotation, data cleaning and a local shutdown function. Use `serve --no-browser` for an unattended console session. Remote office clients need a reachable server hostname and appropriately configured HTTPS, not their own `localhost`.

## LicenseCounter and LogPath

Configure `LicenseCounterPath`, `LicenseCounterMethod`, `LicenseCounterIntervall` and `LicenseCounterAnon` in the add-ins' `redink.ini`. The active server normally expects `WebExtended;RequireAck=true`; copy the full collection URL from the server's authenticated Settings page. Red Ink may report usage whether a Pro, private or no verified licence is stored; counting does not change licence entitlement. Monthly users are deduplicated across Word, Excel and Outlook.

`LogPath` records feature activity on client computers. The server reads configured accessible folders or copied log files; it does not automatically download files from client machines. Tool calls and agent-run starts are shown separately and do not increase monthly licence numbers. Interpret individual statistics according to your organization's privacy policy.

## What the public Windows ZIP contains

Only the frozen Windows application, runtime libraries, operational launchers and autostart helpers, the administrator and user guides, and applicable third-party licence notices. It excludes project source files, build scripts, test suites, QA artefacts, example transport vectors, change reports, development notes and build-machine metadata. **The compiled application inherently contains executable code; it is not a source-code release.**

Do not distribute a ZIP containing deployment data, credentials, SQLite databases, local logs or private TLS/signing keys. Keep those separately secured on the destination server.
