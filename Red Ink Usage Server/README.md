# Red Ink Usage Server 1.4.6

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

The local dashboard is normally at <http://localhost:8765>. The web interface includes Settings, download URLs with copy controls, optional Windows-login autostart, token rotation, data cleaning and authenticated shutdown from both local and remote admin sessions. Use `serve --no-browser` for an unattended console session. Remote office clients need a reachable server hostname and appropriately configured HTTPS, not their own `localhost`.

## LicenseCounter and LogPath

Configure `LicenseCounterPath`, `LicenseCounterMethod`, `LicenseCounterIntervall` and `LicenseCounterAnon` in the add-ins' `redink.ini`. The active server normally expects `WebExtended;RequireAck=true`; copy the full collection URL from the server's authenticated Settings page. Red Ink may report usage whether a Pro, private or no verified licence is stored; counting does not change licence entitlement. Monthly users are deduplicated across Word, Excel and Outlook.

`LogPath` records feature activity on client computers. The server reads configured accessible folders or copied log files; it does not automatically download files from client machines. Tool calls and agent-run starts are shown separately and do not increase monthly licence numbers. Interpret individual statistics according to your organization's privacy policy.

## What the public Windows ZIP contains

Only the frozen Windows application, runtime libraries, operational launchers and autostart helpers, the administrator and user guides, and applicable third-party licence notices. It excludes project source files, build scripts, test suites, QA artefacts, example transport vectors, change reports, development notes and build-machine metadata. **The compiled application inherently contains executable code; it is not a source-code release.**

Do not distribute a ZIP containing deployment data, credentials, SQLite databases, local logs or private TLS/signing keys. Keep those separately secured on the destination server.

### Administrator sign-in on a trusted device

The login form offers **Remember this browser for 30 days** (unchecked by default).
This stores a random HttpOnly, SameSite=Strict cookie in the browser and only
a SHA-256 digest in the server database. It works after restarting the server.
Logout revokes the device token; changing the administrator password also
invalidates older tokens. Only use on trusted personal browsers; keep TLS
enabled for network access. Passwordless localhost logins are never remembered.

### Source archive for support

Run `scripts\create-source-zip-on-desktop.bat` from the checked-out
source project. The resulting ZIP contains source, build scripts, assets,
tests and documentation, **not** running instance secrets, configuration,
user databases, keys, certificates, logs or compiled distributions.
The desktop folder is resolved by Windows (including OneDrive redirection).

### Observed Office add-in versions

The Function activity view shows **distinct SharedLogger users for each Office
add-in and latest observed version**, rather than a count of log entries.
The version panel has a separate time window (7 days by default; optionally
14 or 30 days). A user with records from several versions of Word, Excel,
or Outlook contributes only once to each host. Missing version information
is shown as "Version unavailable" for an otherwise observed add-in. A dated `V.DDMMYY` version
is ordered by its embedded date, and a numeric `Vx.y.z` version is ordered
numerically. Incomparable formats fall back to the last observation.
This analysis is based only on imported LogPath records; it is not a live
installation inventory or a licence billing statistic.

### Chart colours and Office version grouping

Chart bars, activity lines, ranking bars and ordinary accent-coloured text
retain the brighter dashboard colours. The exact red sampled from the Red Ink
logo remains reserved for headings, input/control outlines and decorative
brand accents, including the administrator login.

The observed-version panel groups **Word**, **Excel** and **Outlook** separately,
showing each host's distinct observed user total and its version distribution
for the chosen 7-, 14- or 30-day window. The one-version-per-user-and-host
rule and the independence from monthly licence reporting are unchanged.
