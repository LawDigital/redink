# Red Ink Usage Server 1.4.8

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

## Additional static company websites

Settings can define up to 32 read-only static websites with their own bind
address, port, URL prefix, server-side document root and optional SPA fallback.
Sites on the same host and port can share that port under distinct paths.
Folders automatically serve `index.html` or `index.htm`; no directory listings
or server-side code execution are enabled. Enabled HTTPS sites reuse the
existing server TLS certificate. Settings changes apply through the supervised
listener switch without a manual process restart, with rollback if a new port
cannot be bound. The additional sites never expose the Usage Server's admin,
collector or token/configuration routes.

**Security:** Static websites do not use the Usage Server admin login. Any
browser able to reach a site's port can access its published files. On the
**same hostname and port as the administrator**, a JavaScript app has the
same browser origin and can access an authenticated administrator's API;
therefore share the admin origin **only with fully trusted application code**.
Use a separate hostname or port for other apps. Publish
only vetted public-to-the-intranet assets, limit each port with Windows and
network firewalls, and keep site roots outside the protected Usage Server data
directory. The root must be an absolute path on the machine running the
server. To serve applications needing back-end code, authentication, upload,
or API execution, deploy a separate application service/reverse proxy.

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

### Websites on the Usage Server port and per-site Save

An enabled static website can now use the same bind address, TCP port and HTTP/HTTPS
mode as the Usage Server (for example, `https://server:8765/apps/`). All sites on
that shared listener are routed by URL prefix; `/` and the administrator, login,
collector, configuration, health and static-asset namespaces remain exclusively
assigned to the Usage Server. To publish at `/` use a dedicated port.

In **Settings → Additional static websites**, each website now has its own
**Save website** button (and a persistent Remove action). The global Save
configuration button manages only general server settings. A website save
validates and applies only that website together with all previously saved
websites, without capturing edits to other unsaved cards. Live listener changes
are activated in the background; the card shows the application/rollback state.
Existing separate-port websites remain unchanged.
