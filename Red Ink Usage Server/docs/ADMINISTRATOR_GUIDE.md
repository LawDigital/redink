# Red Ink Usage Server — Administrator Guide

**Edition:** 31 October 2026 · **Applies to:** Red Ink Usage Server 1.4 and later, including version 1.4.2  
**Audience:** IT administrators, deployment managers, licence/reporting administrators and privacy officers  
**Scope:** installation, deployment, configuration, monthly licence counting, `LogPath` analytics, security, monitoring, retention and recovery  
**This guide does not require knowledge of Python, VB.NET, SQL or application source code.**

> **Important:** The Usage Server collects and reports usage evidence; it does **not** grant, validate or enforce a Red Ink licence. Feature logs are not a substitute for monthly licence evidence. Use the approved, signed Red Ink configuration files supplied by your organisation.

## 1. What the server does

The Usage Server has three independent functions:

1. **LicenseCounter collection:** accepts monthly licence-use evidence from Red Ink for Word, Excel and Outlook, including explicit days on which an add-in was loaded. A user is counted once per month for the same licence identity even if several Office programs were used.
2. **LogPath analytics:** reads existing Red Ink diagnostic/feature log files from administrator-configured directories. It reports observed command frequency, distribution across Office hosts, days, weeks and pseudonymous log users, and explicitly recorded agent-run starts separately from tool operations.
3. **Configuration publishing:** provides controlled downloads of `redink.ini`, `allmodels.ini` and `specialservices.ini` and their optional detached `.sig` files. The server preserves their bytes; the add-ins remain responsible for any signature validation.

The dashboard also offers licence reports (including printable and CSV exports), user statistics, backup and retention controls, import-source diagnostics and Settings. The collection service works without keeping a browser open.

**Terminology:** *Monthly licence users* are distinct licence identities observed in that month, not the number of invocations or Office applications. *Daily users* are identities with an explicitly reported client-local add-in load date; they are not active sessions. *Feature calls* are logged invocations, not proof of task completion or time saved. *Agent runs* are explicitly instrumented top-level loop starts; nested tools are a separate measure.

## 2. Decide on the installation design

| Design | When to use it | Requirements |
|---|---|---|
| **Local test server** | A single administrator machine or proof of concept | Browser and add-ins on the same PC; `http://localhost:8765` or `http://127.0.0.1:8765` |
| **Organisation-wide HTTPS server** | Production with multiple PCs | A reachable DNS hostname, trusted HTTPS certificate, firewall policy and access-controlled server account |
| **Central file-share import** | Existing file-based LicenseCounter or LogPath deployments | Client write access to a share; server account **read** access to the share; periodic server scans |
| **Combined** | Collect licence evidence via HTTPS and feature logs from a share | Both mechanisms can operate together and are counted separately |

**Do not put the database on a network share.** Use a server-local writable disk for the data directory. Only input directories may be UNC shares. The browser opens at startup only in interactive mode; use headless mode for unattended operation.

## 3. Install and initialise (Windows)

### Option A — packaged Windows application

Extract the entire distributed `RedInkUsageServer` application directory, including its `_internal` (when provided), `docs` and `scripts` subdirectories. Do not run only the `.exe` while leaving the other packaged files behind.

Choose an **independent data directory**, normally `C:\ProgramData\RedInkUsageServer`, outside the program folder and outside Git. Grant write access only to the service account/administrator, and read access to authorised backup operators as appropriate.

Initialise **once** from PowerShell or Command Prompt:

```powershell
.\RedInkUsageServer.exe init --data-dir "C:\ProgramData\RedInkUsageServer"
.\RedInkUsageServer.exe check --data-dir "C:\ProgramData\RedInkUsageServer"
.\RedInkUsageServer.exe serve --data-dir "C:\ProgramData\RedInkUsageServer"
```

Initialisation displays the initial administrator password and **two distinct access tokens**. Store them in your organisation's password manager. Do not publish them in emails, tickets, source control, screenshots or web server access logs.

### Option B — Python source deployment

Python 3.11 or later is required. On the machine where the source directory is installed:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe run_server.py init --data-dir "C:\ProgramData\RedInkUsageServer"
.\.venv\Scripts\python.exe run_server.py check --data-dir "C:\ProgramData\RedInkUsageServer"
.\.venv\Scripts\python.exe run_server.py serve --data-dir "C:\ProgramData\RedInkUsageServer"
```

On an **existing installation, never run `init` again.** A server update uses the same data directory and automatically upgrades supported older SQLite schemas. Stop the server, back up its database **and** its configuration/secrets, replace the application files and start the new version. The server creates a verified pre-migration database backup when the schema needs upgrading and refuses unsupported future schema versions.

### Start without a browser (headless)

```powershell
.\RedInkUsageServer.exe serve --no-browser --data-dir "C:\ProgramData\RedInkUsageServer"
```

For Python installations, replace the executable with `.\.venv\Scripts\python.exe run_server.py` and keep the other arguments. Headless means **no browser opened**, not that the server has no HTTP interface. It still serves requests until stopped.

### Windows login autostart

The **Windows startup** controls under **Settings** can create or remove a Task Scheduler task for the **currently logged-in Windows user**, without requiring administrator rights in the usual current-user setup. The task launches the existing application headless at **Windows logon**. It does not run before logon. The scheduled process needs read/write access to the data directory and read access to its import locations. A dedicated pre-logon Windows service/account is a separate deployment choice and should be configured by IT.

The UI's **Shut down server** action requests an orderly shutdown of the current process; it does not disable future autostart. To prevent a launch on the next sign-in, **disable Windows autostart separately**. Remote browsers cannot change the local autostart or invoke the local shutdown function.

## 4. Open the dashboard and sign in

Use **http://localhost:8765** or **http://127.0.0.1:8765** for local HTTP. Some Windows systems resolve `localhost` to IPv6 `::1` before IPv4; if the former URL fails while the service is bound only to IPv4, use `127.0.0.1` or adjust the listener/network policy. Neither address makes the server available from other computers.

Use your administrator password or the optional **passwordless local sign-in** configured on Settings. Passwordless mode is allowed only for genuinely local administration on a loopback-bound service; it is not a replacement for HTTPS and controlled credentials in a shared deployment. Do not expose the desktop or local browser session to untrusted users. The passwordless setting is protected by a key in the separate `secrets.json`, not a freely editable `true` flag.

### Making the server available to multiple computers

Use a DNS name such as `usage.company.example` and TLS. Configure native HTTPS or an HTTPS reverse proxy in front of the loopback listener. Set `server.trusted_hosts` to include the hostname; ensure the server certificate covers it. Client settings **must use the address reachable from the client**, not `localhost` unless the server actually runs on that same client PC. Use the certificate configuration instructions below and your organisation’s certificate management procedures.

### Configure HTTPS for network clients

For native HTTPS, obtain a certificate for the public DNS hostname and a corresponding private key, in PEM format. Put both files under the restricted data directory (for example `C:\ProgramData\RedInkUsageServer\tls\`) and make them readable only by the server's Windows account. In **Settings**, configure a network listener, trusted hostnames, secure cookies and the native TLS certificate/key paths; save and restart the server. The equivalent values in `server.toml` are:

```toml
[server]
host = "0.0.0.0"
port = 8765
trusted_hosts = ["usage.company.example"]
secure_cookies = true
allow_insecure_remote = false

[tls]
enabled = true
certificate = "tls/fullchain.pem"
private_key = "tls/private-key.pem"
key_password_env = "REDINK_TLS_KEY_PASSWORD"
```

Use an encrypted PEM private key, with its decryption password stored separately from `server.toml` and provided through the named environment variable for the account that starts the server. The optional packaged `scripts/run-server-with-tls-password.ps1` supports a DPAPI-protected password file for unattended Windows use under the **same Windows account** that created it. A PFX/P12 certificate can be imported into encrypted PEM form with `RedInkUsageServer.exe import-pfx --data-dir ... --source ... --name ...`; the command prompts for the necessary passwords. **Do not confuse the Windows EXE signing certificate with the HTTPS server certificate.**

Alternatively, terminate HTTPS at an organisation-managed reverse proxy and keep the Usage Server bound only to `127.0.0.1`. In either case, verify trusted DNS names, firewall rules and certificate trust on the client machines before distributing the complete collector/configuration URLs.

## 5. Configure the server from Settings

**Settings** shows the effective server configuration. Click **Save configuration** and check the status message. Most import, publishing, billing and retention changes take effect immediately. Changes to the bind address, port, TLS and listener-level security **require a restart**. Unreachable import paths or missing published files produce warnings. A saved path does not cause remote PCs to send files automatically.

| Settings area | Administrator purpose | Effective when |
|---|---|---|
| `server.host`, `port`, `threads`, `trusted_hosts` | Listener address, port, allowed browser hostnames, capacity | Restart |
| `server.secure_cookies`, `allow_insecure_remote` | HTTPS cookie and remote HTTP controls | Restart |
| `tls.*` | Native HTTPS, certificate and private key | Restart |
| `collection.*` | Request size and legacy compatibility options | Restart |
| `imports.counter_paths` | Folder(s) containing file-based LicenseCounter reports | Save, then scan |
| `imports.log_paths` | Folder(s) containing SharedLogger `LogPath` files | Save, then scan |
| `imports.recursive`, `scan_seconds`, `maximum_file_bytes`, `maximum_files_per_scan` | Traversal and scan budget | Save |
| `publishing.source_directory` | Alternate folder containing approved INI files | Save |
| `authentication` | Optional passwordless localhost login, separate protected control | Save |
| `billing.unit_price`, `currency` | Optional report unit pricing; not a tax engine | Save |
| `retention.cutoff_date`, `keep_months`, `auto_cleanup` | Ingestion floor and cleaning rules | Save |
| Ingestion/configuration tokens | Separate credentials for collection and downloads | Rotation takes effect immediately |

**Never edit the internal name `INI_LogPath` in the UI.** The add-in parameter is **`LogPath`**. The Usage Server counterpart is `imports.log_paths`.

### Full publishing URLs and Copy link buttons

Settings displays and allows copying the complete URLs for:

```text
https://usage.company.example/config/<CONFIG_TOKEN>/redink.ini
https://usage.company.example/config/<CONFIG_TOKEN>/allmodels.ini
https://usage.company.example/config/<CONFIG_TOKEN>/specialservices.ini
https://usage.company.example/config/<CONFIG_TOKEN>/redink.ini.sig
https://usage.company.example/config/<CONFIG_TOKEN>/allmodels.ini.sig
https://usage.company.example/config/<CONFIG_TOKEN>/specialservices.ini.sig
```

It also shows the complete collection endpoint:

```text
https://usage.company.example/collect/<INGESTION_TOKEN>
```

These are **examples**; copy the actual addresses from Settings. Use a hostname reachable by the target add-ins. The `.sig` links work only when the corresponding approved files exist. For migrated deployments whose original tokens were stored only as hashes, the plaintext is unavailable: rotate the token intentionally and update every affected client.

### Token rotation

**Rotate ingestion token** changes the `LicenseCounterPath` credentials. **Rotate configuration token** changes all INI and `.sig` download URLs. Each newly generated token invalidates its predecessor immediately. Rotate one role at a time, capture the replacement link, update relevant clients and verify both collection and downloading. Do not rotate during an unattended rollout without a distribution plan. The collector token is not an administrator password and does not authenticate the content of a claimed licence identity.

## 6. Configure LicenseCounter in `redink.ini`

### Recommended HTTP collector configuration

For clients and server on the **same PC**, use the complete URL supplied by Settings, for example:

```ini
LicenseCounterPath=http://localhost:8765/collect/REPLACE_WITH_INGESTION_TOKEN
LicenseCounterMethod=WebExtended;RequireAck=true
LicenseCounterIntervall=360
LicenseCounterAnon=True
```

For organisation-wide clients, use your **HTTPS DNS hostname** instead of `localhost`. Treat the token as a secret. The spelling is `LicenseCounterIntervall` (double *l*, as implemented in the add-ins).

| Add-in parameter | Meaning |
|---|---|
| `LicenseCounterPath` | Enables collection when set; one endpoint, file path or a supported list of targets. Empty/disconnected disables this reporter. |
| `LicenseCounterMethod` | Reporting transport. For the Usage Server, prefer `WebExtended;RequireAck=true`; conventional methods remain for legacy receivers. |
| `LicenseCounterIntervall` | Check/retry interval **in minutes**. Default `360`; allowed `0–10080`. `0` means startup-driven only, not a continuously running timer. Positive intervals include local day-boundary checks. |
| `LicenseCounterAnon` | `True` transmits pseudonymous identity only; `False` permits the supported recoverable user identity field in licence reports. The choice does not change counting units. |
| `LicenseUserID` | The organisation's stable identifier for a licensed user, for example `%USERNAME%` where suitable. Avoid shared values that merge multiple people. |
| `LicenseProductID` | The configured/product-issued licence identity where applicable; it is not replaced or validated by the Usage Server. |

**Important current behaviour:** with `LicenseCounterPath` set, the revised add-ins record usage **regardless of `LicenseNoCheck=True`, and even if a private rather than a Pro licence is stored**. Counting is independent of entitlement. The Usage Server does **not** upgrade or grant rights. Product information must not be interpreted as a newly validated commercial licence. Older client versions may still suppress records because of their previous Pro-licence gating; update those clients if you need complete coverage.

### Confirmed-delivery protocol

`RequireAck=true` sends a JSON POST to the server. Delivery is successful **only** after an HTTP success response containing JSON `accepted: true` from a committed transaction. Timeouts and errors leave the marker pending for a later attempt. Repeated deliveries are deduplicated. For a historical event that is older than the configured server cutoff, the server can respond successfully with a **discarded** outcome; it will not add the old evidence to retained reports.

A single user in a given month, licence identity and organisation/network scope is one monthly unit across WD (Word), EX (Excel) and OL (Outlook). Additional Office hosts can update coverage without incrementing that monthly unit. Dates loaded on a client are stored as that client's calendar dates; a late September event arriving in October still belongs to September.

**What it does not mean:** a daily load is **not** an active-session estimate, an LLM request or proof of working time. A month-only legacy marker remains valid monthly evidence but does not create invented daily history.

### Alternative collection without HTTP

For an internal share, authorised administrators can configure:

```ini
LicenseCounterPath=\\fileserver\RedInkUsage\LicenseCounter
LicenseCounterMethod=File
LicenseCounterIntervall=360
LicenseCounterAnon=True
```

Then set that folder in **Settings → imports.counter_paths**; clients need write permissions, the server needs read permissions. The Usage Server also supports appropriate existing legacy file, append and web-log transport/import formats. Use the add-in's `NetworkScope=company.example` method option for classic licences if you need one consistent organisation scope across domains. A signed offline licence retains its signed scope. For transport and resubmission settings, use the LicenseCounter configuration guidance in this guide.

### Local markers, delayed reporting and diagnostics

Each client first writes local marker evidence under the Windows user's `%APPDATA%\redink` profile, and then attempts delivery to its configured targets. Marker retries survive temporary outages. A successfully acknowledged same-day host is not repeatedly billed because of a timer. Historical markers may lead to legitimate historical month rows in the dashboard. **Do not delete local markers just to remove old chart months.** Configure the server's retention cutoff and clean retained data with a verified backup.

The patched add-ins write a local licence-reporter diagnostic file:

```text
%APPDATA%\redink\licensecounter\RI_LC_Diagnostics.log
```

It reports whether the reporter started, skipped, wrote a marker or received a collector response, without exposing the collector token. Check it when a host does not appear in the report. After changing INI settings, close **all** Office processes and start the relevant host again.

## 7. Configure `LogPath` and feature analytics

`LogPath` is an **add-in log output directory**. It is not a URL and does not upload logs over HTTP. For example, a central share:

```ini
LogPath=\\fileserver\RedInkUsage\Logs
```

Ensure each user's Office process can write to that share. Then enter the matching path in **Settings → imports.log_paths** so the server can read it. If clients write only to their own `C:` drive, the Usage Server cannot read those files unless the paths are made available by an organisation-approved share or copy process. The server does not remotely browse workstations.

The analyser recognises Red Ink SharedLogger records, incrementally scans growing files and stores relevant event labels and timestamps, not raw exception stack traces or document/prompt content. New add-ins write segmented logs per month and Office host, e.g. `redink-<hash>-202610-WD.log`. Older unsplit files remain importable.

### Understanding the graphs and users page

| Measure | How to interpret it |
|---|---|
| Direct commands | Recorded Red Ink commands invoked; can include user interface actions such as Settings |
| Top-level agent run starts | **Explicit** agent-loop entry markers; internal tool calls and repeated iterations do not inflate this series |
| Tools, Skills, delegated agents | Separate counts of meaningful subordinate operations; internal housekeeping such as `tool_loader` and `context_expand` is excluded from primary activity |
| Feature groups | Related direct commands/run starts; expand or hover a group to see constituent commands and counts |
| Per calendar day | Observed primary activities divided by all calendar days in the selected range |
| Per active day | Activities divided by days with at least one observed primary event |
| Per 7 days | Seven times the per-calendar-day rate; not identical to active-week averages |
| Per active week | Activities divided by Monday–Sunday weeks in which activity is observed |
| Median/peak activity per day | Typical and maximum count among observed active days |

**Identity limitation:** `LicenseCounterAnon=False` may expose a readable licence user ID, while SharedLogger uses a **different hashed log identity**. The two records are **not automatically linked**: do not infer a named employee's feature activity from the licence-user table without a verified mapping. Even `Anon=True` is pseudonymisation, not guaranteed irreversible anonymity in every deployment. These statistics do not demonstrate staff productivity, outcome quality or completed AI tasks.

## 8. Data retention and controlling growth

The server and the clients have different storage locations and policies.

**Server database:** In **Settings → Data retention and cleaning**, choose an earliest retention date (`YYYY-MM-DD`), a rolling period in months (`keep_months`) or both; the *newer* floor applies. Preview the affected counts and explicitly confirm before deleting. `auto_cleanup` can enforce the policy on scheduled scans. Inbound offline reports **older than the effective cutoff** are acknowledged as discarded but not stored, preventing old events from reappearing after cleaning. The boundary month is evaluated using actual client dates where available; do not assume a monthly-only marker proves activity on or after the boundary day.

**Client marker files:** A reporter retains unacknowledged markers. The server cannot erase files on each employee's PC. Configure endpoint availability and retries; do not promise unlimited offline retention and bounded storage simultaneously.

**SharedLogger logs:** Monthly segmentation limits the size/time scope of individual files but not cumulative disk usage. Use the CLI `archive-logs` **preview** to identify fully imported closed monthly files, then archive only after review. An optional `--apply` saves and verifies the archive before removing eligible source files. The current month, legacy unsplit logs and incomplete/unimported segments are held. Example:

```powershell
.\RedInkUsageServer.exe archive-logs --data-dir "C:\ProgramData\RedInkUsageServer" --keep-months 6 --destination "D:\RedInkLogArchive"
# Review the output before repeating the command with --apply.
```

Archive **before** database retention deletes evidence required to verify an archive. Protect exported archives: they are not automatically encrypted. Set a separate expiry for backups and archived ZIPs under your organisation's records policy. Never delete old client data just because the usage chart already includes it.

## 9. Publishing `redink.ini` and companion files

Place organisation-approved files under the **default publishing directory**:

```text
C:\ProgramData\RedInkUsageServer\config-files
```

Or choose another existing server folder under **Settings → publishing.source_directory**, save and check the **Current publishing directory** and file list. The **Open default/current folder** buttons operate on the Windows server computer in a local interactive session. A browser on another PC cannot open the remote machine's Explorer window.

Published names are restricted to `redink.ini`, `allmodels.ini`, `specialservices.ini` and their `.sig` files. The download URL includes the **configuration token**, not the collection token. Publishing does not rewrite, generate, sign or validate INI content: prepare and verify the source files and signatures through your normal Red Ink release process. Changing the publishing directory affects subsequent downloads immediately; add-ins still follow their own configured retrieval schedule.

## 10. Reports, terminology and privacy

**Overview:** Displays monthly distinct licence identities, explicitly observed daily load identities and host coverage; old offline submissions can revise previous months.

**License reports:** Select the month/range and export CSV/JSON or print. A saved report snapshot captures what was known **at the time of export**; a later submission does not rewrite that historical snapshot. Pricing is optional and does not calculate tax.

**License scope:** The reported network/organisation grouping used as part of the unique licence identity, derived from the signed offline scope or the configured classic-licence `NetworkScope` fallback. It does **not** mean an extra licence, Office host, person or subscription tier. Two incompatible scopes can count as distinct identities; use consistent settings across a company.

**Anonymity:** `LicenseCounterAnon=True` suppresses the recoverable licence-user field in the marker/transport; `False` allows it. Restrict report/export access to staff with a legitimate need and follow applicable employment and data protection rules. Neither mode authorises guessing the relationship between the licence and LogPath pseudonyms.

## 11. Backups, upgrades and disaster recovery

1. Before a deployment or cleaning, stop or quiesce imports and create a **verified SQLite backup** using `backup --output ...`, selecting a new, unused filename. The online backup includes committed WAL transactions.
2. Back up **`server.toml`, `secrets.json`, `config-files` or the configured publishing source, and the TLS private-key files** separately with restricted access. A database backup without the credentials is not a complete disaster recovery set.
3. Keep the same **data directory** when upgrading. Start the new server; schema migrations are automatic and transactional with a verified pre-migration backup. Do not run `init` again. Never downgrade the application against a newer unsupported schema.
4. Check the dashboard, sample configuration downloads and collection endpoint, historic monthly reports, LogPath imports, and both token roles after the upgrade.
5. Securely restore a matching application/database/configuration set if recovery is required. Replacing credentials alone invalidates old tokens, even when the database is correct.

An example backup command:

```powershell
.\RedInkUsageServer.exe backup --data-dir "C:\ProgramData\RedInkUsageServer" --output "D:\SecureBackups\usage-2026-10-31.sqlite3"
```

## 12. Monitoring and troubleshooting

| Symptom | What to check |
|---|---|
| Dashboard does not load | Is the server running? Correct port? Try `http://127.0.0.1:8765/healthz`; check local firewall, host and TLS settings. |
| `localhost` fails but `127.0.0.1` works | IPv6 DNS preference versus IPv4-only listener. Use IPv4 URL or review the bind policy. |
| `LicenseCounter` shows no users | `LicenseCounterPath` configured, updated add-in running, token current, correct receiver hostname, `RequireAck=true`, no cutoff rejection. Inspect the local counter diagnostics. |
| Word appears but Excel/Outlook does not | Restart each Office host, confirm the updated add-in is loaded, and check host coverage and the diagnostic file. Multiple hosts **do not** increase monthly user totals. |
| Historical July/September rows appear | Old valid markers were delivered later; use retention Preview/Clean and cutoff if policy requires removal. Do not delete arbitrary database rows by hand. |
| LogPath statistics remain empty | `LogPath` on the clients points to a real shared/copy location; the server account can read it; Settings reports a successful source scan. |
| Save says successful but files not imported | Read the source warnings; check permissions, file patterns and scan status. A saved client path alone is not a file transfer. |
| Published INI file returns 404 | Missing or unapproved filename; missing actual file in the active publishing directory, or an unsafe symlink. |
| Published file returns 401 | Rotated/incorrect **configuration** token. Collection token cannot download INI files. |
| Collector returns 401 | Rotated/incorrect **ingestion** token. Check full endpoint in Settings. |
| Feature counts seem inflated | Distinguish direct commands, explicit agent run starts and subordinate tool calls. Missing historical run markers are **unknown**, not zero. |
| Cleaning does not remove a report | Check effective cutoff, boundary date, preview counts and confirmation. Exported external reports/backups require separate retention. |
| Windows autostart not running | User has not logged in, Task Scheduler policy blocked it, scheduled user lacks folder rights, or the server is already running. |
| Shutdown button unavailable | Requires authenticated **localhost** production server session. It does not terminate a separately managed Windows service or remote server. |

Use `check --data-dir "..."` for configuration and database integrity, and review `server.log` in the data directory. Access tokens can appear in URL strings or deployment files: redact them before sharing diagnostics. The server intentionally avoids request-path logging for token-bearing endpoints.

## 13. Administrator rollout checklist

- [ ] Select local or production HTTPS design and use a stable accessible DNS hostname.
- [ ] Install the packaged folder or a managed Python environment.
- [ ] Initialise **once**, secure the administrator password and both tokens.
- [ ] Protect the data directory, publishing folder, backup locations and SQLite database with filesystem permissions.
- [ ] Publish and verify approved INI files and signatures.
- [ ] Copy **complete** collection and configuration links from Settings; distribute through a secure channel.
- [ ] Set the four `LicenseCounter*` settings and a stable `LicenseUserID`/scope as appropriate.
- [ ] Configure a shared/copy destination for `LogPath`; set matching server import folders.
- [ ] Restart Word, Excel and Outlook; verify one monthly licence identity with three host labels, without triple-counting.
- [ ] Verify actual loaded dates and delayed offline recovery using a test environment.
- [ ] Verify direct commands, agent runs and tools are separate; check user-statistic denominator notes.
- [ ] Decide on data retention, archive policy and test a backup/restore.
- [ ] Enable headless Windows login autostart if desired; check after next sign-in.
- [ ] Confirm shutdown and restart procedures and nominate operational owners.

## 14. Reference material included in the Windows package

- `README.txt` — first installation and upgrade commands.
- `docs/ADMINISTRATOR_GUIDE.md` — this operations and configuration guide.
- `docs/USER_GUIDE.md` — interpreting dashboards, statistics and reports.
- `third-party-licenses/` — applicable third-party license notices and references.

The public Windows download intentionally excludes source files, build scripts, regression tests, internal QA reports, release engineering notes and development history. Administrative data remains in the separately configured data directory.

**Document status:** This is the practical administrator operating guide. Windows Task Scheduler, executable startup and Office deployments must be validated in the target environment. Usage evidence is not a legal determination of an organisation's licence terms or privacy duties.
