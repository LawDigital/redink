# Red Ink Semantic Archive Worker

## Quick start

For a normal local Red Ink installation no `--ini` argument is required. The worker resolves the same active `redink.ini` source as the Office add-ins and synchronizes the configured Semantic Archive library before archive selection.

```powershell
redink-sa-worker.exe refresh "DP Know-how"
```

This selects the archive by its unique display name (a stable archive ID is also accepted), processes all documents and drains the refresh operation until completion or deferred work. `reindex`, `extract`, `retry` and `permissions` can be used instead of `refresh`. Advanced flags remain available through `--help`.


The same source filtering, extraction/OCR coverage, storage, permission and indexing rules apply in Office and this worker. Keep `redink-sa-catalog.json`; generated Semantic Archive state lives under `sa-archives`.

This Windows .NET Framework 4.8 console host calls the same `SemanticArchiveBuilder`, catalog, exporter, model resolver and permission-reconciliation APIs used by Red Ink. It does not automate Word or Outlook. Run it under the Windows account whose configured source and generated-output permissions should apply.

Build **Red Ink Semantic Archive Worker** in the solution, or build the existing `SemanticArchive.Worker.vbproj` path, with the Windows Visual Studio/MSBuild toolchain and restored solution NuGet packages. The executable is **`redink-sa-worker.exe`**, produced under `bin\Debug` or `bin\Release`; its generated configuration file is named `redink-sa-worker.exe.config`. The source folder, project path and project GUID stay unchanged so existing project references remain valid. Deploy its complete output directory, including the shared-library dependencies. A working, licensed Red Ink configuration must already exist. Configure its archive catalog with `SemanticArchiveCatalogPathLocal=%LOCALAPPDATA%\RedInk\SA`. Previous development keys `SemanticArchiveCatalogPath` and `SemanticIndexArchivePath` must be renamed to `SemanticArchiveCatalogPathLocal`; they are not read aliases. A blank canonical value disables Semantic Archives. By default the worker resolves the active configuration source with the same registry/default-path precedence as Red Ink. `--ini` remains available to override that source explicitly. The worker never changes the Office configuration-source registry setting or starts the setup wizard.

```powershell
# One bounded refresh batch for all documents in one archive.
.\redink-sa-worker.exe --ini C:\RedInk\redink.ini --once --archive archive_id --all-documents

# Drain selected documents with bounded discovery and batches.
.\redink-sa-worker.exe --ini C:\RedInk\redink.ini --loop --archive archive_id --document document_id --operation reindex --batch-files 4 --discovery-entries 250 --discovery-seconds 5 --max-cycles 10 --max-seconds 600

# Large existing archives: keep discovery resident and drain cached imports promptly.
.\redink-sa-worker.exe --ini C:\RedInk\redink.ini --loop --all-archives --all-documents --batch-files 128 --interval-seconds 1 --discovery-entries 1000 --discovery-seconds 10 --max-seconds 3600

# Reconcile generated-artifact permissions without model configuration, OCR or LLM calls.
.\redink-sa-worker.exe --ini C:\RedInk\redink.ini --loop --all-archives --all-documents --operation permissions --max-seconds 300
```

Archive arguments accept either a stable archive ID or an exact, unique archive display name; filesystem paths are not accepted. Document arguments remain stable IDs. Advanced scope can use repeated `--archive` arguments or `--all-archives`; when no document option is supplied, all documents are selected automatically. Selected document IDs require one explicitly selected archive. `--all-archives` means every registered archive, including archives disabled for automatic processing: these are explicit manual commands. Background-enabled flags do not narrow that explicit selection. Rights reconciliation can also repair artifacts for disabled archives.

| Operation | Work requested |
| --- | --- |
| `refresh` (default) | Reconcile the selected scope and process changed or new source jobs. |
| `retry` | Retry failed jobs within the selected scope. |
| `reindex` | Rebuild semantic metadata while reusing suitable extraction artifacts. |
| `extract` | Force source extraction and subsequent semantic processing; configured per-root OCR settings apply. |
| `permissions` | Reconcile generated-artifact protection against current source permissions; no model setup or content processing. |

`--once` performs one bounded batch per selected archive. The default mode (and `--loop`) drains that same operation with a default 1-second cancellable delay between cycles; it continues past individually deferred jobs while runnable content or discovery remains, and exits when complete or when only deferred/host-dependent work remains, another writer blocks the archive, or no progress is possible. It does not continuously start new indexing operations. Schedule fresh refresh or permission commands externally for periodic maintenance. For very large directories, prefer one resident `--loop` process: restarting short `--once` processes repeatedly can spend their discovery time re-seeking a native directory iterator. The worker defaults (64 files per processing batch, up to 5000 discovery entries / 60 seconds per discovery pass, and 1 second between cycles) are intended for an explicit unattended indexing run and are deliberately faster than Office background processing. Use a smaller `--batch-files` value when model/OCR load should be throttled, or larger values such as 128/256 for large cached imports after choosing suitable resource limits.

Each invocation logs an `operationId`. To resume a bounded, cancelled or deferred command, supply that GUID through `--operation-id` and keep the exact archive/document scope and operation flags. Durable checkpoints prevent a repeated ID from forcing already completed discovery again. To intentionally start a fresh reconciliation/rebuild, omit `--operation-id`. Do not reuse an ID for a different scope or operation. The runner retains the same ID and options throughout its own drain loop.

Worker option defaults are defined centrally in `SharedLibrary/Code/Core/Definitions/SharedMethods.Constants.vb` and consumed by the parser; omission retains the same conservative behavior. An archive scope is still required, but it may be supplied by the quick positional syntax; document scope defaults to all documents.

Controls are `--batch-files 1..4096` (default 64), `--discovery-entries 1..10000` (default 5000), `--discovery-seconds 1..120` (default 60), `--interval-seconds 1..86400` (loop only, default 1), `--max-cycles 1..1000000` (loop only), and `--max-seconds 1..604800`. Writer contention yields after 250 ms without opening the work queue or resolving a model. Ctrl+C requests cancellation at the builder's safe checkpoints. Completed checkpoints remain available; cancellation does not authorize publishing an incomplete generation.

Interactive progress is rendered as one self-overwriting line on stderr, while lifecycle/batch events remain JSON lines on stdout. This keeps human progress visible without breaking scripts that parse stdout. `Ctrl+C` requests graceful cancellation and the worker prints that it is stopping at a safe checkpoint. `WORKER_HELP.txt` is copied next to the executable and contains the operational quick reference.

Logs are JSON lines on stdout. Optional `--log` writes the same lines to a **new** file outside all configured source roots and the archive control directory, creating/validating a private Windows ACL for its directory/file. Existing files are never appended or overwritten; choose a fresh filename for each invocation. Mapped-drive and UNC aliases are compared through physical paths. An unavailable/unverifiable root or unsafe log path fails with `log_initialization_failed`; retry without `--log` to use stdout alone. Logs contain operation/archive IDs, bounded-work counts and fixed diagnostic codes. They omit source paths, content, configuration URLs, provider responses, API keys, and license values. Handle stdout redirection with the same care as other administrative logs. Full operational diagnostics remain in the normal protected archive state/UI.

| Exit code | Meaning |
| --- | --- |
| `0` | Selected operation completed. |
| `1` | Configuration, license, authentication or fatal runtime failure. |
| `2` | Invalid/ambiguous command line; use `--help`. |
| `3` | Partial, pending, deferred or failed archive work; use the logged ID to resume. |
| `130` | Cancelled by Ctrl+C or the time limit. |

The worker never opens a dialog, browser, device-code flow, OAuth listener or license-activation form. A configuration/authentication path requiring interaction fails explicitly (`headless_interaction_required` or `noninteractive_auth_required`). Normal licensing still applies, including its established validation and state behavior. Rights-only commands skip model setup after the same license decision. Per-user automatic settings use `SemanticArchiveBackgroundIndexingEnabled` (default `False`), `SemanticArchivePermissionMaintenanceEnabled` (default `True`), `SemanticArchiveBackgroundIndexingWindow` and `SemanticArchivePermissionMaintenanceWindow`; explicit worker commands do not depend on those automatic-maintenance switches. Existing legacy preferences are read compatibly, and an explicit settings save persists the canonical names with `SemanticArchiveSettingsVersion=1`. API-key and supported service-account authentication use the existing implementations; already valid in-process OAuth cache entries can be reused. A separate worker process does not inherit an Office process's token cache, and this project does not add an interactive-token refresh mechanism.

`.doc` and other inputs requiring a host reader remain `PendingHost`; open/process them through a supported interactive host. The worker supplies no `HostReaderDispatcher`. It indexes as its actual Windows identity and does not create requester search grants. AutoPilot and scheduled search requests still require independently verified, request-bound authority; a worker command or claimed sender identity cannot grant those rights.

Windows compilation and execution are required to validate ACL, Office-independent loading, model authentication and cancellation in the deployment environment. Temporary verification may be run outside the product repository; it is not a substitute for the required Windows build and runtime validation.
