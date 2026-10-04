# Red Ink Semantic Archives — Technical Documentation

**Audience:** Red Ink administrators, power users, support staff, publishers of shared archives, and users who consume Semantic Archives.  
**Implementation basis:** current Semantic Archive source state from this conversation, 4 October 2026.  
**Scope:** filesystem-based Semantic Archives in Red Ink Gen2, including local catalogs, central library subscriptions, extraction/OCR, shared artifacts, retrieval, automatic maintenance, AutoPilot integration, and the standalone worker.

---

## 1. What Semantic Archives are

Semantic Archives (SA) let Red Ink make a set of files searchable through a bounded semantic retrieval layer while preserving the original files as the source of authority.

An archive is defined by:

- a stable archive ID;
- a human-readable name and description;
- one or more source roots;
- extraction, OCR and file-type rules;
- semantic-routing settings and retrieval budgets;
- optional automatic-processing settings;
- optional central-library publication/subscription metadata.

The main design principle is that **the original file remains authoritative**. Semantic Archive text extracts, cards, indexes and routing metadata are derivatives. They may accelerate retrieval but never replace the current access check against the source.

The feature is available only when `SemanticArchiveCatalogPathLocal` is configured. A configured central library alone does not enable Semantic Archives.

---

## 2. Components

### 2.1 Local catalog and private state

The per-user local catalog is configured with:

```ini
SemanticArchiveCatalogPathLocal=<directory>
```

The value is a directory. Red Ink stores:

```text
<SemanticArchiveCatalogPathLocal>\
    redink-sa-catalog.json
    sa-archives\
        catalog.lock
        <archive-id>\
            current.json
            writer.lock
            writer-fence.json
            generations\
            state\
            work\
```

The local catalog contains archive definitions and local subscription state. The `sa-archives` tree contains private generation metadata, durable queues, checkpoints and other personal processing state.

The default setting is empty, which disables SA.

### 2.2 Central catalog library

The optional central publication library is configured with:

```ini
SemanticArchiveCatalogLibraryPath=<directory or UNC path>
```

It contains one bounded JSON descriptor per published archive. It is not the user's index store and does not contain original documents.

The library is used for:

- publishing a locally created archive definition;
- updating a published definition;
- withdrawing a definition;
- automatically subscribing users who can read a published descriptor;
- synchronizing published revisions into the local catalog.

The library requires a configured local catalog. A library path without `SemanticArchiveCatalogPathLocal` does not activate SA.

### 2.3 Shared per-document artifacts

Reusable extracted text and document indexes can be stored near the source or in an explicitly configured shared artifact root. The implementation uses a `.redink-sa` namespace.

The shared artifact layer is separate from the user's private local catalog. Shared artifacts are intended to avoid repeated extraction/OCR/index work between authorized users.

If safe shared publication is not possible, Red Ink can fall back to a private derivative location.

### 2.4 Worker

`redink-sa-worker.exe` is a .NET Framework 4.8 console application that executes the same Semantic Archive build logic without requiring Word or Outlook UI.

It supports:

- discovery;
- extraction and OCR;
- semantic metadata generation;
- indexing;
- publication;
- permission reconciliation;
- central-library synchronization.

---

## 3. Global configuration

The important global settings are:

| Setting | Purpose |
|---|---|
| `SemanticArchiveCatalogPathLocal` | Required local catalog directory. Empty disables SA everywhere. |
| `SemanticArchiveCatalogLibraryPath` | Optional central publication/subscription library. |
| `SemanticArchiveBackgroundIndexingEnabled` | Personal default for automatic SA content processing. |
| `SemanticArchiveBackgroundIndexingWindow` | Optional local-time allow/deny window. |
| `SemanticArchivePermissionMaintenanceEnabled` | Controls automatic permission reconciliation. |
| `SemanticArchivePermissionMaintenanceWindow` | Optional local-time window for permission maintenance. |

All SA filesystem paths are expected to pass through Red Ink's existing `SharedMethods.ExpandEnvironmentVariables()` behavior and subsequent SA path validation.

Supported placeholders therefore include the normal environment variables plus Red Ink-specific convenience placeholders such as:

```text
%APPDATA%
%LOCALAPPDATA%
%PROGRAMDATA%
%USERPROFILE%
%DESKTOP%
%DOCUMENTS%
%MYDOCUMENTS%
%DOWNLOADS%
%TEMP%
%PROGRAMFILES%
%COMMONDOCUMENTS%
```

Unresolved `%...%` placeholders are rejected rather than silently treated as literal paths.

---

## 4. Archive definition

Each archive contains these principal fields:

- `ArchiveId`
- `SemanticArchiveName`
- `SemanticArchiveDescription`
- `SemanticArchiveEnabled`
- `SemanticArchiveRoots`
- `SemanticArchiveBackgroundEnabled`
- `SemanticArchiveBackgroundWindow`
- `SemanticArchiveSourceIndexThresholdBytes`
- `SemanticArchiveMaxChildrenPerNode`
- `SemanticArchiveMaxRoutingCharacters`
- `SemanticArchiveExtractionProfileVersion`
- `SemanticArchiveSemanticProfileVersion`
- `SemanticArchiveAllowPartialSearch`
- `SemanticArchiveRetrievalBudgets`
- optional `SemanticArchiveLibrary` registration metadata

The logical source identity is intentionally independent of OCR, extraction policy, semantic model and batch size. Changing those settings may invalidate a representation, but must not create a second logical document for the same physical source.

---

## 5. Source roots

Each source binding contains:

- `BindingId`
- `SemanticArchiveRootPath`
- `SemanticArchiveRecursive`
- `SemanticArchiveExclusions`
- `SemanticArchiveFileTypeFilterEnabled`
- `SemanticArchiveSupportedExtensions`
- `SemanticArchiveArtifactPlacementMode`
- `SemanticArchiveSharedArtifactRoot`
- `SemanticArchiveShadowArtifactRoot`
- `SemanticArchiveEnableOcr`
- `SemanticArchiveOcrBatchPages`
- `SemanticArchiveExtractionOptionsSignature`
- `SemanticArchiveScopeTags`

### 5.1 Default file-type filter

The filter is enabled by default.

Default extensions are:

```text
.pdf
.doc
.docx
.docm
.rtf
.xlsx
.xlsm
.pptx
.pptm
.txt
.csv
.eml
.msg
.png
.jpg
.jpeg
.gif
.bmp
.tif
.tiff
.webp
.svg
```

A custom list replaces the defaults.

Turning the filter off does not magically add converters; it only allows any format supported by the existing text-export pipeline to be considered.

### 5.2 Always-excluded generated content

Knowledge Store `.redink` content and Semantic Archive generated output are not re-ingested. This prevents feedback loops where generated KB/SA artifacts become sources of another archive build.

### 5.3 Recursive traversal and path safety

Source traversal rejects unsafe paths, including:

- device paths;
- alternate data streams;
- traversal through reparse points where disallowed;
- paths outside the registered source root;
- paths exceeding supported Windows source/generated path budgets.

UNC source paths are supported subject to the same containment, permission and path rules.

---

## 6. Extraction pipeline

Semantic Archive ingestion reuses the existing Red Ink `text_export_to_text` implementation. SA does not maintain a separate PDF/Office conversion engine.

The high-level sequence is:

1. enumerate candidate sources;
2. filter by archive/source rules;
3. detect source changes;
4. validate or calculate source fingerprint/hash;
5. reuse a compatible existing extraction where possible;
6. otherwise extract through the existing exporter;
7. validate extraction coverage;
8. optionally create a per-document semantic section index;
9. create semantic metadata/card;
10. update archive routing structures;
11. validate and atomically publish a new generation.

### 6.1 Source fingerprint

A source fingerprint includes:

- length;
- last-write UTC ticks;
- SHA-256.

Length/time can be scan hints; validated reuse depends on stronger identity/version checks.

### 6.2 Representation contract

An extracted representation stores:

- `RepresentationId`
- original `SourceHash`
- `ExtractorVersion`
- `OptionsSignature`
- `TextPath`
- `TextFileHash`
- `TextByteLength`
- `EncodingName`
- `Completeness`
- `SourceMapJson`
- `ExtractedUtc`

Completeness is one of:

```text
complete
empty
incomplete
unknown
```

`unknown` is deliberately not treated as `complete`.

### 6.3 UTF-8 hash semantics

The exported text file and semantic index payload have distinct hash meanings.

Do not assume:

```text
exported text-file hash == indexed payload hash
```

The text exporter may write UTF-8 with a BOM, while semantic index payloads use their own byte semantics. Implementations must preserve the distinct source hash, extracted-file hash and indexed payload hash.

---

## 7. PDF OCR behavior

OCR is selective.

Red Ink first checks the native PDF text layer. OCR candidates are selected from pages that require it rather than sending every page to OCR by default.

Typical strong OCR signals include:

- no usable text together with images;
- very sparse text on image-heavy pages;
- text-quality/encoding problems.

A single sparse page is treated as a heuristic observation, not automatic proof that content is missing.

### 7.1 OCR batching

`SemanticArchiveOcrBatchPages` controls the maximum candidate pages grouped into one OCR model call.

- default: `16`
- accepted range: `1..75`

Only candidate pages are batched. Native-text pages are retained and are not needlessly OCRed.

A successful batch must still prove which requested pages were returned. Coverage is evaluated per requested page even when model work is batched.

### 7.2 Completeness after OCR

A PDF can be `complete` when:

- the native text layer was inspected over the full known page range;
- every OCR-required page was included in a successful verified OCR result;
- no required page is missing;
- there was no truncation/cancellation/other unresolved coverage gap.

If required pages fail, the extraction is `incomplete`.

If the system cannot prove the coverage, it remains `unknown`.

### 7.3 Partial-search policy

`SemanticArchiveAllowPartialSearch` defaults to `False`.

Therefore incomplete/unknown representations are normally excluded from search. This is a quality boundary, not a processing failure.

---

## 8. Per-document section indexing

`SemanticArchiveSourceIndexThresholdBytes` uses the **original source-file size**.

Default:

```text
65536 bytes
```

`0` disables the additional per-document section index.

For large documents above the threshold, Red Ink can build a semantic section index. When such an index exists, content retrieval uses semantic section selection and reads only the selected chunks. Literal hits do not bypass that section-selection contract.

Small documents without a section index can be read directly within bounded byte limits.

---

## 9. Semantic cards

A document card contains:

- `CardId`
- generic level: `CONTAINER`, `DOCUMENT`, or `SECTION`
- `TargetId`
- `PartitionKey`
- `SourceVersion`
- `RepresentationId`
- `Title`
- `Summary`
- `Topics`
- `UserIntents`
- `Identifiers`
- `ExactTerms`
- richer metadata object
- `RetrievalText`
- routing-reduction information
- optional byte range
- permission-neutral flag

Cards are navigation/ranking metadata. A summary is not itself authoritative evidence for answering a user question unless the retrieval contract explicitly says so.

Exact answers are grounded by the read operation against the extracted text.

---

## 10. Immutable generations

Published archive state is generation-based.

A generation manifest contains, among other fields:

- archive ID;
- generation ID;
- previous generation ID;
- configuration signature;
- creation time;
- writer fence token;
- root node ID;
- node descriptors;
- document-shard descriptors;
- optional term shards;
- inventory/counts;
- validation status;
- diagnostics.

`current.json` atomically points to the active validated generation.

Repeated `refresh`, `reindex` and `extract` must not multiply logical documents. Retired/tombstoned records are not supposed to remain as active records in a newly activated generation.

---

## 11. Writer locking and publication safety

Mutating archive operations use a storage-backed writer lock and fencing counter.

This protects against concurrent Word/Outlook/worker writes to the same archive state.

Publication is atomic:

1. build work is written into new immutable artifacts;
2. validation succeeds;
3. a generation pointer is updated atomically;
4. the prior validated generation remains safe if the new build is cancelled or fails.

This is why a graceful worker cancellation is recoverable.

---

## 12. Access-control model

### 12.1 Source files remain authoritative

The primary rule is:

> A derivative does not grant access to an original.

Before search results disclose document metadata/path information, and again before evidence is read, Red Ink rechecks the requesting principal's current access to the source.

A local directly interacting user uses the current Windows identity. Remote, delegated and unattended scenarios require an independently verified requesting-principal policy.

### 12.2 Shared derivatives

Shared derivative artifacts may be writable even if the original is read-only.

However, derivative ACLs must not broaden source readability. Shared-artifact publication includes integrity and security checks; unsafe shared storage falls back to a personal derivative area rather than weakening source protection.

### 12.3 Private catalog

The personal catalog and its generated private state are protected as a per-user access domain, with SYSTEM/Administrators handling according to the implementation's private-storage policy.

### 12.4 Central library descriptors

Central library descriptors control who can discover/subscribe to an archive definition. They do **not** grant access to source documents.

A subscriber must pass both:

1. central descriptor/library authority checks;
2. current original-file access checks.

A cached private index cannot authorize access when the central publication or original-file permission is no longer valid.

---

## 13. Central-library publisher/subscriber model

### 13.1 Publishing

A user can create an archive locally, test it, and then publish the definition.

The central descriptor contains archive configuration, not source content.

For portable publication, shared source roots should be meaningful to other users. Publishing local/mapped paths into a UNC library is rejected where the path cannot represent the same source for other users.

### 13.2 Updating

Publishing the same archive again updates the same library identity and increments the library revision when the definition changed.

New documents under an existing source root do **not** require republishing the definition. They require a normal archive refresh.

Changes to source-root definitions or archive processing/search settings do require republishing the definition.

### 13.3 Withdrawal

Withdrawal publishes a withdrawn revision.

Subscribers stop using the archive after reconciliation, while:

- original files remain untouched;
- the publisher's local archive remains;
- the stable publication identity is retained for possible republishing.

### 13.4 Automatic subscription

Readable library entries are synchronized into the user's local catalog.

A subscriber may opt out. The opt-out persists locally and is not undone by the next library sync.

Temporary library unavailability is not interpreted as withdrawal.

### 13.5 Synchronization interval

The normal library synchronization interval is currently 300 seconds.

The standalone worker also synchronizes the library before resolving the requested archive.

---

## 14. Retrieval scope and archive selection

The host resolves an authorized archive scope before exposing retrieval to the model.

For direct interactive Word/Outlook use, the intended order is:

1. explicit session/archive selection;
2. configured defaults;
3. enabled locally visible archives when the host explicitly allows the interactive fallback.

An explicit empty selection remains an opt-out and must not silently broaden.

Remote/delegated/AutoPilot scenarios do not gain interactive fallback authority from model tool arguments.

A model call can narrow an already authorized scope; it cannot broaden it.

---

## 15. Freestyle trigger syntax

Supported SA controls include:

```text
(sa)
(sa: query terms)
(sa: archive:"Archive Name")
(sa: archive:"Archive Name" query terms)
(sa: archive:"Archive Name" mode:content query terms)
(sa: archive:"Archive Name" mode:files query terms)
```

Modes:

- `content` — search and then read exact evidence;
- `files` — return relevant original-file references.

`(sa)` uses the surrounding user task as the query.

The parser reads only the authoritative current user request. Retrieved document content is never treated as control syntax.

Word and Outlook Freestyle also expose a Sources menu. Selecting a source replaces the existing scope-only SA/KB control rather than appending duplicate triggers.

Freestyle displays visible feedback such as:

```text
Querying Semantic Archive...
```

while the SA request is being resolved.

---

## 16. Search behavior

The public search contract returns:

- status/message;
- query;
- hits;
- optional continuation reference;
- detailed coverage information.

A hit contains:

- opaque `HitReference`;
- archive ID;
- generation ID;
- document ID;
- source path and authorized `SourceUri`;
- display name;
- title/summary;
- relevance/reason/channel;
- representation ID;
- source hash;
- extraction completeness;
- optional literal-match offset.

The `HitReference` is run-scoped and opaque. Consumers must not manufacture or persist arbitrary references as durable document IDs.

### 16.1 Search strategies

Current strategies include:

- exact metadata;
- bounded flat document-card selection for small archives;
- hierarchical semantic routing for larger/continued searches;
- optional literal-text inspection under a separate bounded budget.

For small complete inventories up to the configured threshold (currently 64 source records), the service can rank document cards directly.

### 16.2 Continuations

A partial search can return a continuation reference.

A continuation is bound to:

- the run;
- original query;
- archive scope;
- literal-search parameters when applicable.

Changing the query or scope requires a new search.

A partial result is not a tool failure. It means additional coverage may be available.

### 16.3 No-match semantics

“No hit in explored material” is not automatically an exhaustive negative statement. Coverage diagnostics distinguish explored records from excluded/unavailable/incomplete sources.

---

## 17. Exact evidence read

`semantic_archive_read` accepts opaque hit references returned by search.

Before returning evidence it:

- validates the run binding;
- validates the pinned generation/document;
- rechecks current source access;
- verifies source/extract/index integrity;
- uses the per-document section index when required;
- returns exact extracted text and provenance.

Default read budget is bounded. The service rejects excessive byte/excerpt counts.

The read result includes source identity, representation identity, hashes, byte offsets, source-map metadata, and exact extracted text.

---

## 18. Retrieval budgets

Default archive retrieval budgets currently include:

```text
MaxNodesVisited              40
MaxModelCalls                24
MaxElapsedSeconds            120
MaxCandidateFiles            120
MaxEvidenceBytes             65536
InitialBranches              4
MaxExactLookupDocuments      2000
MaxSectionCandidates         48
MaxPromptCharacters          32000
MaxRequestTokens             65536
MaxLiteralScanBytes          16777216
MaxLiteralScanDocuments      64
```

Hosts apply hard caps in addition to archive settings.

For multi-archive selection, the effective budget is conservatively bounded.

---

## 19. Source links and unattended delivery

Interactive results can expose an authorized reference to the original source. Red Ink constructs the source URI from the validated original path; it must not point to private extracted text or `.redink-sa`.

In the embedded Outlook UI, `file://` navigation is intercepted and opened through the existing host path-opening handler rather than relying on the browser's direct local-file navigation.

AutoPilot/unattended delivery must not assume the recipient can use the interactive user's local or UNC path. Where supported, original documents actually read as evidence are delivered as artifacts/attachments after current access checks.

---

## 20. Automatic processing

Semantic Archive automatic maintenance is hosted by **Outlook only**.

Word continues to support:

- manual SA administration;
- manual refresh/reindex/extract actions;
- interactive search/read;
- its independent Knowledge Store maintenance.

Outlook may host:

- central-library synchronization;
- content indexing;
- permission maintenance.

These run through the generic background-maintenance coordinator and are subject to idle/activity/cancellation rules and maintenance windows.

This is not a Windows service. Outlook must be running for Office-hosted automatic processing.

---

## 21. Standalone worker

### 21.1 Quick start

First build or normal update:

```powershell
redink-sa-worker.exe refresh "Archive Name"
```

The worker resolves Red Ink's active configuration source automatically unless `--ini` is supplied.

A unique archive display name or archive ID can be used.

### 21.2 Operations

```text
refresh      discover changes, reuse valid work, extract/index what is needed, publish
retry        retry failures/deferred jobs that are ready
reindex      rebuild semantic metadata/indexes while reusing valid extracted text
extract      force re-extraction/OCR and then rebuild semantic metadata
permissions  reconcile derivative permissions without content indexing
```

Use `refresh` for first-time indexing.

### 21.3 Worker defaults

```text
Processing files per batch        64
Discovery entries per pass        5000
Discovery time per pass           60 seconds
Continuation delay                1 second
```

Example throttling:

```powershell
redink-sa-worker.exe refresh "Archive Name" --batch-files 8
```

### 21.4 Progress

Interactive consoles display an in-place progress line on stderr.

stdout remains JSON-lines for automation.

### 21.5 Cancellation

Ctrl+C requests graceful cancellation. Durable checkpoints and the last validated generation remain usable.

Avoid hard process termination when Ctrl+C is available.

### 21.6 Exit codes

```text
0    complete
1    configuration/auth/runtime error
2    invalid arguments
3    incomplete/deferred/failed work remains
130  graceful cancellation
```

### 21.7 Office-host limitation

The worker does not automate Word/Outlook UI. A source requiring an Office-host-only reader can remain host-required.

---

## 22. Administration console

The Semantic Archives administration UI includes areas for:

- Archive
- Library
- Source roots
- Retrieval budgets
- Documents and selected actions
- Status and diagnostics
- Personal automatic processing

Typical administrator workflow:

1. create archive;
2. set name/description;
3. add source root(s);
4. configure filter/OCR/derivative roots;
5. save;
6. `Refresh archive`;
7. review searchable counts;
8. optionally publish the archive definition to the central library.

Changing selected archive refreshes the displayed status.

Unsaved edits are not silently discarded.

### 22.1 Diagnostics

Normal diagnostics are intentionally user-oriented.

Technical details such as:

- stable IDs;
- source paths;
- routing splits;
- permission internals;
- processing signatures;
- shared-publication events;
- queue state;
- generation identifiers

belong behind **Show technical details**.

---

## 23. Common operational scenarios

### Create a private archive

1. Configure `SemanticArchiveCatalogPathLocal`.
2. Open Semantic Archives administration.
3. Create archive.
4. Add source root.
5. Save.
6. Refresh.

No library path is required.

### Publish a central archive

1. Configure both local and library paths.
2. Create/test archive locally.
3. Ensure source roots are portable/shared as required.
4. Use **Publish / Update library**.

### Subscribe automatically

Normal users need no manual import. If the library path is configured and they can read a descriptor, it is synchronized into their local catalog unless they opted out.

### Add files to a published archive

Add files under an existing registered source root and run/allow refresh. Republishing is not required unless the definition itself changed.

### Rebuild semantic metadata

Use `reindex`.

### Force new extraction/OCR

Use `extract`.

### Recover from interrupted processing

Run `refresh` again. Durable discovery/extraction checkpoints and immutable published generations are designed for continuation.

---

## 24. Troubleshooting

### SA does not appear

Check `SemanticArchiveCatalogPathLocal`. If blank, SA is deliberately disabled.

### Central archive does not appear

Check:

- local path is configured;
- library path is configured;
- descriptor can be read under the current Windows identity;
- publication is not withdrawn;
- local opt-out is not set.

### Archive shows fewer jobs than files

Discovery is bounded and checkpointed. If discovery has a continuation, the current job count is not the final physical file count.

### Document is excluded as incomplete/unknown

Inspect extraction diagnostics. If scanned pages are involved, enable OCR and refresh/extract as appropriate. Do not simply enable partial search unless accepting incomplete evidence is an intentional policy choice.

### Shared artifact reuse is unavailable

Check the shared-artifact path and ACL/security diagnostics. Private fallback can continue without weakening source permissions.

### Search is slow

Inspect:

- `Coverage.RetrievalStrategy`
- `Coverage.ModelCalls`
- `Coverage.ElapsedMilliseconds`
- archive size;
- retrieval budgets;
- whether a continuation/hierarchical route was needed.

### Model metadata parser failure

The current metadata adapter tolerates some common shape deviations (for example `Title` or `Summary` returned as a string array) before strict validation.

---

## 25. Backup and cleanup

Back up:

- `redink-sa-catalog.json`;
- local `sa-archives` if preserving processing state matters;
- central library descriptors if used;
- original sources independently.

Do not treat derivative storage as the only copy of source content.

Generated Semantic Archive directories may be rebuilt, but deleting them discards private generations/checkpoints and can cause expensive extraction/OCR/model work to repeat.

---

## 26. Current validation status and operational caveat

The implementation has been extensively reviewed and exercised during development, including real user runs for search, OCR and worker execution.

However, this documentation must not be read as a claim that every combination has completed a clean Windows/Visual Studio/VSTO/UNC/multi-user acceptance matrix.

Important remaining production validation areas include:

- clean VS2022 builds of all hosts and worker after the latest changes;
- multi-user shared-artifact reuse;
- live ACL removal/restoration;
- target UNC/SMB environments;
- automatic library subscription with multiple Windows users;
- AutoPilot original-document delivery;
- representative mixed/native/scanned OCR batches;
- final performance measurement after search-strategy changes.

