# Semantic Archives — User and Administrator Guide

**Updated:** 5 October 2026  
**Audience:** Red Ink users and administrators  
**Purpose:** setup, day-to-day use, maintenance, sharing, and troubleshooting  
**This guide intentionally omits programming and internal implementation details.**

---

## 1. What a Semantic Archive is

A Semantic Archive lets Red Ink search a managed collection of files by meaning rather than only by exact words.

An archive can contain one or more folders. Red Ink prepares the files so that users can find relevant documents and then read exact supporting content from the original material.

Examples include:

- compliance collections;
- legal know-how libraries;
- data-protection guidance;
- project documentation;
- policies and procedures;
- collections on shared drives.

Access to the original files remains decisive. Creating or sharing an archive does not grant users access to files they could not otherwise read.

---

## 2. Prerequisites

Semantic Archives require a **personal catalog location** for the current Windows user.

If no personal catalog location is configured, Semantic Archives are disabled.

The catalog is separate from your source folders. Do not place your working documents inside the catalog merely because they belong to an archive.

A central library is optional. It is used to share archive definitions with other users; it is not required for private archives.

---

## 3. Opening the Semantic Archives administration window

Open the **Semantic Archives** administration window from Red Ink in the host application where the command is available.

The window lets you:

- choose the personal catalog location;
- create and unregister archives;
- add and remove source folders;
- configure OCR and source filtering;
- control search visibility and default selection;
- refresh and repair archives;
- rebuild semantic indexes;
- manage generated-file permissions;
- inspect archive status and diagnostics;
- publish or synchronize archive definitions through a central library.

---

## 4. Choosing the personal catalog location

At the top of the administration window, use one of these options:

- **Enter catalog path…**
- **Browse…**
- **Use recommended location…**

Choose a private writable location for the current Windows user.

The recommended location is normally the safest default.

Changing the catalog location does not automatically move or delete an existing catalog. If changing locations, verify the intended catalog before continuing.

---

## 5. Creating an archive

1. Click **New archive**.
2. Enter an archive name.
3. Optionally enter a description.
4. Decide whether the archive should be enabled for search.
5. Decide whether it should be used by default.
6. Add at least one source folder.
7. Save the changes.
8. Run **Refresh archive** to prepare the files.

Creating the archive entry does not create or move the source folders.

---

## 6. Adding source folders

In the archive's source area you can type a folder path or choose one with **Browse directory…**.

Both local Windows paths and accessible network paths can be used.

For each source folder, configure the relevant options.

### Include subfolders

Enable **Include subfolders** when the complete directory tree should be part of the archive.

### Allow OCR

Enable **Allow OCR** when PDFs or images may need optical character recognition.

OCR is particularly useful for:

- scanned PDFs;
- image-only pages;
- PDFs where some pages contain images but little usable text;
- documents whose native PDF text layer cannot be read reliably.

OCR can increase processing time and model usage.

### File type filtering

Use **Filter source file types** to limit the archive to configured supported file types.

The default list covers common Office, PDF, text, email, and image formats supported by the Red Ink text converter.

Use **Restore office/image defaults** if you want to reset the extension list.

### Exclusions

Use exclusions to skip files or folders that should not be indexed.

Examples include:

```text
Temp
*.tmp
Archive\Old\*
```

Excluded directories are skipped rather than traversed file by file.

---

## 7. Generated-file storage

Each source can use automatic or private generated-file storage.

### Automatic

Use automatic placement when reusable generated document artifacts should be shared where Red Ink can safely preserve the original source's reader access.

If safe shared placement is not possible, Red Ink can fall back to private storage.

### Private

Use private placement when generated document artifacts should remain only in the current user's private storage.

The original files are not made writable and their permissions are not changed by Semantic Archive indexing.

---

## 8. Search visibility and default use

### Enable archive

An enabled archive can be offered for search.

Disable an archive if it should remain configured but should temporarily not be available for normal search.

### Use by default

Enable **Use by default** when this archive should be included in the user's normal default Semantic Archive scope.

Explicit archive selections still take precedence when the search interface allows them.

---

## 9. Incomplete extracted text

The option **Allow incomplete extracted text** controls whether Red Ink may search documents whose complete extraction could not be verified.

For most production archives, leave this **off**.

When it is off:

- complete documents are searchable;
- incomplete or unknown extraction is excluded from search;
- the Admin diagnostics show the affected documents so they can be repaired.

Turning this option on is a policy choice, not a repair. It does not improve extraction quality.

---

## 10. The main maintenance commands

The most important Admin commands have deliberately different meanings.

### Refresh archive

Use **Refresh archive** for normal ongoing maintenance.

It:

- discovers new files;
- detects changed and removed files;
- reuses valid existing work;
- processes only what is needed;
- publishes an updated searchable archive.

Use this after adding files below an already configured source folder. You do not need to republish the central library definition simply because new files appeared under the same root.

### Retry failed

Use **Retry failed** when some documents failed, were deferred, or were excluded because their extraction was missing, empty, incomplete, or unknown.

Current intended behavior:

- a missing/empty/incomplete/unknown extraction is performed again, including OCR where enabled;
- a verified complete extraction is reused if only later semantic/index processing failed.

This is usually the right first action after reviewing failed or coverage-excluded documents.

### Rebuild semantic index

Use **Rebuild semantic index** when extracted text is already valid but the semantic navigation/index should be rebuilt.

This command does **not** run extraction or OCR.

If a document requires new extraction, the operation reports that requirement instead of silently starting OCR.

### Re-extract/OCR all

Use **Re-extract/OCR all** when you intentionally want to recreate extracted text for the entire archive.

This can be expensive. Use it when extraction quality itself must be regenerated, not for normal updates.

OCR is used only for source folders where OCR is enabled.

### Permissions: all

Use **Permissions: all** to reconcile access on generated files with the current original-source access situation.

This operation does not change the permissions of the original documents.

---

## 11. Working with selected documents

The **Documents** area lists filenames, readable status, a recommended action and the source folder. The window starts larger and can be resized to make more room for the list.

To find documents that need attention:

1. Choose **Needs attention** (the initial filter), or a specific status such as **Incomplete / unknown extraction**, **No readable text**, **Processing failed** or **Needs extraction / indexing**.
2. Optionally enter part of the filename or folder path.
3. Click **Find documents**. Only matching published records are displayed; this does not rescan the original source folders or start maintenance.
4. Click a column header to sort the loaded matches. Click again to reverse the order.
5. Click **Load more matches** to append more results while keeping existing selections.
6. Select the files to repair and run the appropriate selected command.

**Source access unavailable** is a separate view. Unavailable sources also appear in **All documents**, with their cached names, paths, processing status and diagnostics hidden. Issue filters do not reveal the hidden status of inaccessible files.

Other views include **Searchable** and **Removed**. An incomplete document can also be searchable if incomplete-text search was deliberately enabled.

Each retrieval loads up to 200 matching rows. Sorting applies to the rows already loaded, not all matching documents in the archive. The display holds at most 5,000 rows; refine the filters if that limit is reached. Filtering searches the published metadata in bounded background steps and may take time for a large archive with few matches.

Use **Select loaded matches** or select individual rows, then use:

- **Rebuild semantic index (selected)**
- **Re-extract/OCR selected**
- **Permissions: selected**
- **Retry selected**

At most 1,024 document IDs can be selected for one operation. An empty selected-document list never means "all documents". Changing a filter clears the previous results and selection. **Find documents** starts a fresh published-status snapshot; use it again after maintenance to see the new state.

Stable document IDs and raw diagnostics are available under **Show IDs / technical details** when troubleshooting. Normal selection does not require copying cryptic IDs.

---

## 12. Recommended repair sequence

When an archive shows a mixture of incomplete, unknown, failed, or deferred documents:

1. verify that OCR is enabled for source folders that need it;
2. run **Retry failed**;
3. refresh the status;
4. review any documents that remain incomplete or unknown;
5. use **Re-extract/OCR selected** for specific problem files if needed;
6. use **Rebuild semantic index** only after extraction is valid when the semantic index itself needs rebuilding.

Do not use **Allow incomplete extracted text** merely to hide a processing problem.

---

## 13. Status and diagnostics

Use **Refresh status** to inspect the current published archive and outstanding work without starting indexing.

The normal diagnostics focus on:

- source count;
- searchable count;
- complete/incomplete/unknown extraction;
- failed or deferred work;
- items requiring extraction;
- whether further discovery or permission work remains.

Use **Show technical details** only when troubleshooting with a developer or support person.

Use **Copy diagnostics** to copy the displayed diagnostic information.

---

## 14. Pausing work safely

Use **Pause** to request that current content/permission processing stop at a safe checkpoint.

Published search remains available while maintenance is paused.

Closing the Admin window while work is active also requests a safe stop rather than intentionally corrupting a running operation.

---

## 15. Automatic background processing

Automatic Semantic Archive maintenance is hosted by **Outlook**.

Word can still run manual Semantic Archive operations and searches, but it does not host automatic Semantic Archive maintenance.

Personal automatic settings include:

- **Enable automatic content indexing**;
- a content-indexing time window;
- **Enable automatic permission maintenance**;
- a permission-maintenance time window.

Each archive can also independently allow automatic updates and define its own processing window.

Outlook must be running and idle for Outlook-hosted background processing to occur.

---

## 16. Processing windows

A blank processing window means processing may occur at any local time when the other conditions allow it.

Examples of supported user-facing time rules include:

```text
allow:22:00-06:00
```

or

```text
deny:08:00-18:00
```

Multiple ranges can be separated with semicolons.

Use overnight windows if OCR or indexing could interfere with daytime work.

---

## 17. Using the standalone worker

The standalone executable is:

```text
redink-sa-worker.exe
```

It uses the same normal Red Ink configuration as the Office add-ins unless an explicit configuration source is supplied.

### Normal update

```powershell
.\redink-sa-worker.exe refresh "Archive Name"
```

### Retry failed/incomplete work

```powershell
.\redink-sa-worker.exe retry "Archive Name"
```

### Repair extraction and rebuild the current semantic index

For an archive already using the current index format, use `repair` when you want one unattended operation that repairs problematic extraction/OCR where necessary and then rebuilds the current semantic index.

For one archive:

```powershell
.\redink-sa-worker.exe repair "Archive Name"
```

For multiple archives in one process:

```powershell
.\redink-sa-worker.exe repair "VISCHER Compliance" "VISCHER DP Know-how"
```

This is the recommended command for an overnight maintenance run when the archives contain missing, failed, incomplete, or unknown extracted content and should finish with a rebuilt current semantic index.

### Rebuild semantic index without extraction/OCR

```powershell
.\redink-sa-worker.exe reindex "Archive Name"
```

### Force extraction/OCR again

```powershell
.\redink-sa-worker.exe extract "Archive Name"
```

### Permission maintenance

```powershell
.\redink-sa-worker.exe permissions "Archive Name"
```

---

## 18. Worker progress output

Worker processing messages are written as normal lines underneath one another, for example:

```text
Processing source — policy.pdf | 18 completed; 112 queued
Extracting text — policy.pdf
Building semantic metadata — policy.pdf
```

The worker no longer depends on repeatedly overwriting one console line. Resizing the PowerShell/console window may cause normal text reflow, but it should not cause the worker to erase or overwrite earlier progress lines.

Press **Ctrl+C** for graceful cancellation. To resume the **same operation**, copy the `operationId` from its `started` event and add `--operation-id "<operation-id>"`, retaining exactly the same operation, archives and selected-document scope. Without that ID, the command starts a new operation; reusable completed work can still be retained. Never reuse an operation ID for a different scope.

Avoid forcibly terminating the process when Ctrl+C is available.

---

## 19. Worker result codes

The worker uses these result codes:

- **0** — operation completed;
- **1** — configuration, authentication, or runtime error;
- **2** — invalid command-line arguments;
- **3** — work is still incomplete, deferred, failed, or requires further attention;
- **130** — graceful cancellation was requested.

For unattended overnight processing, a final code `3` is important: the process ran, but not everything became ready.

A batch showing `published=true` and `failed=0` is not sufficient proof that every selected document became searchable. Check the final `published_scope_verified` event: `complete=true`, `unresolved=0` and `missingSelectedDocuments=0`, together with the published inventory. Completion applies to the command's requested scope. Outstanding permission discovery or artifact repair is reported separately from content searchability.

---

## 20. OCR troubleshooting

If a PDF remains incomplete:

1. verify **Allow OCR** is enabled for that source;
2. use **Retry failed** or **Retry selected**;
3. if necessary, use **Re-extract/OCR selected** for that document;
4. review the diagnostics again.

A document can remain incomplete if Red Ink cannot verify every page that requires OCR. This is preferable to falsely marking a partially read document as complete.

Sparse text on a page is a hint that OCR may be required; it is not by itself proof that the page was not extracted.

If the diagnostics identify a missing assembly during PDF preparation (`pdf_chunk_creation`), OCR has not yet reached its model request. Install the complete current worker output, including its configuration file and dependency files, before retrying. Installing a .NET 8 runtime is not the remedy for this .NET Framework worker. Provide the missing assembly name to support rather than copying arbitrary DLL versions into the output folder.

If a document still fails after a fresh OCR attempt, provide the document-specific diagnostics to support/development rather than globally enabling incomplete-text search.

---

## 21. Office document troubleshooting

If a Word document is shown as `unknown` or `incomplete`, use **Retry failed** after updating to the current version.

A current retry performs a fresh extraction when the stored extraction is not verified as complete.

If the file contains damaged or unreadable document parts, it may correctly remain incomplete rather than being marked complete based only on the main body text.

---

## 22. Central library setup

A central Semantic Archive library is optional and is configured separately from the personal catalog.

The central library is used to distribute archive definitions to other users.

It does **not** copy the source documents into the library.

For a personal archive that is ready to share:

1. configure the central library location;
2. create and test the archive locally;
3. make sure shared/network source folders use paths that other intended users can reach;
4. select the archive;
5. choose **Publish / Update library**.

The archive name, description, and source definition become available to permitted library users. Their access to the source files is still governed by the source locations themselves.

---

## 23. Updating a published archive

Use **Publish / Update library** after changing the shared archive definition, such as:

- source folders;
- archive settings;
- description or other published configuration.

You do **not** need to republish merely because new files were added below an already published source folder. Use **Refresh archive** for content changes.

---

## 24. Synchronizing the central library

Use **Sync library** to immediately read the configured central library and reconcile local subscriptions.

This does not itself run OCR.

New or changed source content is prepared through normal archive maintenance or the worker.

---

## 25. Withdrawing a published archive

Use **Withdraw from library** to publish a withdrawn revision.

Withdrawal:

- stops the shared definition from remaining normally usable by subscribers after synchronization;
- does not delete your local archive;
- does not delete original files;
- does not broaden or change original-file permissions.

---

## 26. Subscriptions and local opt-out

Readable central library entries can be subscribed automatically.

For a subscribed archive, publisher-managed archive-definition fields are not edited locally.

Unregistering/unsubscribing a subscription acts as a local opt-out. It does not delete the publisher's library entry or the original documents.

A user can later re-enable the subscription through library synchronization/availability.

---

## 27. Permissions and shared artifacts

Semantic Archives may reuse generated per-document work between authorized users when configured for shared artifacts.

Important user-facing rules:

- the original file remains authoritative;
- Red Ink does not grant new source access merely because a derivative exists;
- original source permissions are checked again during retrieval;
- if safe shared generated-file storage is not available, Red Ink can use private storage instead;
- permission maintenance never rewrites the original file's permissions.

If access to an original file is removed, users should not expect a previously generated artifact to remain a way around that restriction.

---

## 28. Unregistering an archive

**Unregister** removes the archive registration from the personal catalog.

It does not delete the original source files.

Generated artifacts are intentionally retained rather than being automatically destroyed as part of unregistering.

If the selected archive is a locally published archive with an available library definition, **Unregister** asks you to confirm both withdrawal from the central library and removal from the personal catalog. Withdrawal runs first. If withdrawal fails, the local registration remains; resolve the reported library or publisher-permission problem and retry. A separately withdrawn definition stays withdrawn even if subsequent local removal is interrupted.

For subscribed archives the button is shown as **Unsubscribe**. This remains a local opt-out and does not withdraw the publisher's definition.

An old archive can report `unsupported_semantic_index` because it has no current routing graph. The old index cannot be migrated by **Retry**, **repair** or **Rebuild semantic index**. The Admin window shows this compatibility problem and disables incompatible content actions. Library withdrawal and unregistering remain available independently of the old index. If the collection is still needed, create and configure a new archive using the source folders, then prepare a new current-format index.

---

## 29. Recommended administrator operating practices

For production use:

- keep **Allow incomplete extracted text** off unless there is a deliberate policy reason to enable it;
- enable OCR only on sources that may genuinely need it;
- use **Refresh archive** for routine updates;
- use **Retry failed** for failed/coverage-excluded items;
- use **repair** in the worker for unattended comprehensive recovery plus semantic-index rebuild on archives already using the current format;
- use **Rebuild semantic index** when extraction is already good;
- avoid **Re-extract/OCR all** unless a true extraction reset is needed;
- schedule heavy OCR/index maintenance outside working hours;
- monitor worker exit code `3` and Admin diagnostics rather than assuming a completed process means every document became searchable;
- test source access using the same Windows accounts that will use the archive;
- use UNC paths for shared source folders when central library users must reach the same network content.

---

## 30. Recommended nightly recovery command for the two VISCHER archives

For the two archives discussed during implementation, provided both already use the current index format, use:

```powershell
.\redink-sa-worker.exe repair "VISCHER Compliance" "VISCHER DP Know-how"
```

This is preferable to manually chaining `retry` and `reindex`, because `repair` is intended to repair missing/failed/incomplete/unknown extraction where required, reuse already verified complete extracts, and finish by rebuilding the current semantic index for the selected archives.

After the run, check the worker exit code and refresh the Admin status for both archives.


---

## 31. Searching and understanding coverage

Ask Red Ink to use the Semantic Archives and select the relevant archive scope when the interface offers that choice. Examples:

- "Which documents do we have on professional secrecy? Give me a list with source links."
- "Read the relevant documents and identify the passages supporting this conclusion."

A document list uses indexed names, titles and summaries to find candidate files. It is not the same as reading all original text. For substantive conclusions or quotations, ask for supporting passages from the relevant documents. Summaries help discovery but are not themselves exact source evidence.

Search may return a **partial** result because each request has bounded processing and model budgets. Red Ink can continue the same search using the returned continuation. If it stops while continuation remains, treat the answer as a partial document list, not an exhaustive inventory. A missing match does not prove that the collection contains no relevant document.

By default, Red Ink continues a unique pending search for the same archive selection rather than restarting it merely because the model rephrases the query. The original question remains active; diagnostics identify a deferred formulation that was not searched. A deliberately different question can explicitly start a new search, which may repeat work and previously returned files. When several searches are pending, the exact continuation identifies the one to advance. Continue an existing search before requesting alternative formulations merely to obtain more matches. Requesting a larger useful document list can reduce repeated small result pages, within the search limits.

An unavailable or obsolete published archive requires maintenance or a new current-format archive. Rephrasing the question does not repair its publication. A fully searchable published inventory confirms that the documents are ready for retrieval; it does not mean every document was examined for a particular question.

Search time includes model selection, current source-access checks, continuation and answer generation. OCR/indexing time is a separate maintenance cost. For a performance report, provide the complete tooling log from the same question and archive scope. The log can contain source names, paths and text; share it only with the intended support recipients.
