# Tasks

## Current implementation basis

- Baseline: `RI_Gen2_260926_2355.zip`.
- This file is cumulative for the current implementation stream.
- Product invariant: no integrated/runtime self-tests are to be added. Regression checks stay external/manual or in separate tooling and must not run during Red Ink startup.
- Architecture invariant: shared behavior remains model-, provider-, tool-, skill-, organization-, template-, document- and language-agnostic. Outlook and Word use the same shared implementation wherever the concern is common.

## Completed in package 01

- [x] Remove the dedicated Agent self-test sources from the maintained code basis.
- [x] Remove `TextFileSnapshotSelfTests.vb` from `SharedLibrary.vbproj`.
- [x] Remove the now-orphaned `ClearForTests` and `ResetForTests` hooks used only by the deleted self-tests.
- [x] Confirm that no Agent self-test invocation is present in the supplied Outlook/Word startup paths before removal.
- [x] Correct `ToolResultStore.GetWindow` so EOF, beyond-EOF and empty stored content return an empty window instead of attempting a one-character substring.
- [x] Preserve the existing direct-store compatibility rule that non-positive `maxChars` still means one character when content remains.
- [x] Correct `ContextExpandTool.Execute` to derive the returned window and all navigation metadata from one captured stored body.
- [x] Normalize negative/beyond-EOF starts in the returned metadata and avoid `startChar + window.Length` overflow on extreme input.
- [x] Preserve the existing public `context_expand` schema, default, 500..100000 clamp, unknown-reference error, tool name, result reference, JSON field names and types.
- [x] Do not change compaction retention, replay suppression, payload budgets, model/provider selection, retry behavior, permissions, finalization, Office threading or tool authorization.
- [x] Run external static/range regression checks for this package; no tests were inserted into runtime code.

## Short user smoke test after package 01

- [ ] Start Outlook once and Word once; confirm Red Ink loads normally and there is no new test UI/output.
- [ ] Run one ordinary AutoPilot job that reads/extracts a document and completes successfully.
- [ ] Run one document/job large enough to use `context_expand`; confirm it completes without an `ArgumentOutOfRangeException`/substring error at the end of a stored result.
- [ ] If tooling logging is enabled, check that a terminal `context_expand` window has `returned_chars = 0`, `next_offset = total_chars`, `truncated = false` when EOF is requested.

## Next implementation packages

- [x] P02 compaction/reread contract: reproduced with the supplied long-task log and implemented the narrow visibility-aware replay correction in package 02.
- [x] P03 reference lifecycle: workflow-bound reference access and response-owned generic replay-reference reuse implemented after the package-02 visibility/reread correction; cleanup/delegation contracts externally characterized.
- [x] P04 bounded diagnostics: real P03 log confirms the existing per-turn replay/retention diagnostics are sufficient; no aggregate runtime summary added because it would add log volume without resolving a material ambiguity.
- [x] P05 safeguard characterization: current timeout, repeated-failure, ordered-batch and final-delivery safeguards reviewed against real logs; no generic relaxation justified, so production guard behavior remains unchanged.
- [ ] Reassess concurrency only after per-run mutable state, model configuration isolation, Office COM/STA access, replay/idempotency and rate-limit handling have explicit contracts.
- [ ] Reassess request timeout/hedging only after cancellation, duplicate side effects, provider semantics and cost accounting are provider-agnostically defined.
- [ ] Reassess batching/caching only after partial-failure, invalidation and provider-neutral fallback contracts are defined.

## Release gates still open

- [ ] Windows/Visual Studio 2022 compile of SharedLibrary, Outlook add-in and Word add-in. NOT RUN in this Linux environment.
- [ ] Outlook VSTO smoke test. NOT RUN here.
- [ ] Word VSTO smoke test. NOT RUN here.
- [ ] One real AutoPilot regression run with the normal model/tool mix. NOT RUN here.
- [ ] One large-result/compaction regression run. NOT RUN here.
- [ ] Review startup timing logs separately if startup remains slow; the supplied source contained no startup invocation of the removed Agent self-tests, so no startup-speed claim is made from their removal alone.

## Completed in package 02

- [x] Review `RI_Tooling_Log(20260926-225452).txt` as a long-task regression sample before proceeding.
- [x] Confirm P01 paging behavior in the log: sequential windows advance correctly and the terminal window ends at `next_offset = total_chars` with `truncated = false`.
- [x] Reproduce the P02/P03 interaction from the log: a previously returned `context_expand` window was compacted to a stub and a later legitimate reread of that exact window was incorrectly suppressed by the run-wide no-progress set.
- [x] Replace run-wide "ever expanded" suppression with replay-visibility-aware suppression in both Outlook and Word.
- [x] Canonicalize expansion-window identity from the actual stored result and normalized returned range, so omitted/default arguments, clamped `max_chars`, negative starts and beyond-EOF starts are treated consistently.
- [x] Count only successful prior `context_expand` results when deciding whether a repeated read is redundant; failed attempts do not poison a later corrected reread.
- [x] Continue suppressing exact duplicate expansion windows when their full response is still available to the model or when the same window was already produced in the same tool-call batch.
- [x] Allow the exact same window to be read again after its full body has been compacted out of active replay.
- [x] Ensure a current-turn `context_expand` result is delivered losslessly at least once in sub-agent mode before historical compaction may replace it with a navigation stub.
- [x] Add bounded diagnostics distinguishing `context_expand` suppression while visible from legitimate reread after compaction.
- [x] Keep payload budgets, the public `context_expand` schema/limits, provider/model selection, generic tool retry/failure behavior, result-store lifetime, permissions, Office threading and artifact finalization unchanged.
- [x] Do not add integrated/runtime self-tests.

## Short user smoke test after package 02

- [ ] Re-run one long document task that uses several `context_expand` calls. It should complete without a `no_progress_suppressed` failure when the model legitimately rereads a window that has already been compacted to a stub.
- [ ] In the Tooling log, if such a reread occurs after compaction, look for `context_expand reread allowed after compaction` followed by a normal successful `context_expand` result.
- [ ] Confirm normal paging still advances (`start_char` / `next_offset`) and the final window still ends with `truncated = false`.
- [ ] Run one agent/skill with a large stored result. The first response to each requested `context_expand` window must be present in full for the next sub-agent model turn; only older windows may become compact stubs afterwards.
- [ ] Confirm Outlook and Word still load normally; package 02 adds no startup tests or test execution path.

## Next after package 02 validation

- [ ] Finish P03 reference-lifecycle characterization across generic large results, canonical parent-to-sub-agent source handles, voluntary compaction and cleanup. Do not introduce reference stabilization unless an actual instability remains after the package 02 replay fix.
- [ ] Finish P04 low-overhead diagnostics with an end-of-run compaction/expansion summary only if the new per-decision diagnostics are insufficient in real logs; do not add high-volume logging.
- [ ] P05: characterize timeout, repeated-failure, queue and final-delivery safeguards against current logs before changing behavior.
- [ ] Reassess concurrency only after per-run mutable state, model configuration isolation, Office COM/STA access, replay/idempotency and rate-limit handling have explicit contracts.
- [ ] Reassess request timeout/hedging only after cancellation, duplicate side effects, provider semantics and cost accounting are provider-agnostically defined.
- [ ] Reassess batching/caching only after partial-failure, invalidation and provider-neutral fallback contracts are defined.

## Additional package 02 release gates

- [ ] Windows/Visual Studio 2022 compile after package 02. NOT RUN in this Linux environment.
- [ ] Outlook long-task regression run proving the former suppressed reread now succeeds. NOT RUN here.
- [ ] At least one sub-agent run with a newly expanded window proving first-delivery visibility before historical compaction. NOT RUN here.
- [ ] Word tooling regression run covering the same replay behavior. NOT RUN here.


## Completed in package 02A - PDF to Word Word-only hardening

- [x] Make `pdf_to_word` a hard Microsoft-Word-only conversion path.
- [x] Always create a dedicated `Microsoft.Office.Interop.Word.Application` instance for PDF conversion; do not reuse another application's handler or Windows file association.
- [x] Remove every `pdf_to_word` execution path to `IPdfWordLayoutAdapter` / registered layout OCR adapters.
- [x] Reject legacy `conversion_mode=layout_ocr` instead of dispatching to a non-Word adapter.
- [x] Keep `auto`, `native`, `layout_preferred`, and `layout_required` input modes for compatibility, but route all of them through the same Word-only engine.
- [x] SUPERSEDED IN P02B: the `FileConverters` preflight proved too conservative because an installed/advertised converter does not establish what direct `Word.Documents.Open` will actually use.
- [x] SUPERSEDED IN P02B: converter-inventory inspection is no longer a gate; the direct Word COM attempt itself is authoritative.
- [x] SUPERSEDED IN P02B: `Word.Documents.Open` + `Document.SaveAs2(...wdFormatXMLDocument)` is now attempted directly, irrespective of installed third-party PDF software.
- [x] Preserve the existing image-only quality check; `layout_required` now fails without invoking an OCR/layout adapter, while other compatible modes may return Word's image-only result explicitly marked degraded.
- [x] Add structured result metadata `converter_policy=word_only_fail_closed`; keep prior provenance/quality fields and leave adapter ids empty.
- [x] Update the host tool description, lazy tool-loader guidance, shared tool list, and `.inky` guidance so models cannot be instructed to use an external PDF application through `pdf_to_word`.
- [x] No integrated/runtime self-tests added.
- [x] No changes to compaction, context replay, model/provider selection, Office mutation tools, generic retry/finalization, or `.inky` behavior outside the PDF-to-Word routing guidance.

## Short user smoke test after package 02A

- [ ] With Acrobat/PowerPDF installed and associated as the default PDF application, run `pdf_to_word` on a normal text PDF. Verify that no Acrobat/PowerPDF window/process is opened by Red Ink and the Tooling log says `Converting PDF to Word with Microsoft Word only`.
- [ ] Confirm the successful result JSON contains `converter_policy":"word_only_fail_closed"` and `converter_provenance":"word_native_verified"`.
- [ ] SUPERSEDED BY P02B TEST: with a third-party PDF product installed, confirm Red Ink still attempts direct `Word.Documents.Open` first and does not reject merely because Word advertises a PDF converter.
- [ ] Test a scanned/image-only PDF once. In ordinary modes the result may be explicitly `degraded_image_only`; with `conversion_mode=layout_required` the tool must fail and must not start any OCR/layout adapter.
- [ ] Confirm the previous package-02 long-document/context-expand behavior is unchanged when convenient.

## Next after package 02A validation

- [ ] Resume P03 reference-lifecycle characterization across generic large results, canonical parent-to-sub-agent source handles, voluntary compaction and cleanup.
- [ ] Finish P04 low-overhead diagnostics only if the current per-decision diagnostics remain insufficient in real logs.
- [ ] P05: characterize timeout, repeated-failure, queue and final-delivery safeguards before changing behavior.
- [ ] Reassess concurrency, hedging, batching and caching only after their existing roadmap gates are satisfied.

## Additional package 02A release gates

- [ ] Windows/Visual Studio 2022 compile after package 02A. NOT RUN in this Linux environment.
- [ ] Real Windows test with Microsoft Word plus at least one third-party PDF application installed. NOT RUN here.
- [x] SUPERSEDED IN P02B: FileConverters inventory is no longer used as a conversion gate.


## Completed in package 02B - direct Word conversion first + alternative-path recovery

- [x] Review `RI_Tooling_Log(20260927-054755).txt` and distinguish the successful markup retry from the unrelated finalization failure.
- [x] Confirm the 10 requested `word_markup` operations in that run were all eventually applied: 9 in the first batch and the one `no_match` operation on the targeted retry.
- [x] Remove the `FileConverters` preflight as a blocker. Installed/advertised third-party PDF converters no longer prevent a direct Word attempt.
- [x] Keep a dedicated `Microsoft.Office.Interop.Word.Application` instance and invoke `Word.Documents.Open(...wdOpenFormatAuto)` followed by `Document.SaveAs2(...wdFormatXMLDocument)` as the primary path.
- [x] Keep Windows file associations, Shell/Open-With, external PDF applications and Red Ink layout/OCR adapters out of the `pdf_to_word` execution path.
- [x] Preserve the controlled OCR/text reconstruction fallback only after the direct Word attempt itself fails or is unsuitable for `layout_required`.
- [x] Put stable PDF-to-Word failure identifiers in `ToolResponse.ErrorCode` and human-readable detail in `ToolResponse.ErrorMessage`; do not leave the host repair log with an empty error code.
- [x] Add a generic tool-declared `AllowCrossScopeAlternativeRecovery` contract to both Outlook and Word `ToolResponse` surfaces.
- [x] Keep cross-scope recovery disabled by default. Only the concrete failed tool can opt in.
- [x] In the shared sequencing state, allow an opted-in non-terminal skip failure to be superseded by a materially different successful path even when opaque operation ids differ.
- [x] Do not clear such a failure merely because another call ran: administrative/read-only progress remains insufficient, and a created-deliverable task requires a deliverable-capable successful call plus a host-validated deliverable.
- [x] Support an optional tool-declared replacement artifact extension; `pdf_to_word` requires a validated `.docx` replacement so an unrelated deliverable type cannot clear the conversion failure.
- [x] Preserve exact same-tool retry scope matching; the new cross-scope permission is for a materially different fallback path, not for disguising a retry with another operation id.
- [x] Keep the implementation generic: common recovery logic knows no PDF, Word, organization, provider, model or skill names. Tool-specific fallback eligibility is declared by the tool response, not hard-coded in the host.
- [x] No integrated/runtime self-tests added.

## Short user smoke test after package 02B

- [ ] With PowerPDF/Acrobat still installed, run the same PDF→Word task. The log should say `Converting PDF to Word through direct Microsoft Word automation` and Word should get the first conversion attempt instead of `word_only_external_pdf_converter_detected`.
- [ ] On direct Word success, confirm result metadata contains `converter_policy":"direct_word_com_first` and `converter_provenance":"word_documents_open_saveas2`.
- [ ] Confirm no PowerPDF/Acrobat window is opened by Red Ink and no Shell/Open-With path is used.
- [ ] If direct Word conversion genuinely fails, allow the existing reconstruction fallback to run. After a valid replacement Word deliverable is created, the log should contain `Recovered prior tool failure after successful logical replacement` (or equivalent recovery diagnostic) and finalization must not loop five times on the obsolete `pdf_to_word` failure.
- [ ] For the markup case, confirm a partial batch followed by a successful targeted retry is accepted as completed and the run can end `Success: True` when no other blocking failure remains.
- [ ] Confirm the P02 long-document/context-expand behavior remains unchanged when convenient.

## Next after package 02B validation

- [ ] Resume P03 reference-lifecycle characterization.
- [ ] Finish P04 diagnostics only if current logs remain insufficient.
- [ ] Continue P05 failure/finalization characterization using this recovered-fallback case as a regression fixture; do not relax unrelated failures.
- [ ] Reassess concurrency, hedging, batching and caching only after their existing roadmap gates are satisfied.

## Additional package 02B release gates

- [ ] Windows/Visual Studio 2022 compile after package 02B. NOT RUN in this Linux environment.
- [ ] Real Word COM conversion test with PowerPDF/Acrobat installed. NOT RUN here.
- [ ] One forced direct-Word failure followed by reconstruction fallback, proving obsolete failure state is cleared only after validated replacement output. NOT RUN here.
- [ ] Word-host regression for the generic `AllowCrossScopeAlternativeRecovery` plumbing. NOT RUN here.

## Completed in package 02C - isolate Word PDF reflow from third-party Word integration

- [x] Review `RI_Tooling_Log(20260927-061736).txt`: the run completed successfully with 14/14 tool calls and the P02 repeated-`context_expand` fix behaved correctly, but the direct `Word.Documents.Open` call still allowed installed PowerPDF integration to take over interactively.
- [x] Research Microsoft Word PDF-open behavior and PowerPDF Office integration. Microsoft exposes no PDF member in `WdOpenFormat`, so `Documents.Open(...wdOpenFormatAuto)` cannot explicitly select a built-in PDF parser.
- [x] Replace `New Microsoft.Office.Interop.Word.Application()` as the PDF conversion bootstrap with an explicitly resolved `WINWORD.EXE /safe /n` process. Microsoft documents `/safe` as Office Safe Mode and `/n` as a separate Word instance.
- [x] Resolve `WINWORD.EXE` without any file association: first from Windows App Paths, then from the registered `Word.Application` CLSID `LocalServer32` path.
- [x] Create a unique temporary RTF bootstrap document, launch only that file in the isolated Word process, enumerate the Running Object Table, and bind only to the Word `Document` whose full path and owning process id both match the process Red Ink started.
- [x] Do not use `Marshal.GetActiveObject("Word.Application")` for this path; it may bind to an unrelated pre-existing Word instance when multiple instances exist.
- [x] In the isolated Word instance, inspect the Word `FileConverters` inventory before opening the PDF. If an open-capable PDF converter is still advertised, fail before `Documents.Open` and allow the existing controlled `.docx` reconstruction fallback instead of risking interactive third-party conversion UI.
- [x] If converter inventory inspection itself cannot be completed reliably, fail closed to reconstruction rather than opening the PDF with unknown converter provenance.
- [x] Only after the isolated instance passes that guard, call `Word.Documents.Open(...wdOpenFormatAuto)` and `Document.SaveAs2(...wdFormatXMLDocument)`.
- [x] Keep the isolated Word process separate from any user Word instance; verify the automation `Hwnd` process id before touching it, then close/release COM objects and terminate only the process Red Ink itself started if normal `Quit` does not finish.
- [x] Keep the P02B alternative-path recovery contract unchanged: a failed isolated Word attempt can be superseded only by a validated replacement `.docx`.
- [x] No provider/model/organization/skill-specific runtime branch added. No integrated/runtime self-tests added.

## Short user smoke test after package 02C

- [ ] Repeat the same PDF→Word review with PowerPDF still installed. The Tooling log should say `Converting PDF to Word through isolated Microsoft Word Safe Mode`.
- [ ] PowerPDF/Convert Assistant must not request any manual interaction.
- [ ] If isolated Word can use its native PDF reflow, `pdf_to_word` should succeed and return `converter_policy":"isolated_word_safe_mode_first` plus `converter_provenance":"word_safe_mode_documents_open_saveas2`.
- [ ] If the isolated Word instance still advertises an external PDF converter, `pdf_to_word` should fail before opening the PDF with `word_safe_mode_external_pdf_converter_present`; AutoPilot should then use the existing OCR/text reconstruction fallback and still finish successfully with a validated `.docx`.
- [ ] Confirm the user's already-open Word documents/windows remain untouched.
- [ ] Confirm P02 long-document reread still logs `context_expand reread allowed after compaction` when applicable.

## Next after package 02C validation

- [ ] Resume P03 reference-lifecycle characterization across generic large results, canonical parent-to-sub-agent source handles, voluntary compaction and cleanup.
- [ ] Finish P04 low-overhead diagnostics only if real logs still leave a material ambiguity.
- [ ] Continue P05 timeout/failure/finalization/queue/delivery characterization without relaxing unrelated failures.
- [ ] Reassess concurrency, hedging, batching and caching only after the existing roadmap gates are satisfied.

## Additional package 02C release gates

- [ ] Windows/Visual Studio 2022 compile after package 02C. NOT RUN in this Linux environment.
- [ ] Real Office test proving `/safe /n` suppresses PowerPDF's interactive Word-open integration on the target installation. NOT RUN here.
- [ ] Real Office test of the guarded fallback path when the isolated Safe Mode instance still advertises a PDF `FileConverter`. NOT RUN here.
- [ ] Word-host regression for the shared P01/P02/P02B runtime changes. NOT RUN here.

## Compile fix after package 02C

- [x] Correct the two P02C Win32 process-id calls to the actual compiled SharedLibrary namespace: `Global.SharedLibrary.SharedLibrary.NativeMethods`.
- [x] Remove the `System.IntPtr` overload ambiguity under `Option Strict On` by converting Word `Hwnd` explicitly with `System.Convert.ToInt64(...)` before constructing `System.IntPtr`.
- [x] No runtime behavior, converter policy, fallback logic, compaction logic or tool contract changed.
- [x] No integrated/runtime self-tests added.
- [ ] Windows/Visual Studio 2022 compile after this compile fix. NOT RUN in this Linux environment.

## Package 02D - remove Office Safe Mode prompt; isolate Word by suspending connected add-ins

- [x] Review the user screenshot showing the modal Office `Default File Types` prompt caused by the P02C Safe Mode launch path.
- [x] SUPERSEDE the P02C `WINWORD.EXE /safe /n` bootstrap. Do not use Office Safe Mode or `/a` for `pdf_to_word`, because diagnostic startup modes can change preference loading/saving and surface first-run/default-file-type UI that blocks unattended AutoPilot runs.
- [x] Restore a dedicated normal `Word.Application` automation instance as the primary Word path. Word is a multiple-instance Automation server, so this does not require attaching to the user's already-running Word instance.
- [x] Before opening the PDF, temporarily disconnect every COM/VSTO add-in that is currently connected in that dedicated Word instance. Record each disconnected add-in by ProgID so its original connection state can be restored before the dedicated Word instance exits.
- [x] If a connected add-in cannot be identified or disconnected safely, fail before opening the PDF and allow the existing validated `.docx` reconstruction fallback.
- [x] After add-in suspension, inspect Word's `FileConverters` inventory. If an open-capable PDF converter is still advertised, fail before `Documents.Open` and use the reconstruction fallback instead of risking interactive third-party conversion UI.
- [x] Only when the dedicated Word instance passes the add-in/converter guard, call `Word.Documents.Open(...wdOpenFormatAuto)` followed by `Document.SaveAs2(...wdFormatXMLDocument)`.
- [x] Restore suspended COM/VSTO add-ins before `Word.Application.Quit`; do not make permanent registry, GPO, Office-profile, PowerPDF-specific, or file-association changes.
- [x] Update `pdf_to_word` tool descriptions, lazy loader guidance, shared tool lists and `.inky` guidance to match the new non-Safe-Mode policy.
- [x] Remove the P02C Safe Mode process bootstrap, ROT binding, WINWORD path discovery and related Win32 process-id code from `ThisAddIn.AutoPIlot.Tools.Office.vb`; the preceding compile-only NativeMethods fix is therefore no longer needed in this path.
- [x] No integrated/runtime self-tests added.

## Short user smoke test after package 02D

- [ ] Run the same PDF→Word review with PowerPDF installed. No `Default File Types` dialog may appear.
- [ ] The log should say `Converting PDF to Word through dedicated Microsoft Word automation with connected COM add-ins suspended`.
- [ ] PowerPDF/Convert Assistant must not request manual interaction.
- [ ] If Word can use its own PDF reflow after add-in suspension, the result should contain `converter_policy":"word_com_addins_suspended_first` and `converter_provenance":"word_documents_open_saveas2_com_addins_suspended`.
- [ ] If a PDF `FileConverter` remains visible after add-in suspension, `pdf_to_word` should fail before opening the PDF with `word_external_pdf_converter_still_present`; AutoPilot should then use the existing reconstruction fallback and still complete with a validated `.docx`.
- [ ] Confirm a separately opened user Word window/document remains unaffected.

## Next after package 02D validation

- [ ] Resume P03 reference-lifecycle characterization across generic large results, parent/sub-agent handles, compaction and cleanup.
- [ ] Finish P04 diagnostics only if real logs still leave material ambiguity.
- [ ] Continue P05 timeout/failure/finalization/queue/delivery characterization without relaxing unrelated failures.
- [ ] Reassess concurrency, hedging, batching and caching only after the existing roadmap gates are satisfied.

## Additional package 02D release gates

- [ ] Windows/Visual Studio 2022 compile after package 02D. NOT RUN in this Linux environment.
- [ ] Real Office test with PowerPDF installed proving there is no Safe Mode/first-run dialog and no interactive PowerPDF takeover. NOT RUN here.
- [ ] Real Office test proving suspended add-ins are restored after the dedicated Word instance exits. NOT RUN here.

## Completed in package 03 - reference lifecycle

- [x] Characterize `result_ref` creation, replay reuse, canonical parent-to-sub-agent handles, voluntary compaction interaction and end-of-workflow cleanup against the current P02D basis.
- [x] Keep the package-02 visibility-aware reread contract unchanged: a stable reference does not make a compacted window count as currently visible.
- [x] Bind active-workflow `result_ref` resolution to the current `WorkflowId`. A reference owned by another workflow, or an unscoped stored reference presented inside an active workflow, is treated as unavailable.
- [x] Apply the same workflow ownership check before canonical `context_expand` window-key construction and before accepting `canonical_source_result_refs` for a delegated agent.
- [x] Stabilize only response-owned generic replay references. Re-rendering the same immutable `ToolResponse` may reuse its prior `result_ref` only when workflow, producing tool and full body still match exactly.
- [x] Do not add global or semantic content deduplication. Identical content from a different response does not become the same stored result merely because its text matches.
- [x] Rebuild the compact preview on every replay pass using the current threshold/preview settings; a previously larger preview is never reused merely because the stored reference is reused.
- [x] Preserve the established top-level cleanup rule: sub-agents share the parent workflow and do not clear the result store; the top-level run clears that workflow in `Finally`.
- [x] Preserve P01 EOF/range normalization, P02 first-delivery visibility/reread behavior, payload budgets, `context_expand` public schema and all retry/finalization/Office behavior.
- [x] No integrated/runtime self-tests added. P03 QA is external/static/model-based only.

## Short user smoke test after package 03

- [ ] Run one long document task with several `context_expand` calls. Repeated historical replay of the same large source should keep the same `result_ref`; legitimate reread after compaction must still be allowed as in P02.
- [ ] Run one skill/agent path using `canonical_source_result_refs`. Parent and sub-agent should read the same handle successfully while they share the workflow.
- [ ] Confirm a normal run still ends cleanly and a subsequent new run cannot use a stale `result_ref` from the completed workflow (`unknown_result_ref` is the expected safe outcome if such a stale token is deliberately retried).
- [ ] Confirm Outlook and Word show identical replay behavior. No startup/self-test path was added.

## Next after package 03 validation

- [ ] P04: use the existing bounded replay diagnostics first; add only a compact aggregate end-of-run summary if real logs still leave material ambiguity.
- [ ] P05: characterize timeout, repeated-failure, queue and final-delivery safeguards externally before changing behavior.
- [ ] Reassess concurrency only after mutable model/run state, Office COM/STA affinity, reference ownership, idempotency and rate-limit contracts are explicit.
- [ ] Reassess request hedging only after cancellation, duplicate side effects, provider semantics and cost accounting are defined.
- [ ] Reassess batching/caching only after partial-failure, invalidation and provider-neutral fallback contracts are defined.

## Additional package 03 release gates

- [ ] Windows/Visual Studio 2022 compile of SharedLibrary, Outlook add-in and Word add-in after P03. NOT RUN in this Linux environment.
- [ ] Outlook long-task run confirming stable replay refs plus P02 reread-after-compaction behavior. NOT RUN here.
- [ ] Word-host equivalent replay run. NOT RUN here.
- [ ] Canonical parent-to-sub-agent source-handle run. NOT RUN here.

## Real-run validation after package 03

- [x] Review `RI_Tooling_Log(20260927-072722).txt` from the P03 test document.
- [x] Confirm the large `read_attachment` result was compacted under the stable reference `tref_000001` and the same reference persisted across later replay rebuilds.
- [x] Confirm `context_expand` later resolved the same `tref_000001` successfully.
- [x] Confirm the run completed normally: 19 iterations, 36 tool-call records, 36 successful, 0 failed, no unresolved tool failure.
- [x] P04 conclusion: the current per-turn `Tool replay retention` lines plus explicit reread diagnostics provide enough evidence for replay/ref debugging. Do not add a permanent aggregate end-of-run diagnostic unless a future real log exposes a concrete ambiguity.

## P05 safeguard characterization

- [x] Keep the existing consecutive unrecovered tool-failure breaker at 3. Recoverable planning/input failures already reset that counter; no real log shows a safe reason to weaken or accelerate it globally.
- [x] Keep the existing duplicate-execution breaker at 3 and the visibility-aware `context_expand` no-progress guard.
- [x] Keep `MaxContinuationRetries = 5`. The earlier five-round PDF finalization loop was caused by stale unresolved-failure state and was corrected in P02B; the current P03 run finalizes immediately once complete.
- [x] Keep current LLM timeout/transport-retry profiles and heavy-call timeout scaling. No supplied log demonstrates a generic timeout policy defect after the existing safeguards.
- [x] Keep ordered sequential tool-batch execution. In the P03 real run, tool execution time was negligible relative to model time, so concurrency would add state/COM/idempotency risk without material benefit in this scenario.
- [x] Keep artifact/deliverable completion gates unchanged. The P03 run produced and promoted the requested DOCX and then finalized normally.
- [x] No integrated/runtime self-tests added.

## Round-trip reduction after P05

- [x] Add a shared, provider/tool-agnostic sequencing instruction that prefers emitting several already-known independent calls in the same model turn, especially independent read-only lookups.
- [x] Preserve strict ordered execution; this is model-turn batching only, not parallel tool execution.
- [x] Explicitly discourage batching mutations merely for speed and preserve the existing dependency rule: if later arguments depend on an earlier result, wait for that result.
- [x] Apply the instruction through the existing shared `DependentBatchingInstruction`, which is already consumed by both Outlook and Word hosts.
- [x] Do not change tool schemas, payload budgets, compaction policy, retry/failure behavior, Office threading, queues or finalization.

## Short user smoke test after round-trip guidance

- [ ] Re-run the same P03 test document and prompt with the same model. Functional output must remain equivalent and the run must still finish with zero failed tool calls.
- [ ] Compare model iterations/tool batches rather than raw tool execution time. When several independent lookups are known, look for multiple calls in one assistant turn instead of a separate model turn for each lookup.
- [ ] `Tool-call batch execution mode` must remain `sequential_ordered_no_deferral`; there must be no claim or behavior of parallel execution.
- [ ] Stable `result_ref` behavior from P03 and reread-after-compaction behavior from P02 must remain unchanged.

## Next after this package

- [ ] Reassess provider/model-call caching only if a concrete repeated-identical model-request pattern is demonstrated; do not cache generative responses solely by prompt text.
- [ ] Reassess runtime concurrency only on workloads where tool execution itself is a material share of elapsed time and after mutable run state, Office COM/STA, idempotency and rate-limit contracts are explicit.
- [ ] Reassess hedged model requests only after cancellation, duplicate side effects, provider semantics and cost accounting are provider-agnostically defined.
- [ ] Reassess domain/tool-level batching only where a generic partial-failure contract exists; prefer the new same-turn independent-call guidance before adding new batch schemas.

# Host improvement stream – cumulative continuation through H24c

## Current implementation basis after Host improvement stream

- [x] Current source basis for this stream is `RI_Gen2_260928_1311.zip` plus the cumulative changed-only package `RedInk_Gen2_Host_Improvements_H24c_changed_only.zip`.
- [x] The changed-only package contains all cumulative Host-improvement source changes, not merely the final compile-fix file. This is intentional so a partial earlier overlay cannot leave host call sites and SharedLibrary signatures out of sync.
- [x] Original source inventory remains 596 files; cumulative changed-only inventory remains 18 files.
- [x] No integrated/runtime/startup self-tests were added.
- [x] Shared behavior remains model-, provider-, tool-, skill-, organization-, template-, document- and language-agnostic where the concern is common; Word and Outlook use symmetric host plumbing for common behavior.
- [x] Existing P01/P02/P03/P05 replay/reference/safeguard contracts remain protected unless explicitly listed below.

## Completed Host improvements

- [x] H-01: preserve Local Chat terminal error state and already-produced assistant content without falsely converting failed jobs into success.
- [x] H-03: add narrow retry-scope rebinding only for the invalid/incomplete explicit-artifact-identity preflight where no valid logical scope can yet exist; default recovery remains scope-exact.
- [x] H-04: align empty explicit artifact identity presence semantics and improve the corresponding diagnostic without placeholder-word heuristics.
- [x] H-05: correct the generic `skill_use` schema and keep generic/dynamic skill scope enforcement symmetric.
- [x] H5b: make `tool_loader` report `loaded` / `already_loaded` only when the current model adapter can actually expose the tool definition, using the already-existing adapter conversion rather than a second schema engine.
- [x] H-06: deny writes under configured central `skills/` / `agents/` resource roots unless central writes were explicitly allowed; preserve reads, local resource writes and existing opt-in semantics.
- [x] H-07/H-23: remove contradictory authoring guidance such as one-file-per-turn and temporary-workspace guidance for persistent resource edits; keep the smallest complete authorized edit as the generic contract.
- [x] H-08: carry the latest host-authoritative raw user request separately into skill execution as `user_request`; keep the orchestration `input` for compatibility but mark it non-authoritative.
- [x] H-09/H-10: separate Local Chat model capability, persisted Tooling preference and temporary AutoPilot/Gate availability; keep the existing AgentGate serialization and prevent status reads from destroying the stored Tooling preference.
- [x] H-11: enforce the dialog-owner SameThread/watchdog rule on the direct STA `ask_user` path while preserving the already-safe ownerless MTA fallback.
- [x] H-11 rendering remainder: protect Windows-path tokens in the `ask_user` question from Markdig escape interpretation without replacing the global Markdown renderer.
- [x] H-12: use one canonical trailing `TASK_STATUS` envelope interpretation so examples/code/quotes are not mistaken for the protocol footer while true duplicate trailing footers and malformed final envelopes remain detectable.
- [x] H-13: make `DoubleS` a prose-presentation transform that preserves code, structured JSON-like content, paths/URLs, quotes, tags, blockquotes and `TASK_STATUS` protocol data.
- [x] H-14: add bounded/data-minimizing tool-call correlation diagnostics with run/call/operation/step identifiers and path-like target references; avoid free-form content logging.
- [x] H-15: prevent the advisory `js_run` Node-API preflight from triggering on strings/comments/template text while leaving the WebView2 sandbox unchanged as the actual execution boundary.
- [x] H-16/H-24: preserve first visibility of fresh reference-bearing results in the normal budgeted replay path while under the existing total budget; only the final existing budget-pressure stage may compact such current results. Current-turn `context_expand` remains lossless.
- [x] H-18: validate actual loader-selected fallback skill descriptors in the authoring postcondition and reject unsupported YAML block-scalar syntax with a clear diagnostic; keep tolerant loader compatibility for existing resources.
- [x] H-19: enforce the same allowed-tool scope for the generic `skill_use` entry as for named `skill_<name>` entry points while retaining mandatory host runtime primitives.
- [x] H-20: make the Local Chat slash trigger unambiguous and close cancellation turns with `Aborted by user.` when needed so a user turn is not left orphaned.
- [x] H-21a: use `Inky_for_AutoPilot.md` for AutoPilot guidance when present, with fallback to normal `Inky.md`; do not inject the global author-mode prompt into AutoPilot.
- [x] H-22: correct scope-denial diagnostics from the misleading `sub-agent` wording to generic enforced-tool-scope wording without changing the scope decision.
- [x] Reject malformed non-locked `expected_artifacts` metadata before side effects while preserving absent metadata and `[]` compatibility and without relaxing final-delivery gates.

## Characterized and deliberately not changed

- [x] H-02 broad persistence/delivery reinterpretation was not implemented. `expected_artifacts` remains a run-level delivery contract; working/intermediate/resource paths do not prove that the contract is accidental.
- [x] H-17 additional persistent resource write surfaces remain a separate permission/product decision; no new write rights were granted.
- [x] H-21b hard-coded `skill-author` routing text in generic `ToolLoaderTool` is confirmed architecture debt but was not removed without a real routing regression fixture; no new AutoPilot or skill-specific host flag was added.
- [x] H-22 dormant `required-successful-tools-before-final-mutation` was not activated because current authoring guidance documents it as ineffective and existing resources may already contain the key.
- [x] H-25 generic `blocked requires prior tool call` enforcement was not added because legitimate blocked states can exist before any tool call and meaningless calls must not be incentivized.
- [x] O-01 automatic foreign-process attribution/rollback was not added; file hashes can establish mutation but cannot safely attribute the writer.
- [x] Browser `ask_user` resume workflow, full queueing, runtime concurrency, hedged model calls, new caching/batching semantics, new persistent write permissions and `js_run(script_path)` remain separate feature/architecture work.

## H24c – Visual Studio compile feedback and narrow compile correction

- [x] Review the real VS2022 compiler error list supplied after H24b.
- [x] Confirm the primary SharedLibrary source defect: the H13 `DoubleS` literal-protection helper used typographic quote glyphs directly as VB `Char` literals in `SharedMethods.MainLLM.vb`, producing BC30081 / BC30037 / BC30004 / BC30201 around the quote-matching block.
- [x] Replace the problematic direct smart-quote `Char` literals with explicit Unicode values via fully qualified `System.Convert.ToChar(&H201E)`, `System.Convert.ToChar(&H201C)` and `System.Convert.ToChar(&H201D)`.
- [x] Preserve the H13 behavior: straight quotes, German low/left/right smart quotes and unmatched quoted spans remain protected from `DoubleS` normalization.
- [x] Preserve UTF-8 BOM and CRLF line endings in `SharedMethods.MainLLM.vb`.
- [x] Verify that every SharedLibrary member/signature referenced by the dependent-project errors is present in the cumulative source: `InkyPromptBuilder.Build(... isAutoPilot ...)`, `ToolCallSequencing.BuildToolCallAuditDiagnostic`, `ValidateExpectedArtifactArguments`, the three-argument `ValidateExplicitArtifactIdentityArguments`, `NoteToolFailure(... allowRetryScopeRebinding ...)`, `AgentToolRouter.TryHandleAsync(... authoritativeUserRequest ...)`, and `SkillInvokeTool.Execute(... authoritativeUserRequest ...)`.
- [x] Treat those dependent Outlook/Word errors as expected cascade/stale-metadata symptoms while SharedLibrary cannot compile; do not roll back the current APIs merely to match stale compiled metadata.
- [x] H24c external QA round 1: 37/37 checks PASS.
- [x] H24c external QA round 2: 51/51 checks PASS from an independently reconstructed `RI_Gen2_260928_1311.zip + H24b` basis.
- [x] H24c overlay check: original basis + new cumulative changed-only ZIP reproduces the current 596-file source tree hash-for-hash.
- [x] H24c cumulative changed-only package still contains exactly 18 source files; only `SharedMethods.MainLLM.vb` differs from H24b.
- [ ] Windows/Visual Studio 2022 compile after H24c. REQUIRED; not executable in the Linux QA environment.

## H24c targeted Windows compile verification

- [ ] Replace/overlay the repository with the full cumulative `RedInk_Gen2_Host_Improvements_H24c_changed_only.zip`, not only `SharedMethods.MainLLM.vb`, so SharedLibrary declarations and Word/Outlook call sites are guaranteed to be from the same cumulative checkpoint.
- [ ] Clean the Visual Studio solution/build outputs if stale dependent-project diagnostics remain after the SharedLibrary syntax error is gone, then rebuild SharedLibrary first and the full solution second.
- [ ] Confirm no BC30081 / BC30037 / BC30004 / BC30201 remains in `SharedMethods.MainLLM.vb` around the smart-quote protection block.
- [ ] Confirm the earlier dependent-project BC30272 / BC30456 / BC30057 signature/member errors disappear once the current SharedLibrary successfully rebuilds.
- [ ] If any signature/member error remains after a successful SharedLibrary compile, capture the exact new error list plus the on-disk signatures from the referenced SharedLibrary source; do not weaken APIs as a workaround.

## Host-improvement smoke/regression tests still open on Windows

- [ ] Start Outlook and Word once after the successful full rebuild and confirm Red Ink loads normally.
- [ ] Run one ordinary Local Chat request with Tooling off and one with Tooling on; confirm preference/capability text is truthful and normal replies still persist.
- [ ] Start an AutoPilot job and inspect Local Chat state; AutoPilot may make Tooling temporarily unavailable but must not erase the stored Tooling preference.
- [ ] Trigger Local Chat cancellation once; confirm the turn closes with `Aborted by user.` when no assistant turn was already persisted.
- [ ] Exercise `/` in an empty Local Chat box and confirm the Prompt Library trigger; confirm a non-empty/whitespace-prefixed box does not spuriously trigger it.
- [ ] Run one `ask_user` request containing a Windows path with `.inky` or another Markdown-sensitive segment; confirm the displayed path remains exact and the dialog owner behavior is safe.
- [ ] Run one skill through generic `skill_use` and one through a named `skill_<name>` entry; confirm both enforce the same tool scope and receive the authoritative `user_request` separately from orchestration `input`.
- [ ] Run one authoring path using the loader-selected fallback descriptor form; confirm postcondition validation applies without changing loader compatibility for existing resources.
- [ ] Run one malformed `expected_artifacts` call and confirm it is rejected before side effects; confirm absent `expected_artifacts` and `[]` remain valid.
- [ ] Run one long result/read under the normal replay budget; confirm fresh content is visible before historical compaction and P02/P03 reread/reference behavior remains intact.
- [ ] Run one real large-result `context_expand` regression; confirm current-turn expansion is lossless, terminal windows remain correct and historical reread after compaction still works.
- [ ] Run a real AutoPilot regression with the normal model/tool mix and confirm final delivery/failure safeguards remain unchanged.

## Current release gates after H24c

- [ ] Windows/Visual Studio 2022 compile of SharedLibrary, Outlook add-in and Word add-in after H24c.
- [ ] Outlook VSTO smoke test after H24c.
- [ ] Word VSTO smoke test after H24c.
- [ ] One real AutoPilot end-to-end regression run after H24c.
- [ ] One large-result/compaction/context-expand regression run after H24c.
- [ ] Targeted Local Chat / `ask_user` Windows UI smoke tests after H24c.
- [x] Linux/static delivery QA completed; no claim of Windows/VSTO compilation success is made until the above compile gate is actually run.

## H24d – compile compatibility after second VS2022 feedback

- [x] Review the second real VS2022 compiler error list after H24c.
- [x] Identify the first/root SharedLibrary error as `BC30456: TryGetTopLevelOperationIdentity is not a member of ExplicitOperationRegistry` in `ToolCallSequencing.vb`.
- [x] Reconcile this against the uploaded `RI_Gen2_260928_1311.zip`: that uploaded basis already contains `ExplicitOperationRegistry.TryGetTopLevelOperationIdentity`, but the user's actually compiling local repository evidently has an older `ExplicitOperationRegistry` surface.
- [x] Do not add an unchanged `ExplicitOperationRegistry.vb` to the changed-only package merely to mask the mismatch.
- [x] Remove the unnecessary audit-only dependency instead: `BuildToolCallAuditDiagnostic` now reads top-level `operation_id` and `step_id` directly from the already available argument dictionary.
- [x] Preserve audit semantics: operation/step ids are still logged only as bounded metadata; no tool execution, sequencing, recovery, finalization or permission behavior changes.
- [x] Keep the H24c smart-quote compile correction unchanged.
- [x] Verify every SharedLibrary API behind the dependent Outlook/Word errors exists in the cumulative source and is included in the changed-only package: `InkyPromptBuilder.Build(... isAutoPilot ...)`, `BuildToolCallAuditDiagnostic`, `ValidateExpectedArtifactArguments`, three-argument `ValidateExplicitArtifactIdentityArguments`, `NoteToolFailure(... allowRetryScopeRebinding ...)`, `AgentToolRouter.TryHandleAsync(... authoritativeUserRequest ...)`, and `SkillInvokeTool.Execute(... authoritativeUserRequest ...)`.
- [x] H24d external QA round 1: 73/73 checks PASS.
- [x] H24d external QA round 2: 15/15 checks PASS from an independently reconstructed `RI_Gen2_260928_1311.zip + H24c` basis with only the H24d compatibility patch synthesized.
- [x] H24d overlay verification: original 596-file basis + H24d cumulative changed-only ZIP reproduces the current source tree hash-for-hash.
- [x] H24d cumulative changed-only package contains exactly 18 files, all genuinely changed relative to the uploaded 260928_1311 basis.
- [ ] Windows/Visual Studio 2022 compile after H24d. REQUIRED; not executable in the Linux QA environment.

## H24d targeted Windows compile verification

- [ ] Overlay the full cumulative `RedInk_Gen2_Host_Improvements_H24d_changed_only.zip` onto the repository.
- [ ] Clean SharedLibrary output/intermediate folders or use Visual Studio Clean Solution so dependent projects cannot continue to bind to an older SharedLibrary assembly.
- [ ] Build **SharedLibrary first**. The former `TryGetTopLevelOperationIdentity` error must be gone because H24d no longer references that member from `ToolCallSequencing`.
- [ ] Only after SharedLibrary builds successfully, rebuild the full solution.
- [ ] Confirm the dependent Outlook/Word signature errors disappear once they bind to the newly built SharedLibrary.
- [ ] If SharedLibrary still fails, capture **only the SharedLibrary errors first**; dependent-project signature errors are not actionable until SharedLibrary itself succeeds.
- [ ] If SharedLibrary succeeds but a dependent-project member/signature error remains, capture that exact error plus the on-disk signature from the corresponding SharedLibrary file before making any further API change.

## Current release gates after H24d

- [ ] Windows/Visual Studio 2022 compile of SharedLibrary, Outlook add-in and Word add-in after H24d.
- [ ] Outlook VSTO smoke test after H24d.
- [ ] Word VSTO smoke test after H24d.
- [ ] One real AutoPilot end-to-end regression run after H24d.
- [ ] One large-result/compaction/context-expand regression run after H24d.
- [ ] Targeted Local Chat / `ask_user` Windows UI smoke tests after H24d.
- [x] Linux/static H24d delivery QA completed; no claim of Windows/VSTO compilation success is made until the above compile gate is actually run.

## H24d audit correction stream

- [x] CP-GOV-001: Add permanent checkpoint/recovery/double-check QA rules to `CONTRIBUTING.md`; verify original content preservation plus UTF-8 BOM/CRLF independently.
- [x] Reconcile uploaded 2021 ZIP against the audit H24d manifest: common product content hashes match; two XLSX names are archive-decoding aliases with identical payloads; current `Tasks.md` required this cumulative H24d history merge.
- [x] Independently review original host recommendations H-01..H-25, O-01 and N-01 against reconstructed baseline + cumulative H24d, CONTRIBUTING and cumulative Tasks.
- [x] Record 12 correction findings and complete source/evidence/coverage matrix. No production source changes in this audit.
- [x] Run 19 actual extracted Local Chat JS tests with deterministic mocks, 10 integrity/structure checks and isolated Python/static counterexamples. These are NOT a VB/Windows/VSTO build.
- [x] K00: Verify actual local source/project/build basis; current uploaded source content matches the audited H24d manifest (apart from two filename-decoding aliases with identical payload hashes). No .NET/VB/MSBuild toolchain is available here, so Windows/VS2022 compile remains NOT RUN.
- [x] F02 / K01: Close overlapping/strict central write-root bypasses without changing read/local permissions or creating new write surfaces. Shared `PathPolicy` guard verified by two independent external/static checks; native Windows write/runtime verification remains an open K08 gate.
- [x] F05 / K02.1: Make expected_artifacts type validation total and deterministic, with no uncaught structured-ID conversion. Shared validator/source mechanics verified twice; native .NET/Windows execution remains NOT RUN.
- [x] F12 / K02.2: Align tool-owned normalization, exposed schema and generic preflight for host-owned Skill slot contracts; preserve caller-owned/locked/delegated contracts.
- [x] CP-K02.2-006: Neutral opt-in ModelConfig preflight normalizer added; generic and named skill configurations normalize host-owned positive deliverable-count slots before schema/signature/locked/expected-artifact validation. Word signature ordering aligned with Outlook. QA1, independent QA2 and full K02.1 regression rerun PASS; two harness-assumption failures documented as harness defects. Windows/Visual Studio 2022 compile NOT RUN.
- [x] CP-K02.3-007: F04 exact recovery correlation released. Valid operation/step scopes are retained; scope-less `explicit_artifact_identity_incomplete` and `invalid_expected_artifacts` recovery uses canonical host-owned argument snapshots masking only actually repairable invalid fields, and clears only one uniquely correlated candidate. Object-key order is non-semantic; arrays, values/types, opaque IDs and already-valid effect/identity arguments remain binding. Outlook/Word host plumbing is symmetric. QA1 50/50 PASS, independent QA2 35/35 PASS, K02.1 regression PASS and K02.2/F12 regression PASS. One real WIP ambiguity defect plus two QA-harness defects were caught before release. Windows/Visual Studio 2022 compile NOT RUN.
- [x] F04 / K02.3: Preserve valid operation scope; remove same-tool/epoch-only false recovery and require exact host-proven repair correlation.
- [x] F01 / K03: Apply model-exposure validation to already_loaded as well as new loads in both hosts.
- [x] CP-K03-008: F01 truthful loader availability released. Existing selected tools and `tool_loader` self-requests now pass the same current-model exposure decision as new loads; post-model availability requires the same tool-instruction wrapper/template that the next definition builder needs plus successful canonical conversion. Pre-model required-runtime-primitive selection remains unchanged. K03 QA1 46/46 PASS, independent QA2 28/28 PASS, tree QA 26/26 PASS, K02.1/K02.2/K02.3 regressions PASS. Two QA-harness defects were documented and rerun; no product fix was made for either. Windows/Visual Studio 2022 compile NOT RUN.
- [x] F03 / K04: Correct JSON-string-aware footer boundaries and tab/code context; preserve strict footer/memory validation and lossless strip.
- [x] CP-K04A-009: K04 prechange characterization from CP-K03-008 reproduced 6/24 expected F03 boundary/context failures with no product-source change; established protected normal/fence/quote/duplicate/malformed/status/reason/memory/line-ending cases.
- [x] CP-K04-010: Shared TASK_STATUS locator now uses the existing Newtonsoft JSON object parser to disambiguate valid body boundaries, treats literal tags inside JSON strings as data, rejects inline/tab/4-column code starts, keeps unique malformed JSON controlled, and preserves duplicate/footer/memory/egress contracts. K04 QA1 47/47, independent QA2 43/43, tree QA 25/25 and K02/K03 protected regressions PASS. Windows/Visual Studio 2022 compile NOT RUN.
- [x] CP-K05A-011: K05a/F06 prechange characterization from CP-K04-010 reproduced 9/16 path-protection deviations: tilde fence, four-space/tab code, HTML pre/code, Markdown link/image destinations and deliberately Markdown-escaped drive/UNC paths. No product source changed; owner/watchdog/MTA invariants remain intact. Rendering cross-check used an independent CommonMark-compatible proxy only, not VB/Markdig/Windows runtime.
- [x] F06 / K05a: Keep literal Markdown/HTML path blocks unchanged while rendering prose Windows paths exactly.
- [x] CP-K05A-012: F06 released. `ask_user` now selects only native Markdig `LiteralInline` source spans from the same precise-source-location pipeline, leaves structural/literal Markdown and HTML code/pre regions untouched, preserves deliberate escapes, and only protects prose separators that CommonMark would consume. The legacy one-Boolean pipeline overload remains intact; owner/watchdog/MTA code is unchanged. QA1 48/48, independent QA2 55/55, tree QA 40/40 and K02-K04 protected regressions PASS. Windows/Visual Studio 2022 compile NOT RUN.
- [x] CP-K05B-013: K05b/F07 prechange characterization from CP-K05A-012 reproduced 13/30 deviations, including relative/absolute link/image targets, reference definitions/IDs, LF/CRLF multiline code spans, HTML code/pre and TASK_STATUS with a literal close marker inside JSON. No K05b product source changed; existing prose/JSON/fence/indent/blockquote/quote/URL/Windows/UNC/line-ending behavior is the protected baseline.
- [x] CP-REC-014: Recovery metadata reconciliation after CP-K05B-013. `RESUME.md` had remained at CP-K05A-012 although STATE/checkpoint/Tasks/Worklog/QA and the MainLLM hash proved CP-K05B-013 was released. No product source changed; the stale resume was corrected and the evidence cross-check was checkpointed before K05b source work.
- [x] F07 / K05b: Preserve relative link/image/reference targets and multi-line literal/code regions during DoubleS presentation normalization.
- [x] CP-K05B-015: F07 released from CP-REC-014. `SharedMethods.MainLLM.vb` now uses the common precise-source Markdig pipeline and rewrites only proven presentation-prose `LiteralInline` source spans; top-level JSON, code/fences/indented code, HTML code/pre, blockquotes, quoted regions, link/image/reference structure, URL/Windows/UNC tokens and canonical trailing TASK_STATUS spans remain literal. K05b QA1 59/59 (40 contract cases), independent QA2 52/52, tree QA 74/74 and K02-K05a protected regressions PASS. One real WIP quote-across-AST-node risk and several harness-only failures were caught before release. Windows/Visual Studio 2022 compile NOT RUN.
- [x] CP-K06A-016: K06a/F08 prechange characterization from CP-K05B-015 reproduced 7/23 lexical deviations with no product-source change: four RegExp-literal false positives, executable `${process.env}`/`${require(...)}` template expressions hidden by blanket template masking, and a nested-template boundary that can expose literal text as code. Existing strings/comments, normal division and direct Node-global detection remain protected. Windows/Visual Studio 2022 compile NOT RUN.
- [x] F08 / K06a: Remove RegExp-literal Node-preflight false positives without weakening the real sandbox.
- [x] CP-K06A-017: F08 released. The advisory JavaScript lexical masker now treats RegExp bodies/classes/escapes/flags as literal data, re-enters executable code inside template `${...}` expressions (including nested templates), distinguishes division from sufficiently clear RegExp starts, and treats member names that spell JavaScript keywords as member names. Lexical uncertainty masks the remainder for the advisory guard and defers to the unchanged WebView2 sandbox rather than inventing a Node rejection. Fresh QA1 63/63, independent QA2 37/37 with 26 Node-vm oracle cases, 11 protected regression scripts and tree QA 78/78 PASS. One real unreleased member-keyword WIP defect plus the 596/597 local-baseline tree-harness assumption were corrected before release. Windows/Visual Studio 2022 compile NOT RUN.
- [x] CP-K06B-018: K06b/F10 prechange characterization from CP-K06A-017 reproduced the unsupported YAML block-scalar diagnostic gap without product-source change. Six required indicator/comment forms (`|2`, `>2-`, `|-2`, `>+2`, `|2 # comment`, `>2- # comment`) bypass the current guard while the simple runtime loader retains the header token and ignores the intended indented scalar text. Existing `|`, `>`, `|-`, `>+`, quoted scalar values, normal values/lists and runtime descriptor-selection behavior remain protected. Two independent external/static characterizations PASS; the first harness run misclassified expected negative findings as harness failures and was corrected without source changes. Windows/Visual Studio 2022 compile NOT RUN.
- [x] F10 / K06b: Diagnose all unsupported block-scalar header forms, including indentation indicators, without changing runtime YAML compatibility.
- [x] CP-K06B-019: F10 released. `TryGetUnsupportedFrontmatterSyntax` now recognizes YAML block-scalar indentation indicators (`1`–`9`) and the permitted chomping/indent indicator orders, with optional comments, while quoted scalar-looking values and invalid non-header forms remain ordinary values. `ParseSimpleYaml`, descriptor selection, `SkillAuthoringPostcondition`, the project file, Word/Outlook host code and all prior released product hashes are unchanged. QA1 48/48, independent QA2 13/13 over 214 generated grammar cases, 13 protected regression scripts and tree QA 76/76 PASS. One tree-QA attempt failed only because redundant method-span harness anchors included the changed diagnostic helper; source stayed unchanged and the corrected harness reran fully PASS. Windows/Visual Studio 2022 compile NOT RUN.
- [x] CP-K07A-020: K07a/F09 prechange characterization from CP-K06B-019; no product source change. Both hosts currently emit only `preflight/attempt` plus `tool_result/failed`: synthetic preflight/history responses, lazy-load deferral, executed success and controlled abort paths have no matching persisted audit outcome. The shared formatter also leaves several opaque identity fields, operation-step values, argument-key lists, path references and the total diagnostic unbounded; Run/Call/Tool/Recovery fields are not JSON-safe against embedded newlines. Existing full normal path visibility, 32-reference/operation caps, truncation metadata, long non-path reference hashing and exclusion of free-form `text`/`code` values remain protected. Two independent external/static characterizations PASS as negative reproduction; Windows/Visual Studio 2022 compile NOT RUN.
- [x] F09 / K07a: Log correlatable synthetic preflight outcomes and bound diagnostic identity values; no new free-form content dumps.
- [x] CP-K07A-021: F09 released. The shared audit formatter now JSON-escapes/bounds opaque run/call/tool/recovery/operation/step identities, caps operation/argument/reference traversal and total diagnostic size, preserves ordinary full paths, and hashes exceptional oversized values. Both host loops emit one `preflight/attempt` plus exactly one generic `Finally` outcome for executed, synthetic, deferred or aborted calls without changing `AddToolResponseToHistory`, recovery, delivery or tool semantics; diagnostic logger failures are swallowed. Persisted QA1 140/140, independent QA2 140/140, protected regression 134 checks PASS (11 historical harnesses direct PASS; 4 stale whole-file/hash assertions replaced with package-specific invariant checks), tree/encoding QA 83/83 PASS. Harness-only failures are recorded in `qa/K07a/K07A_QA_HARNESS_HISTORY.json`. Windows/Visual Studio 2022 compile NOT RUN.
- [x] F11 / K07b: Distinguish AutoPilot guidance read states and log controlled fallback; do not silently alter H24d global merge/fallback semantics.

- [x] CP-K07B-022: K07b/F11 prechange characterization from CP-K07A-021; no product source change. Current H24d semantics are global: central/local `Inky_for_AutoPilot.md` content is merged central-before-local and selected only when the combined specific content is non-whitespace; otherwise `InkyMdForAutoPilot` falls back to the normal combined `Inky.md`. Empty/whitespace specific files therefore still trigger the existing normal fallback, and mixed normal/specific roots are not merged rootwise. F11 is reproduced because not-configured, missing, empty/whitespace, read/access failure and file-disappeared-after-probe states are not distinguished or diagnosed; `File.Exists=False` can also mask access failure. Existing watchers already mark the cache dirty and `Refresh` reloads both normal and AutoPilot guidance. Two independent static/source-faithful characterizations PASS; Windows/Visual Studio 2022 compile NOT RUN.
- [ ] K08: External actual .NET tests for each correction plus existing P01/P02/P03/P05 and PDF-P02D regression invariants. **NOT RUN:** this environment has no `dotnet`, `msbuild`, `mono` or `vbnc`; static/Python/Node checks do not close this gate.
- [ ] K08: Actual Windows/VS2022 build of SharedLibrary and affected hosts/projects, recorded with current source and output hashes. **Windows/Visual Studio 2022 compile NOT RUN.**
- [ ] K08: Real Outlook/Word VSTO, Local Chat/error/cancel/persistence, ask_user owner/rendering, ordinary AutoPilot and large-result/context_expand regression runs. **NOT RUN:** Office/VSTO/COM runtime is unavailable in this environment.
- [x] K08: New cumulative changed-only ZIP, full reconstructed source/.inky working basis, hashes, diff and checkpoint-resume documentation. Static/integrity delivery is complete at CP-K08-024; native build/runtime gates above remain open.

The earlier broad “all low-risk fixes completed” status is superseded by these audit findings. Earlier historical entries stay intact. Deferred queue/concurrency/hedging/permission expansion, browser ask_user resume, script_path, dormant policies, forced-tool blocked gates and automatic foreign-writer rollback remain deferred. Existing skill-author-specific generic ToolLoader routing is known architecture debt, not newly declared fixed.

- [x] K07b/F11 source WIP completed by CP-K07B-023: only `SharedLibrary/Code/Agents/AgentResources.vb` changed from CP-K07B-022. New private source states distinguish not configured, not present, empty/whitespace, successful read and read error; ordered direct reads replace `File.Exists` probing; diagnostics are bounded metadata only and do not copy guidance. H24d merge semantics are intentionally preserved via the original `Not IsNullOrEmpty(Content)` merge condition, including whitespace-only mixed-source divider behavior. WIP hash `ef0860c6a295caf27313da5a578769ca059116c1a5d8b410ec38b7685d67a2f7`. QA not yet valid; Windows/Visual Studio 2022 compile NOT RUN.

- [x] CP-K08-024: Static/integrity acceptance and cumulative changed-only delivery. Prepackage static acceptance 91/91 PASS; package overlay/integrity 24/24 PASS; independent package QA2 33/33 PASS. The product ZIP contains exactly 15 changed `.vb` files in repository structure, SHA-256 `c67848355cab2345c64b1054a4587357e819423fdfc1a2336104378de15b0520`. Fresh uploaded H24d basis + ZIP reproduces all 595 physical product files (uploaded extraction excluding `CONTRIBUTING.md` and `Tasks.md`) byte-for-byte; `.inky` remains unchanged. `H24d_to_current_cumulative.diff` is generated from the agreed H24d basis. Because the older `RI_Gen2_260928_1311.zip` archive is not present in this work session, the original-basis history is delivered honestly as a two-stage patch series: the audit-supplied original→H24d diff plus the verified H24d→current diff. This checkpoint is a static/delivery checkpoint only, not a native functional release.
- [x] K07b history clarification: the earlier line labelled “K07b/F11 source WIP” records the deliberately unreleased pre-QA state at that moment. It is superseded by the later released CP-K07B-023 entry; its phrase “QA not yet valid” is historical, not the current K07b status.
- [ ] Native release gates after CP-K08-024 remain: actual VB/.NET correction tests; Windows/Visual Studio 2022 SharedLibrary-first build and dependent-host builds; Outlook/Word VSTO and Office/COM smoke/regression tests; real Local Chat/ask_user/AutoPilot/large-result/context_expand runs.


## Independent post-delivery recheck

- [x] CP-R00-025: Recovered CP-K08-024 from the uploaded H24d basis and its released 15-file overlay; checked all 595 product payloads against the old manifest and independently against both ZIPs. Canonical UTF-8 ZIP names are retained for two XLSX paths; the previous manifest aliases are mapped explicitly, not treated as product changes.
- [x] Documentation clarification: CP-K07B-023 is actually released in WORKLOG/checkpoint/QA and matches AgentResources SHA-256 `ef0860c6a295caf27313da5a578769ca059116c1a5d8b410ec38b7685d67a2f7`. The earlier Tasks WIP paragraph is historical; this explicit release entry supplies the previously missing standalone Tasks record.
- [x] Prior CONTRIBUTING checkpoint/double-check rules preserved unchanged. Product source has not been modified during recovery.
- [ ] Independent semantic recheck of the adjusted areas. Prior structural PASS counts do not establish complete code correctness.
- [ ] Windows/Visual Studio 2022 compile NOT RUN. Native VB/.NET tests remain unavailable; current SDK download attempts fail with DNS errors, not product errors.


- [x] CP-R01A-026: New R-F01/K05a defect verified against the exact Markdig 1.3.2 EscapeInlineParser/LiteralInlineParser source contract. Paths are split into adjacent native literals at escapes; K08 scans each node as a new path, so `C:\Temp\.inky` loses a separator and an already escaped path may visibly gain one.
- [x] Five negative source-mechanics/rendering reproductions and an independent partition-invariance countercheck saved. These use actual CommonMark proxy rendering and pinned native parser semantics, not VB/Markdig execution.
- [ ] R-F01: Coalesce only contiguous native eligible LiteralInline source spans before path scanning; retain structural/literal gaps, deliberate escaping and dialog-owner invariants. Product code not yet changed.


### CP-R01B-027 — post-delivery footer counterexample (QA-scoped checkpoint)

R-F02 reproduced: an inline-code line with >=3 matching backticks, or a backtick opener with a backtick in its info suffix, incorrectly opens the custom fence state and hides the following real TASK_STATUS footer. Markdig 1.3.2 rejects that opener. Four direct cases plus seven delimiter-width counterexamples are saved; tilde-info positive control remains valid. No product changes; R-F01 also remains pending. Native gates remain NOT RUN.


### CP-R02-028 — R-F02 footer opener corrected; QA-scoped release

The opening-fence branch now rejects backticks in backtick info strings, while retaining tilde openings and closing-fence rules. Real footers after inline triple-backtick code are recognized and stripped without losing prose. Fresh external QA1: 109 checks including every former K04 case; QA2: independent CommonMark matrix (5760 cases), old-branch mutation countercheck, exact one-hunk diff and 595-file integrity. No native VB/Markdig/Windows runtime test is claimed. R-F01 remains pending; cumulative/K05b integration QA must be rerun.


### CP-R03A-029 — K05a pre-patch regression constraint

A merge-only R-F01 proposal was rejected before implementation: C:\_folder_\x would lose underscores; literal brackets/backticks could become link/code structure. Five visible-text and four independent DOM-structure counterexamples are saved. This is a demonstrated regression of a hypothetical partial fix, not a new claim of native K08 execution. The intended fix must preserve the original escape in addition to adding the path separator, including scanner stop boundaries. No additional product changes.


### CP-R03-030 — R-F01 ask_user path ranges corrected; QA-scoped release

Only ShowMethods.AskUser.vb changed in this step. Contiguous native literal ranges now share path context, without crossing protected gaps. The original punctuation escape is retained so underscores/brackets/backticks do not become formatting, links or code. Private token writing also receives the following eligible source character at a scanner stop. Fresh QA1 PASS 97/97; QA2 PASS 42/42, including 2,000 independent coverage-bitmap cases and 1726 exhaustive two/three-part splits. Existing K05a visible/structural cases rerun with native escape partitions, plus 12 new boundary cases. Dialog/UI/options/owner code and all other current files remain byte-identical to CP-R03A-029. Native VB/Markdig/Office/Windows gates remain NOT RUN; cumulative release validation is pending.


### CP-R03B-031 — additional K05a newline-boundary defect R-F03

A trailing directory separator before LF/CRLF is a native Backslash-LineBreakInline, not LiteralInline. Current ask_user excludes it and displays C:\Temp\ without its final slash when another line follows. Six direct negative cases plus six end-of-input/newline metamorphic cases verify the gap. No product edit yet. Extend eligibility only to that source backslash marker outside literal HTML; preserve CR/LF and ordinary Markdown breaks. Native execution remains NOT RUN.


### CP-R03C-032 — R-F03 final path separator before LF/CRLF retained

Only ShowMethods.AskUser.vb changed. Native backslash hard-break source markers are eligible only outside literal HTML; CR/LF bytes are never admitted. Full current K05a QA1 PASS 129/129; independent QA2 PASS 59/59, including prior R-F01 tests, newline exclusion and actual CommonMark visible-text counterfactuals. The first generated Python harness failed at syntax parsing before tests; its raw-source escaping was fixed without product changes, failure preserved in R03C_ASKUSER_LINEBREAK/failed_attempt_001. Native VB/Markdig/Windows/Office gates remain NOT RUN.


### CP-R04A-033 — R-F04 structural reference-label defect verified

No source change. Twelve shortcut/collapsed reference link/image cases lose resolution after path escaping; six explicit-reference controls retain their link. Two external negative checks PASS 18/18 each: actual CommonMark target resolution and independently normalized reference-key equality. These are not native Markdig tests. The UNC variant predates this review; R-F01 coalescing additionally exposes drive-path variants. The existing DoubleS path already protects shared reference labels and remains unchanged.


### CP-R04-034 — R-F04 shared-reference identities protected

Only ShowMethods.AskUser.vb changed. Native shortcut/collapsed link and image spans are structural and remain source-identical, including their original renderer display; explicit separate-reference labels and ordinary paths remain eligible. QA1 PASS 275/275; QA2 PASS 96/96. Fresh suite includes all R-F01/R-F03 cases, 36 reference contexts, link/alt/definition fidelity and guard-removal/target-redirection mutations. One QA harness attempt failed because markdown-it-py 4.2.0 omits text_special image-alt tokens; failure saved, gold expectations retained, independent Mistune 3.2.1 used for reference DOM checks in agreement with pinned native renderer source. No product workaround for the test renderer. Native build/runtime gates remain NOT RUN.


### CP-R05A-035 — R-F05 stale JavaScript member context verified

No source change. Decimal literals (1.5/.5), optional computed member access, and empty spread followed by return /process.env/ wrongly trigger Node API rejection in the current lexical model. Independent Node vm executes all four safely with no Node globals. Two actual member-keyword division cases still fail with ReferenceError as expected. QA1 and QA2 each PASS 6/6 as negative characterization, not product/native approval.


### CP-R05-036 — R-F05 JavaScript token-context lifetime corrected

Only JsRunTool.vb changed. One narrow state-reset insertion prevents numeric/computed/spread tokens from keeping a pending property-name flag alive. Whitespace/comments before actual property identifiers still work. Full K06a QA1 PASS 187/187 (169 lexical cases); QA2 PASS 167/167, including 26 prior and 124 additional actual isolated Node-vm cases. Native VB/WebView2/Office remains NOT RUN. Exact insertion removal reproduces CP-R05A-035; every other product file, sandbox and dispatch remain unchanged. Current review changes three existing VB files relative to CP-K08-024.


### CP-R06H-037 — External QA provenance repaired

No product changes. Attempt001 incorrectly copied a stale K02_2 QA2 PASS after an aborted fresh test: explicitly invalidated and retained. Attempt002 progress reporting raised NameError; retained and corrected. Attempt003 isolates outputs, requires unique fresh JSON and exit0/PASS. QA1 9/9, independent QA2 23/23 PASS. Historical assertion failures are preserved, not suppressed and not yet current acceptance. Native gates unchanged.


### CP-R06A-038 — Central write authority: verified no-change

Canonical configured-central denial precedes every successful Resolve return; no opt-in widening of workspace/strict allow gates. 512 abstract path/authority combinations plus independent segment-boundary cases. Native reparse/8.3/ACL tests remain open.

Current-scope QA1 7 checks and independent QA2 9 checks PASS. No product change, native gates NOT RUN. External adjudication harness attempt001 failed before tests because one historical report encoded checks as name:boolean rather than name:object. Preserved attempt001_BOOL_REPORT.py; strict normalization now accepts those two explicit report shapes; no product change or product failure.


### CP-R06B-039 — Preflight/normalization/recovery: verified no-change

Current normalizer/validator/recovery semantic tests pass. Direct preflight and recovery shared-block parity verified. The historical all-host-diff asymmetry is exactly later diagnostic labels on pre-existing Outlook power-transition and Word cancellation branches; not divergent recovery.

Current-scope QA1 6 checks and independent QA2 6 checks PASS. Fresh original assertion PASS count: 146; superseded historical-scope assertions: 7, individually retained and adjudicated in R06_SCOPE/HISTORICAL_ADJUDICATION.json. No product change, native gates NOT RUN.


### CP-R06C-040 — Current-model tool exposure: verified no-change

Both exposure and tool-loader methods remain byte-identical across Word/Outlook. Fresh prior converter/permission/model-switch tests pass. Old whole-file non-target equality predates K07a audit changes, not a loader defect.

Current-scope QA1 5 checks and independent QA2 3 checks PASS. Fresh original assertion PASS count: 70; superseded historical-scope assertions: 4, individually retained and adjudicated in R06_SCOPE/HISTORICAL_ADJUDICATION.json. No product change, native gates NOT RUN.


### CP-R06D-041 — DoubleS protected source handling: verified no-change

All original semantic/renderer tests pass; actual DoubleS implementation unchanged since K08. Dependencies explicitly rebound to released R-F02 footer and R-F01/R-F03/R-F04 AskUser hashes; old dependency-hash failures retained as superseded, not silently passed.

Current-scope QA1 4 checks and independent QA2 5 checks PASS. Fresh original assertion PASS count: 369; superseded historical-scope assertions: 2, individually retained and adjudicated in R06_SCOPE/HISTORICAL_ADJUDICATION.json. No product change, native gates NOT RUN.


### CP-R06E-042 — Authored frontmatter grammar diagnostic: verified no-change

Actual-source grammar cases pass. Helper-local diff remains exactly one fully qualified Regex line; parser, descriptor selection, validation and markdown-file selection methods unchanged. Whole-file hash changes are exclusively the later K07b scope.

Current-scope QA1 4 checks and independent QA2 6 checks PASS. Fresh original assertion PASS count: 54; superseded historical-scope assertions: 7, individually retained and adjudicated in R06_SCOPE/HISTORICAL_ADJUDICATION.json. No product change, native gates NOT RUN.


### CP-R06F-043 — Bounded audit events/outcomes: verified no-change

Fresh QA1 140 and independent QA2 140 original checks pass. Current audit helper bodies identical across hosts; shared builder and both host files byte-identical K08. Real logger/Office integration still native-gated.

Current-scope QA1 4 checks and independent QA2 4 checks PASS. Fresh original assertion PASS count: 280; superseded historical-scope assertions: 0, individually retained and adjudicated in R06_SCOPE/HISTORICAL_ADJUDICATION.json. No product change, native gates NOT RUN.


### CP-R06G-044 — AutoPilot guidance fallback/observability: verified no-change

Fresh QA1 174 and independent QA2 141 checks pass, including 729 read-probe sequences. Global ReadInkyMd and prompt-builder code unchanged; actual Windows access-denied/race/cache/UI flows remain native-gated.

Current-scope QA1 4 checks and independent QA2 3 checks PASS. Fresh original assertion PASS count: 315; superseded historical-scope assertions: 0, individually retained and adjudicated in R06_SCOPE/HISTORICAL_ADJUDICATION.json. No product change, native gates NOT RUN.


### CP-R07-045 — Full cumulative recheck acceptance

Five real residual defects (R-F01..R-F05) are corrected in exactly three existing files: AskUser, TaskStatusFooterParser and JsRunTool. Fresh full footer QA 109/109 and independent9/9 (5760 CommonMark combinations); full AskUser QA275/275 and independent96/96; full JS QA187/187 and independent167/167, including26 prior+124 new actual Node-vm cases. All are external and NOT VB/Markdig/WebView2/Office execution. Remaining-area original assertions:1234 PASS,20 explicitly superseded historical scope/hash assumptions retained with independent current-scope replacement proofs; not counted as PASS. Final whole-tree/source QA45/45 and independent12/12. All595 products accounted for, .inky unchanged, same15 cumulative VB paths, no WIP. Windows/Visual Studio 2022 compile NOT RUN. Packaging and transport QA next.


### CP-R08-046 — Source delivery checkpoints and two independent reconstructions

Cumulative H24d source ZIP contains exactly15 genuinely changed VB files; incremental CP-K08-024 ZIP contains exactly3. Both preserve repository paths and exclude tests/governance/unchanged dependencies. Source/archive QA1 15/15 PASS. Independent CLI unzip+sha256sum+byte comparison QA2 16/16 PASS; all595 product payloads match for both overlay paths. Original uploaded ZIP has two known legacy filename-decoding aliases: normalized only inside temporary QA extraction by unique same-directory SHA; no product or .inky rename. Attempt001 FileNotFoundError retained as harness failure. Internal complete current product/.inky snapshot verified; not delivered as a full source ZIP. Native gates unchanged. Final evidence seal: recheck_final/DELIVERY_VERIFICATION.json must be PASS and hash-consistent.


## Feature stream — tracked paragraph insertion (`word_markup`)

- [x] CP-F00-047: Reconstruct the latest released CP-R08-046 source basis and start the feature stream with no product-code change. Mandatory project/audit contracts and the 2026-09-29 feature request were reread; native Windows/VS2022/VSTO gates remain open.
- [x] Verify the feature request against the complete current `SharedLibrary/Code/Agents/WordTools.vb` before changing source (CP-F01A-048).
- [x] CP-F01A-048: Fully reread current `WordTools.vb` (2895 lines) and reproduce the feature gap without source changes: no anchored sibling-paragraph ops, append uses direct body append, `AppendMarkdownContent` has zero call sites, and paragraph-mark insertion is not tracked. Existing inline edit/recovery/artifact/path contracts captured as regression invariants.
- [x] Add `insert_paragraph_after` / `insert_paragraph_before` as a shared `word_markup`/`word_write` capability with tracked paragraph marks only in markup mode (CP-F01C-050).
- [x] CP-F01B-049: Reconcile all mandatory contracts immediately before the first feature source patch and freeze the minimal semantics: shared word_write/word_markup paragraph ops, same-story insertion, per-line Markdown precedence, first-paragraph heading/style/numbering options, tracked paragraph mark only in markup mode. No product change.
- [x] CP-F01C-050: Implement `insert_paragraph_before` / `insert_paragraph_after` in shared `WordTools.vb` for both `word_markup` and `word_write`. Every text line becomes a same-story sibling paragraph; existing Markdown block style wins, otherwise first-paragraph `heading_level` then `style`; `inherit_numbering` copies anchor numbering to the first inserted paragraph only. Markup mode tracks both inserted text and the new paragraph mark via the existing revision-id allocator. QA1 45/45, independent QA2 52/52, tree QA 19/19 PASS. One QA1 harness-position assumption was recorded and rerun; no source change resulted.
- [x] Separately correct body `append` placement before final `w:sectPr`, track appended paragraph marks, and verify `word_verify_revisions` compatibility with inserted paragraph marks (CP-F02B-052).
- [x] CP-F02A-051: Negative-characterize the remaining feature slice with no product change. `append` still uses `Body.AppendChild`, can follow a final `w:sectPr`, and does not track its paragraph mark; `word_verify_revisions` ignores `w:pPr/w:rPr/w:ins`, so rejecting a feature-created paragraph can leave an extra paragraph boundary. Characterization 11/11 PASS.
- [x] Correct `append` placement/tracking and inserted-paragraph-mark verification without broadening existing deleted-paragraph support (CP-F02B-052).
- [x] CP-F02B-052: Correct `append` and inserted-paragraph-mark verification in shared `WordTools.vb`. `append` now stays before body `w:sectPr`, tracks each appended paragraph mark only in markup mode, and the verifier counts/accepts/rejects `w:pPr/w:rPr/w:ins` while keeping deleted paragraph marks unsupported. QA1 49/49 PASS; independent QA2 52/52 PASS; F01C protection 21/21 PASS; tree QA 18/18 PASS. Release postcheck pending final released-state verification.
- [x] Build the final changed-only feature ZIP, cumulative diff/manifest and evidence after checkpoint postcheck (CP-F04A-054, CP-F05A-057); native Windows/VS2022/VSTO gates remain open.
- [x] CP-F03A-053: Final feature-request acceptance with no product change. All 43 requirement/contract checks PASS against released CP-F02B-052, including schema exposure, per-line paragraph creation, Markdown heading mapping, style/heading/numbering options, tracked paragraph mark metadata, operation_id behavior, append-before-sectPr and verifier coherence. The request-named `RI_Gen2_260929_0631.zip` was not supplied; the described gap had been independently reproduced on the actual released working basis before implementation.

- [x] CP-F04A-054: Build deterministic feature delivery artifacts with no product-source change. Incremental ZIP contains only `SharedLibrary/Code/Agents/WordTools.vb` for CP-R08-046/recheck users; cumulative changed-only ZIP contains the prior 15 released corrections plus current `WordTools.vb` for `RI_Gen2_260928_2021.zip`. QA1 22/22 PASS, including both 595-file overlay reconstructions. First QA1 attempt failed only on a manifest-shape comparison; product/source unchanged.
- [x] Run separately implemented package/evidence QA2 and seal current evidence (CP-F04C-056, CP-F05A-057).

## CP-F04B-055 — independent feature package QA2

- [x] Independently verify the F04A source archives without rebuilding product source. QA2 25/25 PASS: exact member sets/payloads, path-safety, current 595-file manifest, exactly 16 changed source files vs `RI_Gen2_260928_2021.zip`, only `WordTools.vb` changed vs CP-R08-046, deterministic ZIP metadata, BOM/CRLF and truthful native gates.
- [x] Build the final evidence ZIP and external preseal hashes (CP-F05A-057); final delivery seal is F05B.
- [ ] Windows/Visual Studio 2022 compile NOT RUN.
- [ ] Office/VSTO runtime NOT RUN.

## CP-F04C-056 — feature evidence package

- [x] Build deterministic feature Evidence ZIP from CP-F04B-055 evidence only. Evidence QA 15/15 PASS; 69 members; no `.vb` members, no source ZIPs and no repository product-source root members. SHA-256 `df0e7af932257a856589bcb8da96d33524cc13885bbf96998a5d22a234728e20`.
- [x] Create external SHA256SUMS/final delivery report and run a separately implemented final seal (CP-F05B-058 candidate; release only after PASS).
- [ ] Windows/Visual Studio 2022 compile NOT RUN.
- [ ] Office/VSTO runtime NOT RUN.

## CP-F05A-057 — current evidence bundle

- [x] Build a fresh feature evidence ZIP from CP-F04C-056 state; do not reuse the earlier F03B evidence archive. Evidence contains the feature request, mandatory audit contracts, all feature checkpoint JSONs, feature QA history, current governance and delivery manifests/diffs, but no nested ZIP and no product `.vb` source. QA1 15/15 PASS; independent CLI QA2 12/12 PASS; evidence ZIP SHA-256 `46b2c4e9374333cc526160b60b0626b6a3ca9ab3b6bbabe10995cacdb4c675b8`.
- [x] CP-F05B-058 released after pre-release 44/44 PASS, corrected independent final seal 59/59 PASS and corrected candidate validation 34/34 PASS. Harness-only failures are retained; no product/package change resulted.
- [ ] Windows/Visual Studio 2022 compile NOT RUN.
- [ ] Office/VSTO runtime NOT RUN.


## CP-F05B-058 — final feature delivery seal

- [x] Freeze product source at `SharedLibrary/Code/Agents/WordTools.vb` SHA-256 `709f26517c2656359e3c5fe3d0daa99286b54ae69fe72613cecb30ae804db4b7`; no product change after CP-F02B-052.
- [x] Final pre-release seal 44/44 PASS against all 595 product files, authoritative F04A/F04C package QA and F05A evidence QA.
- [x] F05B attempt001 recorded as harness failure only: 56/57 checks; `/base` was incorrectly assumed to include the 15 CP-R08/recheck overlay files. Product/source/ZIP/Evidence checks all passed; no product change.
- [x] Corrected F05B final seal 59/59 PASS using the verified three-layer overlay chain: base + released 15-file recheck ZIP + one-file feature ZIP. All 595 product payloads reconstruct exactly.
- [x] Corrected candidate validation 34/34 PASS with line-based Unified-Diff headers; CP-F05B-058 can release.
- [x] Deliver one-file incremental feature ZIP for CP-R08-046/recheck users and 16-file cumulative changed-only ZIP for `RI_Gen2_260928_2021.zip`; both source archives retain their independently verified hashes.
- [x] Deliver current evidence ZIP, cumulative diffs, product/delivery manifests, final governance/recovery documents and SHA-256 list.
- [ ] Windows/Visual Studio 2022 compile NOT RUN.
- [ ] Office/VSTO runtime NOT RUN.


## CP-F05C-059 — final postcheck closure

- [x] Preserve CP-F05B-058 source/packages unchanged and record the independent release postcheck: 47/47 PASS.
- [x] No product-source change; all 595 product payloads remain exact, `WordTools.vb` remains SHA-256 `709f26517c2656359e3c5fe3d0daa99286b54ae69fe72613cecb30ae804db4b7`.
- [x] Recovery state advanced to this no-change closure checkpoint; final standalone governance and hashes regenerated from the released bytes.
- [ ] Windows/Visual Studio 2022 compile NOT RUN.
- [ ] Office/VSTO runtime NOT RUN.

## Semantic Archive — current status (2026-10-04)

- [x] Add Semantic Archive administration, source roots, catalog discovery, search/read tools and Word/Outlook integration.
- [x] Keep all externally configurable Semantic Archive defaults in `SharedMethods.Constants.vb` and wire global configuration through the existing Settings/ConfigWizard contracts.
- [x] Add default-on per-source office/image extension filtering and always exclude Knowledge Store `.redink` output plus Semantic Archive generated output from ingestion.
- [x] Reuse the existing `text_export_to_text` extraction path and preserve native text before selective OCR.
- [x] Add shared reusable per-document artifacts, private fallback, current source-permission rechecks and background permission reconciliation.
- [x] Consolidate Semantic Archive generated state under `sa-archives` and remove the earlier separate work/lock/staging top-level layout.
- [x] Make direct local archive discovery available to the model and add `semantic_archive_list`.
- [x] Fix PDF text-layer coverage so verified native extraction can become searchable.
- [x] Add selective OCR and configurable bounded OCR page batching while retaining per-page coverage verification.
- [x] Keep logical document identity independent of OCR/extraction policy and prune retired/tombstone records from the active generation so repeated refresh, reindex and re-extract cannot multiply source records.
- [x] Allow OCR batches from 1 to 75 pages, mark OCR batch-size edits dirty so Save is enabled, and keep the default at 16.
- [x] Simplify the normal Semantic Archive diagnostics/progress view while retaining raw queue, generation, routing, permission and identity details behind Show technical details.
- [x] Intercept interactive local `file://` source links in the Outlook web UI before browser navigation and open them through the existing host path handler.
- [x] Host automatic Semantic Archive content and permission maintenance in Outlook only; Word retains manual Semantic Archive commands and independent Knowledge Store maintenance.
- [x] Remove the obsolete `SemanticArchiveDerivedOutputRoot` / “Legacy private folder” compatibility path; private derivatives now use `SemanticArchiveShadowArtifactRoot` or the per-user LocalAppData fallback.
- [x] Add a unique high-confidence exact-metadata fast path so obvious archive hits can skip semantic hierarchy model routing and proceed directly to exact evidence reading.
- [x] Show a visible “Querying Semantic Archive…” progress indicator in Word and Outlook Freestyle when `(sa)` is requested.
- [x] Separate concise user diagnostics from optional technical diagnostics.
- [x] Provide original-source references for Semantic Archive evidence; unattended delivery must not rely on a local user link.
- [x] Remove feature-specific test/QA/report artifacts from the product repository; temporary verification stays external.
- [ ] Run a clean Windows/Visual Studio 2022 build of SharedLibrary, Word, Outlook, Excel and optional Semantic Archive Worker after the latest OCR batching changes.
- [ ] Run Word and Outlook smoke tests for Semantic Archive administration, refresh, search/read, source links and dialog-owner behavior.
- [ ] Verify selective OCR batching with representative native-text, mixed text/scan and full-scan PDFs; confirm all requested OCR pages are accounted for before marking extraction complete.
- [ ] Verify a second Windows user reuses shared artifacts without repeating extraction/indexer work and cannot retrieve sources they cannot currently read.
- [ ] Verify UNC archive/source/shared-artifact paths and path-length handling on the target SMB environment.
- [ ] Verify AutoPilot attaches only original documents actually read as evidence and does not expose unusable local/UNC links.


## Semantic Archive — local catalog / Freestyle sources / search strategy (2026-10-04)

- [x] Diagnose the supplied First/Second logs: search took 120.081/32.769 seconds with 15/17 routing model calls; the earlier exact-metadata shortcut did not execute. The first run stopped at the search deadline. The logs do not identify the individual slow routing request or prove a specific provider fault.
- [x] Rename the private catalog setting to `SemanticArchiveCatalogPathLocal` and context property to `INI_SemanticArchiveCatalogPathLocal` across ConfigWizard, Settings/load/save/reset/export, Word, Outlook, Excel, Worker and shared consumers. The directory value and existing index data stay unchanged. Report old-development-key use without introducing an alias or silent migration.
- [x] Route SA filesystem paths through the existing shared environment/Red Ink placeholder helper; retain containment, Windows limits and permission checks. Include source input, derivative paths, exclusions, source URIs and worker log arguments.
- [x] Add a cached Sources menu in the existing Word/Outlook Freestyle footer. Show SA first and KB second, descriptions as tooltips, insert explicit source triggers, keep form height unchanged. Discovery is background-only with independent single-flight caches and read-only catalog APIs; it creates no registry entries/directories and does not enumerate original files or call a model.
- [x] Add bounded direct document-card ranking for at most 64 source records using the existing semantic selector, with no keyword eligibility filter. Keep large-archive hierarchy navigation, complete-record prompt limits, current authorization and exact read contracts. Retain hierarchy continuation after a metadata-only shortcut instead of claiming the hierarchy was traversed.
- [x] Implement the central descriptor library and managed subscriptions in source; see the subscription section below for behavior and remaining native validation.
- [ ] Measure both new search routes on Windows with the user's 42-document generation and compare `Coverage.RetrievalStrategy` / `ModelCalls` / `ElapsedMilliseconds`. Do not infer measured speed gains from static checks.
- [ ] Native VS2022/VB/VSTO build, Outlook/Word Freestyle layout at different DPI settings, real `%DESKTOP%`/`%DOCUMENTS%`/UNC behavior and live ACL-change tests remain unexecuted in this environment.
- [ ] Deployment: rename existing INI `SemanticArchiveCatalogPath` (or `SemanticIndexArchivePath`) to `SemanticArchiveCatalogPathLocal`, retaining the same directory. Rebuild SharedLibrary and hosts. No index deletion/re-extraction is required for this rename or query/UI changes.

## Semantic Archive — library subscriptions and Freestyle scope editing (2026-10-04)

- [x] Confirm the supplied latest Local Agent log: flat document-card search used 5 model calls and 9.650 seconds internally (9.721 seconds tool total); exact read used 195 ms tool time. This is a measured prior-patch result, not a benchmark of this patch. No Freestyle log proves the duplicate-query hypothesis.
- [x] Add `SemanticArchiveCatalogLibraryPath`, default empty, through ConfigWizard, Settings/read/write/reset/export, SharedContext and Word/Outlook/Excel property/setting descriptions. Retain centrally provisioned library configuration on ordinary settings reinitialization, like the local catalog location; full reset uses the empty constant.
- [x] Add local-author Publish / Update library, Withdraw from library and Sync library actions in the existing admin console. Only local definitions can publish; subscribers can opt out/re-enable and refresh content, not edit centrally managed source/processing definitions. Library synchronization remains accessible even with no local archive yet.
- [x] Store central per-archive definitions with publisher ownership, bounded strict JSON, inherited reader policy, protected writer ACL, per-entry exclusive lock and atomic replacement. Preserve existing descriptor audience on update. Reject nonportable local/mapped roots when publishing to a UNC library. Original files are not copied and retain their own rights.
- [x] Implement automatic revisioned local subscriptions with stable IDs, opt-out, withdrawal suspension, recovery after restored access, anti-rollback/publisher checks and local catalog revision conflict handling. Identical publish retries do not add revisions. Failed local bookkeeping after a central commit is explicit and recoverable by identical retry. Private authored archives stay separate.
- [x] Gate every operational SA path on the local setting. Library-only configuration leaves SA unavailable; explicit disabled requests return a configuration result rather than reading/indexing data. Hide SA source-menu/ribbon choices when local configuration is absent.
- [x] Synchronize library metadata in the background at configuration/source-overview warm-up and periodically in the existing Outlook maintenance coordinator (300-second library interval). Subscription content and permission maintenance require no per-user enable switch, but still respect source authorization, configured windows, idle cancellation and the existing writer locks. Word does not acquire automatic content indexing.
- [x] Check current central authority for subscribed search/read/list/build and evidence disclosure; preserve the existing current-original ACL/hash checks and independent unattended requester authorization. Offline/denied/unverified central definitions are not usable merely because a private index exists.
- [x] Make source selection replace scope-only SA/KB controls, prefer unique readable SA names, deduplicate case/spacing-equivalent source controls and make Ctrl+P restoration idempotent without deleting unrelated task text or explicit source queries. Deduplicate equivalent resolved inline SA requests as a second boundary.
- [x] Preserve prior extraction/OCR batching, stable source identity, shared artifact validation, source links/AutoPilot delivery and the flat-card search strategy. No existing source/index directory needs deleting or re-extracting for this patch. First-time subscribers still need their initial rights-filtered private projection.
- [ ] Perform clean VS2022/VB/VSTO builds for SharedLibrary, Word, Outlook, Excel and optional Worker. No VB compiler or native Windows/Office runtime was available during this implementation.
- [ ] Native acceptance: configure only Library vs only Local vs both; cold-start a new subscriber without a catalog; publish/update/withdraw/re-publish; test subscriber opt-out/re-enable, definition write/owner/reader ACLs, UNC shares and failed/offline/interrupted atomic commits. Confirm same entry ID and no duplicate records across repeated sync/refresh.
- [ ] Verify real multi-user source ACL removal/restoration, two concurrent publishers, simultaneous subscriber synchronization and artifact reuse; confirm private archives never publish automatically. The central library must be pre-provisioned: publishers may create files but must not be able to replace other entries via directory/ancestor delete/ACL rights.
- [ ] Native Freestyle acceptance in Word and Outlook: existing `(sa)` then select a named archive; Ctrl+P once/repeatedly/reopen; equivalent KB scope; distinct explicit queries; form height/owner/cancel behavior and empty-library handling. Normal users need Outlook running and eligible for automatic content/permission work; initial preparation is not instantaneous.
