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

