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

## Semantic Archive — scalable routing / index-only reindex / PDF fallback (2026-10-04)

### CP-SA-SR00 — implementation basis and architecture freeze

- [x] Input checkpoint: supplied `RI_Gen2_261004_2026.zip`; reconciled with current `CONTRIBUTING.md`, cumulative `Tasks.md`, the scalable-routing concept and the prior Semantic Archive handoff.
- [x] Step: no product-code change. Fully re-audited the current Semantic Archive builder/work queue, generation/schema/storage hierarchy, search/read contracts, Admin maintenance commands and the now-present `SemanticArchive.Worker` command surface before implementation.
- [x] Root cause/design finding: the existing `SemanticArchiveHierarchy` is a deterministic range/shard persistence tree and is also used for generation/storage locality. Converting it directly into a multi-parent semantic DAG would couple routing changes to document-shard identity and publication. The implementation therefore keeps the range/shard hierarchy as storage and adds an independent semantic routing DAG over document cards.
- [x] Root cause/design finding: current `reindex` only maps to `RebuildSemanticMetadata=True`; `ProcessDocumentAsync` can still call `TextExportService.ExportFileAsync` when an extraction signature changes or a reusable representation cannot be validated. A distinct index-only operation contract is required.
- [x] Root cause/design finding: large-document read already resolves the per-document semantic section index and selected byte ranges; this is a preservation/telemetry task rather than a replacement read architecture.
- [x] Root cause/design finding: deterministic malformed-PDF parser failures belong in the shared extraction/PDF layer, not in Semantic Archive-specific conversion logic.
- [x] Regression invariants: stable logical document identity; original-source authorization at metadata disclosure and evidence read; search/read separation; exact metadata remains parallel and never a semantic eligibility gate; incomplete/unknown extraction remains excluded unless explicitly configured; active generations contain current records only; large indexed documents do not fall back to full-text reads; SA paths keep the shared path expansion/guard; no provider/model/organization-specific branches.
- [x] Checks: source-tree presence of Worker verified; current `reindex`, builder extraction reuse, hierarchy/search/read and Admin mappings inspected; no source modification before this checkpoint.
- [ ] Open/native gates: Windows/Visual Studio/VSTO/Office runtime not available in this environment and must not be claimed.
- [ ] Next: implement a hard index-only build operation and structured `requires extraction` outcomes, then verify it independently before routing changes.

### CP-SA-SR01 — hard index-only reindex contract

- [x] Input checkpoint: CP-SA-SR00.
- [x] Performed step: added an explicit `IndexOnlyRebuild` build mode and wired both Admin `SemanticReindex` and Worker `reindex` to it. Reindex now operates on the currently published logical document inventory instead of crawling for new files; `Refresh archive` remains the discovery operation.
- [x] Affected files: `SemanticArchiveWorkQueue.vb`, `SemanticArchiveBuilder.Scan.vb`, `SemanticArchiveBuilder.vb`, `SemanticArchiveBuilder.Diagnostics.vb`, `SemanticArchiveForm.vb`, `SemanticArchive.Worker/Program.vb`, `Tasks.md`.
- [x] Root cause: `RebuildSemanticMetadata=True` previously forced semantic work but did not prohibit extraction. A changed extraction signature or invalid/missing extract could fall through to `TextExportService.ExportFileAsync`.
- [x] Applied change: invalid/missing/incompatible extracts in index-only mode now produce inactive `needs_extraction` records with explicit `requires_extraction` diagnostics and a counter; no extraction/OCR call is made. Valid extracts are integrity/source-hash validated and reused. Section indexes/cards are rebuilt normally from those extracts and counted separately.
- [x] Regression invariants: Refresh/extract/retry/permissions behavior remains on the existing paths; logical document IDs are unchanged; no new extraction pipeline; reindex does not discover additions/removals; source identity/hash is still checked before derivative reuse.
- [x] QA1: 9/9 static contract checks PASS (worker/admin wiring, inventory binding, semantic-force behavior, guard ordering, explicit status/counters).
- [x] QA2: independent ProcessDocument control-flow check 6/6 PASS proving the index-only invalid-extract return occurs before the sole `TextExportService.ExportFileAsync` call while normal extraction remains available to non-reindex operations.
- [x] Toolchain check: no `msbuild`, `dotnet` or `vbc` executable was available in this environment; no Windows/VS/VSTO compile is claimed.
- [ ] Next: add a separate semantic routing DAG representation/builder over document cards, with multi-parent references, bounded semantic partitioning/splitting and migration without invalidating stored extracts.

### CP-SA-SR02 — additive semantic routing DAG and incremental card reuse

- [x] Input checkpoint: CP-SA-SR01 on the supplied `RI_Gen2_261004_2026.zip` basis.
- [x] Performed step: added a semantic routing DAG as a separate generation artifact over the existing deterministic range/shard persistence hierarchy. The storage tree remains authoritative for generation locality; routing semantics no longer overload filesystem/range structure.
- [x] Affected files: `SemanticArchive.Schema.vb`, new `SemanticArchiveRouting.vb`, `SemanticArchive.Validation.vb`, `SemanticArchiveBuilder.vb`, `SharedLibrary.vbproj`, `Tasks.md`.
- [x] Root cause: the existing hierarchy has one structural parent and deterministic key-range splitting, so converting it directly into a multi-parent semantic graph would couple retrieval semantics to storage identity/publication and risk idempotency regressions.
- [x] Applied change: two independent semantic route families (`content` and `intent`) are built from model-generated document-card semantics. A logical document can therefore have multiple semantic parents without duplicating the document. Groups are bounded by the configured maximum children; semantic SimHash partitioning is derived from title/summary/topics/intents/identifiers/exact terms. Pathological unsplittable collisions switch immediately to a recursively bounded stable-identity fallback rather than producing a long one-child chain.
- [x] Applied change: routing groups carry stable IDs, content/membership signatures, representatives and checksummed immutable generation storage at `routing/graph.json`. Unchanged group cards are reused by signature; only affected groups/ancestors have their cards recomposed. Existing generations with no routing artifact remain readable, so migration can occur by index-only reindex without invalidating extracted text.
- [x] Applied change: representative selection covers semantic child branches before filling remaining slots, avoiding document-ID ordering accidentally hiding a small branch from an ancestor routing card.
- [x] Regression invariants: generation schema version remains unchanged; existing generations are valid without a routing graph; source/document identity remains independent of routing/profile policy; no source extraction occurs in the routing builder; storage range hierarchy remains intact.
- [x] Immediate verification: graph roots/edges/document-parent mappings, cycle/unreachable checks, child bounds, representative bounds, optional-artifact validation and project compile inclusion checked statically. Synthetic topology demonstrates bounded 50,000-document depth (balanced depth 3; complete-collision fallback depth 2 at max child 48).
- [ ] Native compiler/runtime validation remains open; no Windows/VS2022 build was available.

### CP-SA-SR03 — batched semantic beam routing, continuation and large-document telemetry

- [x] Input checkpoint: CP-SA-SR02.
- [x] Performed step: integrated the semantic routing DAG into `SemanticArchiveSearchService` while retaining the existing flat-document-card strategy for small catalogs and the legacy hierarchy as a migration fallback for generations without a routing graph.
- [x] Affected files: `SemanticArchiveSearchService.vb`, `SemanticArchiveSearchService.Read.vb`, `Tasks.md`.
- [x] Root cause: the prior hierarchy could approach one model call per node, and the exact-metadata first-page shortcut could return before semantic routing. Neither satisfies scalable semantic recall.
- [x] Applied change: sibling routing groups at the same level are classified in bounded batches through the existing provider/model-agnostic semantic selector. Selected groups expand at high priority; evaluated non-selected groups remain as low-priority widening frontier and can expand later without another model call. Document leaves are also selected in bounded semantic batches.
- [x] Applied change: exact metadata remains an independent deterministic candidate channel but its first-page short-circuit is disabled. It can add candidates and never determines semantic eligibility.
- [x] Applied change: compact routing entries preserve title, summary, topics, answerable intents, identifiers, exact terms and bounded semantic facets so ~50-document catalogs and sibling batches fit substantially fewer prompt groups without a lexical pre-filter.
- [x] Applied change: continuation retains the original query, pinned archive generations, visited routing groups/leaves, returned/candidate documents, budgets and unexplored routing queues. Coverage now records routing levels, groups considered, model calls, returned candidate count, continuation availability, section-index queries and evidence bytes.
- [x] ACL/security: persisted aggregate group cards are not blindly exposed to a model. Query-time routing entries are recomposed from currently authorized representative documents; if none can be safely exposed, a permission-neutral navigable branch is used. Source access is still rechecked again before hit metadata disclosure and evidence read.
- [x] Large-document invariant: indexed reads remain section-index-only with no full-extract fallback; `SectionIndexesQueried` telemetry was added without changing that read contract.
- [x] Regression invariants: small-catalog flat semantic selection remains available; old generations still use the existing hierarchy; exact metadata is parallel; current source ACL checks remain before metadata disclosure and evidence; continuation never turns a partial semantic beam into an exhaustive-negative claim.
- [x] Immediate verification: QA assertions confirmed DAG preference only above the small-catalog threshold, batched sibling selection, widening retention, ACL-safe routing cards, continuation queues, exact-channel non-gating and strict indexed reads.

### CP-SA-SR04 — deterministic malformed-PDF parser fallback and retry classification

- [x] Input checkpoint: CP-SA-SR03.
- [x] Performed step: fixed reproducible malformed-font `/FirstChar` PDF failures in the shared PDF extraction layer rather than adding an SA-specific converter.
- [x] Affected files: `SharedMethods.FileImporter.vb`, `SemanticArchiveWorkQueue.vb`, `SemanticArchiveBuilder.vb`, `Tasks.md`.
- [x] Root cause: PdfPig deterministic malformed TrueType font-dictionary failures previously surfaced as generic `pdf_read_failed`, so ordinary retry policy could repeat the identical parser path without any chance of recovery.
- [x] Applied change: `ReadPdfAsTextEx` classifies the known deterministic `/FirstChar` failure through the exception chain. The catch only records classification; the asynchronous fallback runs after `End Try`, complying with the project rule that VB must never `Await` inside `Catch`.
- [x] Applied change: the fallback uses existing PdfSharp page discovery and the existing `PerformSelectivePdfOcr` implementation for every page in bounded OCR batches. It marks extraction complete only when all discovered pages have verified processed ranges; otherwise completeness remains false with an explicit error code/provenance basis.
- [x] Applied change: when OCR is disabled/unavailable or the independent fallback cannot determine pages, the builder raises a deterministic explicit-retry processing classification. The generic queue stores its next attempt at `DateTime.MaxValue`; normal maintenance therefore does not retry the identical failure indefinitely, while explicit `Retry failed` still resets the item to due.
- [x] Regression invariants: normal PDF text-layer extraction and selective OCR remain unchanged; the shared `text_export_to_text` path remains the only SA extraction path; page coverage remains strict; no deterministic parser error is silently treated as complete.
- [x] Failed patch attempt retained: the first local source-write attempt failed with a filesystem `PermissionError` before changing bytes. The working tree was made writable and the same targeted correction was then applied successfully; this was a tooling failure, not a product defect.
- [x] QA finding/correction: the first implementation invoked the async fallback from a `Catch`; rereading `CONTRIBUTING.md` exposed the explicit prohibition. It was corrected before release and all relevant QA was restarted from zero.

### CP-SA-SR05 — Admin/Worker operation surface, status taxonomy and reindex telemetry

- [x] Input checkpoint: CP-SA-SR04.
- [x] Performed step: completed the requested Admin/Worker operation semantics and progress/counter surface.
- [x] Affected files: `SemanticArchiveForm.vb`, `SemanticArchiveBuilder.Diagnostics.vb`, `SemanticArchive.Worker/Program.vb`, `SemanticArchive.Worker/WORKER_HELP.txt`, `Tasks.md`.
- [x] Admin now exposes `Refresh archive`, `Rebuild semantic index`, `Re-extract/OCR selected` and `Retry failed`; selected semantic reindex is separately labeled. Help explicitly states that semantic reindex never invokes extraction/OCR and reports incompatible/missing extracts as requiring extraction.
- [x] Maintenance rows distinguish `Needs extraction`, `Needs semantic indexing`, `Failed extraction`, `Failed indexing`, `Excluded because incomplete/unknown`, `Ready/searchable`, plus restricted/removed states without disclosing unauthorized cached metadata.
- [x] Worker `reindex "Archive Name"` is the canonical index-only operation. JSON batch output includes extracts reused, cards rebuilt, section indexes rebuilt, routing groups built and documents requiring extraction; a required extraction makes the worker result incomplete instead of silently extracting.
- [x] Progress includes `Validating existing extracts`, semantic routing-group construction and generation publication. Last-run Admin diagnostics persist the same new counters and accept `completed_with_requirements` under strict validation.
- [x] QA finding/correction: QA1 found that `completed_with_requirements` was initially generated but absent from the strict diagnostic status whitelist. This would have discarded the last-run record for a partially reusable reindex. The status and all new counters were added to strict validation, then complete QA was rerun.

### CP-SA-SR06 — final static/regression acceptance and delivery basis

- [x] Input checkpoint: CP-SA-SR05.
- [x] Product scope: exactly 14 non-governance changed files plus cumulative `Tasks.md`; no generated QA/report/test directory was added to the repository. `CONTRIBUTING.md` required no durable-rule change.
- [x] QA1 final: 52/52 PASS after the final source correction. It covers hard index-only guard ordering, routing schema/persistence/validation, DAG multi-parent families, bounded splitting/collision fallback, exact-channel non-gating, batched semantic routing/continuation/ACL, compact-card routing, large-document section-index contract, PDF fallback/coverage/retry classification, Admin/Worker status/counters, coding rules and synthetic 50/500/5,000/50,000 routing topology.
- [x] QA2 final: 41/41 PASS using an independent diff/control-flow/persistence/ACL/PDF/Worker path, including exact changed-file allowlist, project XML parsing, index-only exporter reachability, optional graph migration, graph integrity rules, current-source checks before model/disclosure, no indexed plaintext fallback, PDF all-page verification and strict diagnostic validation.
- [x] Harness history: QA1 attempt 1 had six failures; five were over-specific/defective harness assumptions and one exposed the real diagnostic-status whitelist defect fixed in CP-SA-SR05. QA1 attempt 2 had one defective lexical block-count assumption caused by existing multiline VB lambdas; it was replaced by targeted structural checks. QA2 attempt 1 had three harness-string mismatches (PDF local variable names and UI label expectation); source inspection showed two harness issues and one unapplied UI-label replacement, which was then correctly applied. Both complete suites were rerun after the final source change.
- [x] Toolchain/native gate: `msbuild`, `dotnet`, `vbc`, `vbnc`/`xbuild` and a Windows/VS2022/VSTO/Office runtime are unavailable here. No native compile or Office runtime result is claimed.
- [x] Synthetic scalability evidence: max child 48 produced representative structural depths of 1 for 50, 1 for 500, 2 for 5,000 and 3 for 50,000 in the balanced topology; a 50,000-document complete semantic-hash collision uses the bounded identity fallback at depth 2. These are structural tests, not real model latency measurements.
- [ ] Native acceptance still required: clean VS2022 Release/AnyCPU build of SharedLibrary, Word, Outlook, Excel and Worker; real model-call/time measurements at representative archive sizes; malformed `/FirstChar` PDFs with OCR enabled/disabled; two-user ACL/shared-artifact/library tests; and actual worker `reindex` proof of zero extraction/OCR calls on valid existing extracts.


### CP-SA-SR07 — native compile correction after VS2022 feedback

- [x] Input checkpoint: CP-SA-SR06 plus the user's native VS2022 compile diagnostics from 2026-10-04.
- [x] Root cause: `SemanticArchiveExplicitRetryProcessingException` inherited from `System.IO.InvalidDataException`. In the target .NET Framework this base type is `NotInheritable`, producing BC30299 and preventing SharedLibrary from compiling.
- [x] Cascading diagnostics: the reported Worker BC30456 errors for `IndexOnlyRebuild`, `ExtractsReused`, `CardsRebuilt`, `SectionIndexesRebuilt`, `RoutingGroupsBuilt` and `DocumentsRequiringExtraction` are not missing-source defects in the delivered tree. Those members are present on `SemanticArchiveBuildOptions` / `SemanticArchiveBuildResult` in `SemanticArchiveWorkQueue.vb`, that file is included by `SharedLibrary.vbproj`, and `SemanticArchive.Worker.vbproj` references the SharedLibrary project. They can therefore arise while Worker analysis sees the previous successfully built SharedLibrary metadata after the SharedLibrary compile failure.
- [x] Applied change: the internal retry-marker exception now inherits from `System.Exception`. Its behavior is intentionally marker-based (`TypeOf ... Is SemanticArchiveExplicitRetryProcessingException`), so no InvalidDataException semantics are required and retry classification remains unchanged.
- [x] Regression invariants: deterministic malformed-PDF failures still bypass generic automatic retry; explicit `Retry failed` remains able to reset them; no public contract, persisted schema, Worker member, extraction behavior or routing behavior changed.
- [x] Verification pass 1: reread full `SemanticArchiveBuilder.vb`, `SemanticArchiveWorkQueue.vb` and Worker `Program.vb`; confirmed the marker type has no dependency on InvalidDataException handling and all Worker-referenced new members exist in source.
- [x] Verification pass 2: independently verified project inclusion/reference wiring (`SemanticArchiveWorkQueue.vb` is compiled by `SharedLibrary.vbproj`; Worker has a `ProjectReference` to `SharedLibrary.vbproj`) and scanned for any remaining inheritance from `System.IO.InvalidDataException`; none remains.
- [ ] Native validation remains required after applying this correction. CP-SA-SR06's statement that no native compile had been run is superseded by the user's compile attempt: that attempt failed with the diagnostics above and therefore is not a successful native build.
### CP-SA-AUD01 — audit correction of routing, continuation, durable reindex and PDF recovery

- [x] Input checkpoint: CP-SA-SR07. Reconstructed the exact cumulative input from `RI_Gen2_261004_2026.zip`, the scalable-routing/reindex changed-only delivery and its compile-fix delivery. The prior `System.Exception` inheritance correction is retained. Full source and `.inky` remain cumulative; no earlier delivered change is silently dropped.
- [x] Audit finding / superseded conclusion: CP-SA-SR06's source/structural checks did not prove a complete semantic implementation or faster retrieval. The old routing builder grouped text hashes rather than model-derived meaning. Its doubly escaped VB Unicode-token pattern returned no tokens for eight varied metadata examples, so those cards all reached the same fallback hash. The old fallback repartitioned rank-ordered identities and was not a sound incremental semantic grouping strategy. These findings supersede the earlier completeness claim; prior checkpoint history remains above.
- [x] Routing changes: replace hash membership with shared-Indexer semantic assignment and balanced local semantic splits in two content/intent projections sharing logical document references. Persist v2 routing/card/profile signatures, reuse unchanged routing cards and update only affected semantic membership/card paths. Old v1 routing and extraction generations remain readable. The host validates complete assignment, unique IDs, fanout, depth, reverse/forward memberships, reachability, cycles and exact summary-contributor provenance. Logged local structural fallback retains documents but explicitly does not claim semantic grouping quality.
- [x] Search root causes/fixes: selected leaf work lost priority to equal-priority groups; selected scores also lost inherited beam priority. Partial/result-limited selection marked all considered documents done, even when only 12 of 32 were retained. Repeated document routes could rerank the same document; failed batches removed arbitrary queue elements. New traversal commits each successful model request, preserves deferred candidates and failed frontiers, rotates deferred work, globally deduplicates document evaluation and prioritizes selected leaves. Full document metadata replaces the lossy compact projection; the exact channel no longer controls small-catalog eligibility and is capped at one second before semantic work. No mandatory lexical gate is introduced.
- [x] Security: persisted v2 summaries may be disclosed only after current original-source checks for all exact contributing document cards. Otherwise the request uses a freshly authorized subset or permission-neutral navigation. Bounded summaries are explicitly not exhaustive evidence; all unexplored branches remain continuable. Search/read separation, source ACL checks and the existing indexed-document evidence path are unchanged.
- [x] Reindex root causes/fixes: discovery selection was not the same as processing/backpressure selection; continuation could derive scope from a different published inventory; index-only intent was not durable on queued jobs; valid local extracts could enter shared-claim waits. Build options are now snapshotted, reindex binds to the operation's fixed published inventory, discovery/processing/backpressure/resume use the same selection, and persisted job/checkpoint mode guards the sole exporter call even during later background drain. Valid local extracts bypass shared import/claims; no index-only path acquires a shared writer claim. Invalid extracts still produce `needs_extraction`, not OCR.
- [x] Extraction compatibility: validate byte length and encoding in addition to private placement, source hash and extract hash. For batch-size-only reindex changes, reconstruct the historical archive signature over supported batch sizes and independently recompute its exporter signature under the current extraction configuration. Reuse is allowed only when both proofs and byte validation succeed. Do not overwrite representation IDs, options signatures, source maps, timestamps, coverage or bytes. Unverifiable or genuinely changed extraction contracts remain `needs_extraction`.
- [x] PDF root causes/fixes: later selective-OCR failure discarded previously completed text chunks; fallback page discovery treated transient file/share errors as permanent parser failures; clamping invalid ranges could overstate coverage. Retain completed OCR chunks and only their supported ranges as an explicitly incomplete representation; require strictly valid all-page ranges before `complete`. Preserve cancellation and classify fallback I/O/access failure separately from deterministic parser failure. No new SA-specific converter or global partial-search enablement was added.
- [x] Affected production files: `SemanticArchiveRouting.vb`, `SemanticArchive.Schema.vb`, `SemanticArchiveSearchService.vb`, `SemanticArchiveBuilder.vb`, `SemanticArchiveBuilder.Scan.vb`, `SemanticArchiveBuilder.Cooperative.vb`, `SemanticArchiveWorkQueue.vb`, `SemanticArchiveQueueIndex.vb`, shared `SharedMethods.FileImporter.vb`; cumulative notes in `Tasks.md`.
- [x] Verification pass A: actual-source/call-contract/control-flow and project XML checks; current summary contributor checks; hard index-only exporter guard; queue mode serialization; no Await in Catch/Finally in changed VB; unchanged Read/Access/RunScope/Inventory and `.inky` bytes. This is not VB compilation.
- [x] Verification pass B: independent external Python algorithm/contract simulations: 50/500/5,000/50,000-document balanced projection structures; fanouts 2/3/4/8/48/256; insertion locality and membership coverage; 32 considered/12 selected counterexample (old loss 20, retained continuation delivers all 32 once across two projections); deferred-card rotation; batch-only compatibility with rejection of source/bytes/profile/options/config changes; native/OCR merging, unsupported-range removal, partial coverage and cancellation/operation-mode cases. These do not execute the production VB, an actual model, a malformed PDF, or Windows ACL APIs.
- [x] Intermediate WIP corrections: the first routing call draft referenced nonexistent access-context interface/type names; actual source contracts were inspected and substituted before checks. The first builder draft lacked an empty-assignment early return; the independent algorithm-to-source comparison caught and fixed that unnecessary recursion. Neither WIP state was delivered.
- [x] Harness failures, not product defects: initial source assertions assumed a variable named `groupPriority`, a filename without the `SemanticArchive.Inventory.vb` dot and a file rather than directory named `.inky`; corrected against actual sources. A broad assignment regex also matched a comparison and was narrowed to statement starts. Re-executed the source suite and independent recovery suite successfully. No failed harness attempt is represented as a native product failure.
- [ ] Toolchain/native gate: no dotnet/msbuild/vbc/Mono/Windows/Office toolchain exists here. Attempts to obtain tooling could not resolve external package hosts. No native compile, OCR recovery, provider request or Windows/SMB permission behavior is claimed to have run successfully.
- [ ] Performance/coverage gates: real model call counts, end-to-end timings, synonym recall and cancellation/resume must be measured with the current source under VS2022. Initial semantic construction costs more offline model work than hashing. The routing graph is still a monolithic validated artifact with a 64 MiB reader limit; cold loading/validation and current-source ACL checks are not constant-time. Bounded summary sampling and fallback paths require recall benchmarks and continuation; they are not a proof of exhaustive first-page recall.
- [ ] Next: final independent diff/project/invariant checks, rerun the relevant external suites after the last production edit, then deliver only changed production files plus `Tasks.md` in repository structure, with external exact find/replace instructions and audit scope. Native build and runtime acceptance remain open rather than being relabeled PASS.
### CP-SA-AUD02 — section-selection countercheck and final source audit

- [x] Input checkpoint: CP-SA-AUD01. During the final independent evidence-path review, the read implementation itself was inspected rather than treating its unchanged hash as sufficient functional QA.
- [x] Established root cause: `ReadIndexedFileAsync` retained potentially missing sections only when at least one section had been selected. An incomplete empty selection could therefore advance the cursor past still-eligible sections. It also accumulated up to 24 selected sections even when a much smaller excerpt page was requested.
- [x] Applied change: retain considered but unselected sections for incomplete/possibly-missing responses even with zero chosen IDs; yield instead of immediately repeating an empty uncertain selection; use one selector request per iteration; stop further section selection once enough candidates exist for the requested excerpt page (bounded by the existing 24-entry limit). No whole-document fallback or literal bypass was introduced.
- [x] Regression correction found by cross-contract comparison: the new reindex encoding check initially omitted the already accepted `utf8` alias. It now trims and accepts the same UTF-8 aliases as the existing small-file reader, while still rejecting incompatible encodings. This was a WIP regression and was corrected before delivery.
- [x] Affected files: `SemanticArchiveSearchService.Read.vb`, `SemanticArchiveBuilder.vb`, `Tasks.md`. CP-SA-AUD01's observation that the read file was unchanged applied to that earlier checkpoint only; its evidence/security contract remains preserved after this targeted correction.
- [x] Verification pass A: reran the complete source/control-flow/project suite after both edits. Independently compared the whole read transaction, source/hash authorization and small-file reader prefix, and all exact selected-range loading/provenance/disclosure/rollback suffix bytes with CP-SA-SR07: unchanged. Confirmed section selection remains mandatory before byte-range loading.
- [x] Verification pass B: reran the external 50/500/5,000/50,000 structure/insertion/continuation suite and the independent recovery suite. Added a 32-section counterexample with 16 considered, no choices and incomplete coverage: previously 16 skipped sections are now retained along with the uninspected tail. The excerpt-page stopping condition avoids additional selection once sufficient pending sections exist. These are source checks and independent Python contract simulations, not native VB or model execution.
- [x] Preserved baseline: original source access, section payload hashes, source-map/offset provenance, evidence-byte budgets, literal-section selection, all-or-nothing read rollback, original compile fix, project/worker references, host code and cumulative `.inky` are unchanged except the explicitly listed production edits.
- [ ] Open acceptance remains explicit: native VS2022 build and Windows/Office runtime; actual OCR recovery fixtures; actual persisted-generation migration/resume; two-user ACL/SMB behavior; real model semantics/recall, model-call counts and timings. Cold monolithic graph loading and source authorization costs remain performance constraints; structural tests do not establish that the implementation is faster in production.
- [ ] Next: validate the changed-only delivery twice against the cumulative working tree, keep temporary checks and find/replace instructions outside product sources, and apply the normal native build/reindex/benchmark acceptance gates.
### CP-SA-AUD03 — changed-only delivery preflight

- [x] Input checkpoint: CP-SA-AUD02; no further production-code edits.
- [x] Preflight checks: changed-only archive contains exactly 10 changed production files plus `Tasks.md`, with original repository-relative paths. First pass checked every archive member byte-for-byte against the cumulative source. Independent second pass overlaid the archive on CP-SA-SR07 and compared all 652 resulting files with the full current source, including all 146 `.inky` files. No source file was removed and no prior fix was dropped.
- [x] Exact find/replace instructions were generated outside the product tree and replayed against the cumulative input; every find block is unique and the complete replay reproduces the new file texts. Diffs, audit notes, verification scripts and delivery manifests remain outside production sources.
- [x] All relevant source/control-flow, structure/continuation and independent recovery simulations were rerun after the last production-code correction. None is a native build or real-model timing result. The full cumulative source/`.inky` base is retained separately; user-facing source delivery remains changed-only.
- [ ] Finalization: regenerate the package and external instructions after this documentation checkpoint and repeat both package validations. Native build, actual reindex/OCR/ACL scenarios and end-to-end speed/recall benchmarks remain the explicit runtime acceptance gates from CP-SA-AUD02.

## Semantic Archive — second audit: OCR, worker output and current indexes only

### CP-SA-R2-00 — reconciled cumulative input and targeted audit

- [x] Input checkpoint: CP-SA-AUD03; reconstructed `RI_Gen2_261004_2026.zip` plus the ScalableRouting/Reindex, compile-fix and AuditFix changed-only packages in that order. All 652 files, including the 146 `.inky` files, are retained in a separate immutable baseline.
- [x] Performed step: read the current contribution rules, prior audit/checkpoints and actual Worker/PDF/extraction/routing contracts; no production-code change in this checkpoint. The user now explicitly rejects backward compatibility with old indexes. This supersedes the earlier old-routing migration requirement; unrelated current catalog settings, source identity and valid current extraction reuse remain required.
- [x] Findings: console progress uses carriage-return overwrites with a fixed 180-character limit, so wrapping/resizing corrupts presentation. PDF batches are accepted by text length rather than an explicit per-page result; ordinary selective OCR discards the retained merged result on partial failure. The existing layout-aware PDF text helper is bypassed by `page.Text`.
- [x] Invariants: stdout/optional private JSON log remain machine-readable; no console cursor/width dependence; no new SA conversion pipeline; native text and verified OCR pages are retained; incomplete/unknown is not silently searchable; current source ACL/hash/evidence checks and hard index-only no-exporter guard remain intact. Remove obsolete index-reader paths, not valid current index format validation or generic reader APIs used outside SA.
- [x] Checks: reconstructed patch/member paths validated, current worker and extractor call sites inspected, existing compile-fix confirmed. Toolchain discovery found no dotnet/msbuild/mono/vbc; downloading official tooling failed with DNS resolution error. This is an environment limitation, not a product test failure.
- [ ] Next: append-only worker progress and focused checks; page-structured OCR/recovery; current-index-only reader cleanup; two independent final source/contract and delivery checks. Native Windows/VS2022/Office, real-model OCR and runtime performance remain explicit open gates.

### CP-SA-R2-01 — append-only worker progress

- [x] Input checkpoint: CP-SA-R2-00. Affected files: `SemanticArchive.Worker/Program.vb`, `WORKER_HELP.txt`, `README.md`, `Tasks.md`.
- [x] Root cause/change: fixed-width carriage-return rewriting assumed one physical console row. Progress now uses complete stderr lines with no padding, truncation, cursor clearing or width queries. Only identical consecutive source/stage events are suppressed; filenames are sanitized for terminal controls. stdout and the optional private JSON log retain their JSON-only contract. Help describes redirected stderr and its filename exposure.
- [x] Checks A: no residual overwrite/clear/width code or progress field/constructor references; unchanged stdout serialization path and exit/cancellation flow. Checks B: independent line-contract probes cover 500-character names, CR/LF/ESC and Unicode separators, Unicode preservation and equal display filenames in distinct directories. These are source/contract checks, not Windows terminal execution.
- [x] Failed intermediate attempts: the help patch assumed CRLF while that supplied file uses LF; the exact-anchor guard rejected it without changing the help file. A broad `_progress` test also matched the unrelated `no_progress` diagnostic; narrowed to a symbol-boundary check. Applied the actual LF block and reran the complete worker-focused check. Neither failure was a product defect; no failed WIP is released.
- [ ] Open: native build and terminal resize/redirection acceptance. Next: fix shared PDF/OCR page accounting, partial-result preservation and adaptive recovery; rerun worker checks after any progress integration.

### CP-SA-R2-02 — shared OCR page contract and actual fresh extraction

- [x] Input checkpoint: CP-SA-R2-01. Affected production files: shared `SharedMethods.FileImporter.vb`, `TextExtractionResourceRegistry.vb`, `TextExportService.vb`, `TextTools.SemanticAndExport.vb`, `SemanticArchiveBuilder.vb`; corrected Worker diagnostic-code allowlist in `Program.vb`. Prior cumulative fixes remain in place.
- [x] Established causes: native PDF extraction bypassed the existing reading-order helper; one deterministic font failure could discard already read native pages; OCR response length incorrectly stood in for complete page coverage; the normal partial-OCR return discarded recovered content. Cache reuse also treated explicitly incomplete results as completed work and could satisfy an explicit force-extract operation without invoking the extractor.
- [x] Changes: read native pages individually with the existing layout helper; preserve native page slots on known font failures. Require typed, unique, bounded attachment-relative page results in a finished JSON envelope. Accept explicit blank/short pages without length heuristics; retain unreadable partial text without counting it as complete. Retry only missing/unreadable pages in smaller bounded ranges; split depth no longer consumes the per-request retry budget. Plain error/empty/refusal responses and transport errors receive bounded same-range retries, not a per-page request explosion. Cancellation/headless requirements propagate, deterministic chunk-parser errors do not retry identically, and temporary PDF chunks are deleted.
- [x] Changes: preserve completed OCR pages plus untouched native text on partial failure. OCR failure diagnostics take priority over repetitive native-page observations within the existing bounded source-text-free log contract. Remove unused duplicate private OCR implementations. Verify alternate OCR model PDF capability and include the actual OCR prompt in its fingerprint. Version the PDF-only extraction contract; non-PDF signatures remain byte-for-byte unchanged.
- [x] Generic extraction change: known incomplete payloads are returned but not session-cached. The explicit `ForceFreshExtraction` execution option bypasses both prior cache results and pre-existing single-flight work, without changing source identity or processing signatures. The durable SA force-extraction flag reaches this option through the existing shared exporter; ordinary completed-cache/single-flight reuse remains unchanged. Force does not override output overwrite permissions. The hard index-only exporter guard is unchanged.
- [x] Verification A: 41 actual-source/contract checks passed; Worker checks rerun after the diagnostic allowlist correction. Verification B: 217 independent Python parser/state-machine probes passed, including malformed/duplicate/truncated envelopes, short/blank/Unicode pages, attachment numbering, partial/native merging, bounded adaptive recovery up to 75 pages, 80 randomized coverage/deduplication cases, cancellation, transport/plain errors and cache/force truth tables. These do not execute production VB, PDFs, a model or Windows APIs.
- [x] WIP counterchecks: initial adaptive draft conflated split depth with retry attempts and would skip leaf retries; corrected before this checkpoint. Unstructured endpoint failures would have triggered unnecessary recursive splitting; now bounded without provider-specific text matching. A partial-only cache fix did not establish the explicit force-extract contract; the generic force option was added and independently checked. OCR page counts are structured declarations, not proof of pixel-perfect transcription.
- [ ] Open: native VS2022 build, real OCR fixtures/configurations and model-response/latency acceptance. No concrete user OCR log or problematic PDF was supplied, so the exact runtime cause/recognition quality is not asserted. Next: remove old index compatibility and the storage-tree search fallback, then rerun complete source/contract and independent delivery QA.

### CP-SA-R2-03 — current-index-only routing and continuation

- [x] Input checkpoint: CP-SA-R2-02. WIP reconciled after context refresh against the immutable cumulative baseline and the actual four changed routing/schema/validation/search files before continuing. Affected production files: `SemanticArchive.Schema.vb`, `SemanticArchiveRouting.vb`, `SemanticArchive.Validation.vb`, `SemanticArchiveSearchService.vb`; contribution rules and Worker help/README document the explicit current-only contract.
- [x] Root cause/change: the previous implementation retained v1 routing exemptions, a missing-DAG fallback and the old node-per-call storage-tree search engine, including small-catalog continuations. Removed those obsolete branches, old traversal state/methods and hash-era `Prefix`. Only current routing schema v2 is accepted; a missing persisted version defaults to invalid zero. Published generations require a current routing graph; builders assign the current version explicitly. Existing current v2 is not gratuitously bumped. Old-index migration is intentionally unsupported by the user's new instruction, superseding the migration statements in CP-SA-AUD01 and older checkpoints.
- [x] Current behavior preserved: small catalogs retain complete direct document-card selection; partial selections continue through the same initialized DAG used for large catalogs. Exact metadata remains independent and cannot exclude semantic candidates. Search group/leaf selection, candidate retention, live-source authorization, document read/section evidence, current storage/shard persistence, source identity, queue/index-only contracts and .inky files are unchanged except for the removed old-format gate. Current validated text reuse and unrelated mutable configuration compatibility remain deliberate, not obsolete index readers.
- [x] Avoided redundant work: generation publication/full audit now call the already-validating graph loader once instead of validating the same complete graph twice. This removes a traversal, not the integrity/security validation. No production latency improvement is claimed without measurements.
- [x] Verification A: 195 source/project/retained-.inky checks passed. Verification B: 219 independent Python current-format/topology/frontier specification probes passed, including version 0/1/future rejection, absent DAG, signature/edge/contributor/depth failures, 50/500/5,000/50,000 structural sets and partial-flat continuation with deduplication. These are not VB/model execution or runtime benchmarks. Worker and both OCR suites (41 source + 217 independent probes) reran successfully after the code changes.
- [x] Failed harness attempt: treating `.inky` as a filename extension matched the directory itself and raised `IsADirectoryError`. Corrected the external harness to enumerate its 146 contained files; no production source correction was necessary. Complete index checks reran successfully.
- [ ] Open: native build/terminal resizing, actual OCR PDFs/model outputs, current-format refresh/reindex/resume and source-ACL scenarios. No automatic deletion/reset or legacy-index import is implemented; obsolete state requires a new archive index. Next: final cross-file/source-syntax checks, full reruns, exact find/replace replay and independent changed-only overlay validation outside the product repository.

### CP-SA-R2-04 — final cross-file regression and release source

- [x] Input checkpoint: CP-SA-R2-03. No additional production code changes. Re-read changed call paths and final diffs, including extraction cache/force, PDF native/OCR/partial returns, graph read/write/publish and small-catalog continuation. Worker help and contribution rules consistently distinguish current indexes, extraction contracts and explicit OCR recovery.
- [x] Final complete rerun: Worker source/line-contract probes PASS; OCR 41 source checks and 217 independent Python contract/state probes PASS; current indexes 195 source/project/.inky checks and 219 independent topology/frontier probes PASS. Additional independent cross-file/limited-lexical pass: 211 checks PASS for BOM/newline preservation, changed-member project inclusion, dangling removed private helpers, public shared reader APIs, compile-marker inheritance, added namespace/IsNot conventions, parenthesis deltas and no Await inside Catch/Finally. None of these is a VB compiler or real model/PDF/Windows test.
- [x] Regression invariants: current access checks and exact/section reads, the hard no-exporter index-only guard, stable logical identities, generic non-PDF extraction fingerprints, normal completed-result reuse, Office reader APIs and original OCR dialog owner/thread/disposal implementation are retained. All 146 files under `.inky` are byte-identical; the full 652-file cumulative source remains separate from changed-only delivery.
- [x] Scope: 12 changed production/documentation files plus `CONTRIBUTING.md` and `Tasks.md`. Removed code is the unused duplicate private OCR pipeline, obsolete routing format acceptance and old storage-node search traversal—not current storage persistence or public extraction APIs. No test, QA, diff, report or tool files were added to production.
- [ ] Native gates still open: clean VS2022 SharedLibrary/Word/Outlook/Excel/Worker build; Windows terminal resizing and redirection; actual mixed/blank/short/malformed/scanned PDF fixtures and configured model output; cancellation/force-cache concurrency and reindex with real persisted generations; SMB/current ACL behavior. Recognition fidelity and end-to-end latency/recall cannot be established from source checks. Cold monolithic routing (64-MiB reader bound) remains a scaling consideration, not addressed by this focused revision.
- [ ] Next: exact find/replace replay and two independent changed-only package validations, append delivery checkpoint, regenerate all artifacts, then repeat package checks. Release notes must state the runtime limits and that an OCR repair requires explicit extraction, not semantic reindex.

### CP-SA-R2-05 — changed-only delivery verification

- [x] Input checkpoint: CP-SA-R2-04; no further production-code changes. External delivery preflight contains exactly the 14 changed files, preserving their repository-relative paths and cumulative source contents. No QA/test/report dependencies or files are present in the product package.
- [x] Verification 1: ZIP member allowlist, path safety, CRC and every member's exact bytes match the working source. Verification 2: independently copied the full immutable input, overlaid the ZIP, and compared all 652 file hashes with the cumulative output; all 146 files under `.inky` remain byte-identical. The full current source/.inky basis is retained independently, while the delivered source ZIP is changed-only.
- [x] Exact instruction verification: 73 unique find/replace blocks replay in order against the input and reproduce every changed file byte-for-byte, including existing BOM/newline conventions. Unified diff, exact instructions, audit and temporary checks remain outside production.
- [x] Source acceptance is bounded by CP-SA-R2-04: native compilation, actual PDF/model/Windows/ACL execution and end-to-end performance are not performed or claimed. No failed WIP or unverified production edit was added after the final suites.
- [x] Finalization: package/instructions regenerated after this checkpoint; both package validations and exact replay repeated successfully. Operational acceptance remains a Windows/VS2022 build plus actual worker resize/OCR/current-index scenarios; use explicit extraction to repair OCR, not index-only reindex.

### CP-SA-R2-06 — late OCR presentation-policy countercheck

- [x] Input checkpoint: CP-SA-R2-05. A final cross-check against the generic LLM postprocessor found an additional genuine extraction defect: ordinary `INI_DoubleS`, `INI_NoDash` and `INI_Clean` presentation settings could rewrite OCR source spelling, dash characters, spacing or hidden marks after the model's page response. The earlier page-structure tests did not cover this host-side transformation. The prior delivery preflight is superseded; its history is retained.
- [x] Applied correction: the existing `CreateIsolatedModelCallContext` helper now isolates the private PDF OCR call before any OCR model selection. The three presentation settings are disabled only on this isolated call after model resolution. The OCR fingerprint records the fixed `postprocess=verbatim` policy. The generic LLM implementation, normal answer-formatting behavior, non-PDF callers and caller/Office context remain unchanged. Durable extraction rules updated in `CONTRIBUTING.md`.
- [x] Verification A: expanded OCR source suite 46/46 PASS; all complete Worker, index (195 source + 219 independent), OCR-state (217 independent) and cross-file/limited-lexical (211) suites reran successfully after the source correction. Verification B: 30 independent formatting-policy/host-immutability probes PASS for every combination of the three settings, retained source symbols/spacing and byte-identical generic model helpers. These are not production VB execution or OCR quality measurements.
- [x] Regression invariants: source text is not silently normalized by presentation policy; no live host settings are changed. Existing structured coverage, missing-page retries, incomplete-text preservation, no-cache partials, force-fresh execution and strict current routing contracts remain intact. No new production file or dependency introduced.
- [ ] Next: regenerate final changed-only artifacts and exact instructions from this corrected cumulative source; repeat both full delivery validations. Native build, actual PDF/model recognition, Windows console and runtime performance gates remain open as above.

### CP-SA-R2-07 — corrected final delivery

- [x] Input checkpoint: CP-SA-R2-06; no further production-code edits. All seven relevant suites reran after the isolated OCR formatting correction: Worker PASS; OCR source 46 PASS; OCR independent 217 PASS; index source/project/.inky 195 PASS; index independent 219 PASS; cross-file/limited-lexical 211 PASS; independent source-formatting/host-policy 30 PASS. Test categories and native limitations are unchanged.
- [x] Rebuilt delivery and repeated the independent validations: exactly 14 changed files; member bytes/CRC/path allowlist match; baseline-plus-ZIP reconstructs all 652 cumulative file hashes; all 146 .inky files retained; all 73 exact replacement blocks replay to the correct bytes. The final full source/.inky basis is maintained separately. Audit, diff, instructions and verification tooling remain outside product sources.
- [x] Finalization includes regeneration after this documentation entry and the same independent ZIP/overlay/exact-replay checks. No code change occurs between the successful final suites and this delivery.
- [ ] Native acceptance: build SharedLibrary and all consuming projects in VS2022; resize/redirect the worker console; test actual OCR PDFs and configured model output, explicit extraction versus zero-extraction reindex, cancellation/resume and current ACLs. No production timing or recognition-fidelity claim is made from these static/simulated tests.

### CP-SA-R3-01 — stage-aware retry/repair root-cause correction

- [x] Input checkpoint: CP-SA-R2-07 plus the user's Outlook Admin runtime diagnostics showing an `incomplete` selective-OCR PDF, an `unknown` DOCX (`reader_coverage_unverified`) and a subsequent `Invalid semantic routing graph` failure during `Retry failed`.
- [x] Root cause 1: retry eligibility and extraction reuse were conflated. An intact representation with `Completeness=incomplete/unknown` passed the representation integrity checks, so `Retry failed` could reuse exactly the extraction that needed repair. Discovery now marks extraction-level retry work with `ForceExtractionRebuild=True`; missing, empty, incomplete and unknown extraction bypasses session cache/single-flight via the existing force-fresh exporter contract. A verified current complete extract is retained when only downstream semantic/index work failed.
- [x] Root cause 2: the DOCX adapter returned successful text with no explicit coverage contract, so normal successful sandboxed OpenXML reads remained `reader_coverage_unverified`. A versioned DOCX coverage contract was added. The shared DOCX reader now reports coverage separately from text success: the main package must parse, and parse failures in present supplementary headers/footers/notes lower coverage to incomplete instead of being silently upgraded. Existing callers retain their previous API behavior through the original overload.
- [x] Root cause 3: maintenance loaded the previous routing graph as an incremental optimization and allowed an invalid derived graph to abort the repair before a new graph could be published. Maintenance now catches the strict current-format validation rejection, discards only that derived routing optimization and rebuilds routing from the current validated document cards. Search remains strict and still rejects the invalid graph; no legacy routing reader or migration path was introduced.
- [x] Admin symmetry: Outlook `Retry failed` uses the same stage-aware builder contract. Help/user status explains that missing/empty/incomplete/unknown extraction is refreshed while verified complete extraction is reused. Coverage-excluded counts are accumulated over the complete bounded operation, not reported only from its final batch. Routing-rebuild recovery is technical diagnostic detail.

### CP-SA-R3-02 — Worker `repair` operation and multi-archive overnight scope

- [x] Added Worker operation `repair`. It combines stage-aware retry with `RebuildSemanticMetadata=True`: extraction is repeated only where the extraction stage is unusable, while verified complete extracts are reused; semantic cards, applicable section indexes and current routing are rebuilt for the requested scope in the same drained operation.
- [x] `reindex` remains strictly index-only and never extracts/OCRs. `extract` remains the explicit force-extract-all operation. `retry` repairs failed/deferred/coverage-excluded work by failed stage without forcing a semantic rebuild of every healthy document.
- [x] Friendly Worker shorthand accepts multiple named archives before advanced `--` options. The intended overnight command is `redink-sa-worker.exe repair "VISCHER Compliance" "VISCHER DP Know-how"`; both archives are drained under one Worker process. Advanced `--archive` syntax remains supported.
- [x] Worker JSON batch output includes `coverageExcluded`; any unresolved extraction coverage keeps the final Worker result incomplete (exit code 3) rather than reporting success. `routing_rebuild_required` is an allowed bounded diagnostic code.
- [x] Durable engineering rule added to `CONTRIBUTING.md`: retry/repair is stage-aware; complete current extracts are reused for downstream failures; repair rebuilds current derivatives/routing; invalid derived routing may be discarded during maintenance while search remains strict.

### CP-SA-R3-03 — regression verification and delivery basis

- [x] Existing complete external suites rerun after the final DOCX coverage hardening: Worker source/line-contract PASS; OCR source 46/46 PASS; current-index source/project/.inky 195/195 PASS; cross-file/limited-lexical 211/211 PASS; independent OCR/cache state machine 217/217 PASS; independent current-index topology/frontier 219/219 PASS; independent verbatim-policy 30/30 PASS.
- [x] New independent retry/repair suite: 41/41 PASS. It checks the user's runtime state table (incomplete PDF -> fresh extraction, unknown DOCX -> fresh extraction, missing extraction -> fresh extraction, downstream failure with complete extract -> reuse), force-fresh exporter propagation, strict-search/current-index behavior, maintenance-only invalid-routing recovery, `repair` operation flags, multi-archive shorthand, unresolved-coverage exit behavior, DOCX structured coverage and technical diagnostic classification.
- [x] Harness history: an intermediate retry/repair check failed because it still searched the old per-batch Admin coverage expression after the product was intentionally changed to an operation-total counter; another intermediate pair still expected the earlier unconditional DOCX-complete implementation after it was deliberately hardened. Both were external harness defects and were corrected; no product rollback was made. Complete suites were rerun after the final source change.
- [x] Regression invariants retained: no mandatory lexical gate, source ACL/search-read separation, large-document section-only evidence, current-routing-format-only behavior, hard no-extraction reindex guard, stable logical identity, append-only Worker progress, force-fresh OCR cache bypass and no silent partial-search enablement.
- [ ] Native acceptance remains required: clean VS2022 Release/AnyCPU build of SharedLibrary and Worker (plus consuming Office projects), actual Outlook Admin `Retry failed` on the reported PDF/DOCX, actual `repair` across both VISCHER archives, malformed/scanned PDF OCR/model output, cancellation/resume and current ACL/SMB conditions. No native compile, Windows console, OCR-quality or production-timing result is claimed from these source/simulation checks.

### CP-SA-R3-04 — changed-only delivery verification

- [x] Input checkpoint: CP-SA-R3-03; no further production-code changes after the final complete regression suites.
- [x] Delivery scope: exactly 12 changed production/governance/documentation files plus cumulative `Tasks.md`, preserving repository-relative paths. No QA/test/report/diff artifact is included in the product ZIP.
- [x] Package verification 1: every changed-only ZIP member is byte-identical to the cumulative working source and the member allowlist exactly matches the changed-file set.
- [x] Package verification 2: independently overlaid the changed-only ZIP on CP-SA-R2-07 cumulative source and compared the complete 652-file tree; every file matches the cumulative CP-SA-R3 source.
- [x] Exact edit instructions: 38 external exact find/replace blocks were generated and replayed against the immediate cumulative input, including precise Markdown placement instructions for `CONTRIBUTING.md` and `Tasks.md`. Diff/instruction artifacts remain outside production.
- [ ] Native gate remains unchanged: VS2022/Windows/Office runtime validation is still required before claiming a successful build or real OCR/repair behavior.

### CP-SA-R4-01 — repair intent, published coverage and retrieval continuation (2026-10-05)

- [x] Input: RI_Gen2_261005_0925.zip / CP-SA-R3-04; reconciled the available WIP with the baseline and the supplied search/Admin status (519 sources, 491 complete extracts, 139 searchable, generation 169353ff3a7c454b89ad5d3981ecc93c).
- [x] Fixed durable queue intent: explicit refresh/retry/repair replaces an earlier index-only intent; background resume preserves it. Explicit retry/repair clears stale index-only flags before processing already queued jobs. Hard reindex still returns needs_extraction before any exporter/OCR call.
- [x] Complete needs_extraction records first undergo current-contract or verified OCR-batch-only compatibility checks. A genuine incompatibility triggers fresh extraction in retry/repair, bypassing stale extraction cache results. Source identity, source hash, text byte/hash verification and exporter/profile contracts remain mandatory.
- [x] Extraction-contract diagnostics include stored/expected archive and exporter signatures. Exporter configuration drift remains an explicit retry-stage failure; output under a mismatched signature is never activated. Historical signature components and the OCR FileNotFoundException dependency are not available in the supplied diagnostics; this change does not claim to identify or repair that missing dependency.
- [x] Inventory derives complete-but-not-searchable counts from existing immutable counters without a schema migration. Worker terminal success checks the published requested scope, current routing and source readability; a drained queue alone is insufficient. Permissions-only operations retain their independent success contract. This supersedes CP-SA-R3-02's per-batch coverage success rule.
- [x] Shared semantic selection reports host-owned rejected IDs per successfully evaluated batch. Search/read retain uninspected, related or selection-limit-omitted entries while avoiding repeated evaluation of proven irrelevant entries. Direct-card and DAG selection use the same progress contract. Metadata candidates remain eligible for semantic selection.
- [x] Search yields a metadata-filled page only after document-card evaluation, can resume one unambiguous pending identical query/scope, rotates continuation references, and supplies NextSearchArguments including an existing literal parameter. Tool guidance explains partial coverage and maintenance exclusions.
- [x] QA: 89 external source-contract/structural/project checks and independent behavioral reference cases passed before task-note packaging; final cumulative checks rerun for release. A second review inspected queue matching, current ACL/routing validation, fresh-export propagation, direct/DAG progress, section deferral and literal continuation. Package checks compare exact changed membership/bytes and independently overlay the patch on the input tree. QA scripts, diffs and edit instructions remain outside the product tree.
- [x] Retained invariants: stable logical source identity; current-routing-format-only; no mandatory lexical gate; strict no-extraction reindex; fail-closed source access; section-only evidence for large documents; partial extraction excluded unless expressly configured; shared host-independent logic; all existing .inky files retained unchanged.
- [ ] Native acceptance: VS2022 Release build of SharedLibrary/Worker/Office consumers, real worker repair, representative search/read timings, actual OCR dependency diagnosis, cancellation/resume and source-rights changes. No native compilation or runtime success is claimed here.

### CP-SA-R5-01 — shared OCR exception diagnostics (2026-10-05)

- [x] Input checkpoint: CP-SA-R4-01 and delivered RI_Gen2_261005_cumulative_Code_inky.zip, byte-reconciled against all 652 current files. Runtime repair log confirms 482 complete/searchable records, zero complete exclusions; supplied Admin diagnostics identify FileNotFoundException during OCR for the video-device guidelines and annual report.
- [x] Root cause of the diagnostic gap: selective OCR persisted only the exception type. Range catches rethrew file/access/parser failures without preserving their exact stage; unattended execution lost the missing filename, inner exception message and call site. The actual missing runtime file remains unknown until the targeted native reproduction.
- [x] Added a reusable bounded exception-observation helper in TextExtractionDiagnostics. It records stage, fully qualified type, HRESULT, missing/load filename, bounded inner exceptions, messages, target methods and up to three stack lines per exception. Essential file identity precedes longer details and repetitive native-page notes. Collection failure produces an explicit bounded fallback and never replaces the extraction failure.
- [x] Shared OCR instrumentation records temporary-path, PDF-chunk creation and model-request stages with requested pages; outer configuration/progress/coverage failures also receive detail. The helper is shared by Word, Outlook and Worker, without archive/template/provider-specific branches. No request payload, attachment, model reply or FusionLog is explicitly collected.
- [x] Existing behavior retained: cancellation/interaction/access/file/parser propagation, request retry budget, temporary-file cleanup, native/verified OCR pages, completeness classification, source rights and current extraction signatures. Diagnostic limits are centralized constants; no extractor version/signature change and no automatic reindex/re-extraction of healthy documents.
- [x] QA: 101 existing source/structure/project/reference checks plus 39 targeted OCR diagnostic checks passed. Independent additive-only diffs confirm no original implementation lines were deleted; bounded-warning reference cases check wrapped missing dependency survival with 0/36/63/64/200 prior observations. Package verification checks exact changed bytes, previous-delivery overlay, complete cumulative basis and .inky preservation. Exact find/replace blocks and diff remain external to the repository.
- [ ] Native acceptance: rebuild/deploy SharedLibrary and Worker/Office consumers; retry only doc_fee7f5270dce06e946501fae1da821b2567d5ef38da240002709ac27b2d3c299; copy Admin technical coverage_excluded/ocr_exception fields. Fix the identified missing dependency/file before retrying the remaining excluded PDFs. No Windows/VSTO compile, actual OCR or real dependency identification was performed here.

### CP-SA-R6-01 — preserve OCR diagnostics when exception metadata fails (2026-10-05)

- [x] Input: CP-SA-R5-01 / RI_Gen2_261005_1312_cumulative_Code_inky.zip, all 652 files reconciled. User's targeted native retry produced ocr_exception: exception_detail_collection_failed=System.IO.FileNotFoundException and remained incomplete; archive coverage stayed at 482 searchable records.
- [x] Confirmed defect in R5 diagnostic helper: it collected basic fields into a temporary list, then accessed optional exception metadata under one outer Try. A later collection failure discarded all earlier observations. The supplied type-only fallback does not prove which getter failed or identify the missing OCR/runtime file.
- [x] Corrected shared helper: file, message, target and stack are read independently behind per-field recovery; partial observations are merged in Finally. Secondary diagnostic failures record observation_unavailable, diagnostic_file and diagnostic_message without reflection or recursive stack formatting. Basic secondary getters are separately guarded. Missing filename/message survive optional reflection/stack resolution failure. This supersedes the R5 all-or-nothing collection implementation.
- [x] No OCR/extraction, cache/signature, retry/repair, ACL, source identity, search, modal-owner or .inky changes. Existing bounds, deduplication and essential-file priority remain intact. Error metadata remains in the existing protected extraction/Admin technical channel.
- [x] QA: 101 existing checks, 39 OCR diagnostic checks and 90 recovery source/reference fault-injection checks passed. Independent review confirms Finally retains partial details and secondary formatting never recurses into the failing metadata path. Prior native failure is recorded as a real product diagnostic defect, not a harness failure. Previous OCR harness source anchors were updated to the new per-field contract.
- [x] Changed-only package: TextExportService.vb and cumulative Tasks.md; exact bytes/member set, previous-delivery overlay, complete Code/.inky basis and exact edit replay verified externally. QA artifacts remain outside the product tree.
- [ ] Native acceptance: rebuild/deploy SharedLibrary and consuming Worker/Office projects, repeat the same single annual-report retry, copy current Admin technical ocr_exception entries. The failing observation and actual missing file remain open until that run. No native compile, successful OCR or dependency repair is claimed from this environment.

### CP-SA-R7-01 — Worker logging runtime references and binding policy (2026-10-05)

- [x] Input: CP-SA-R6-01 / RI_Gen2_261005_1336_cumulative_Code_inky.zip, all 652 files reconciled. User's new targeted retry establishes stage=pdf_chunk_creation, PdfSharp.Pdf.IO.PdfReader constructor, missing Microsoft.Extensions.Logging.Abstractions assembly 8.0.0.0; OCR never reached its model request. R6 diagnostics now retain primary fields despite a stack-formatting load failure.
- [x] Source cause: SharedLibrary and Word/Outlook already reference the existing Logging.Abstractions 10.0.11 package; Office configs redirect to their configured assembly version 10.0.0.11. Worker had neither direct logging dependency references nor corresponding source runtime redirects. The supplied log does not reveal whether the deployed newer DLL, generated redirect or both were absent; the correction covers both build/output paths.
- [x] Added Worker explicit Copy Local references for the existing logging dependency closure: Logging.Abstractions, DependencyInjection.Abstractions, DiagnosticSource, Bcl.AsyncInterfaces, Buffers, Memory, Tasks.Extensions and CompilerServices.Unsafe. Assembly identities and package HintPaths are copied from the current SharedLibrary contract, not guessed or downgraded. Worker App.config uses the same eight binding redirects as Word. Existing automatic redirect generation remains enabled.
- [x] Added a generic non-design-time pre-resolution check for missing Worker HintPath assets; package restore failures no longer silently yield an executable missing declared runtime references. README explains deployment of the generated executable config and complete DLL output together.
- [x] Official references checked: PDFsharp 6.2.4 package declares Logging.Abstractions >=8.0.3; Logging.Abstractions 10.0.11 net462 declares DI.Abstractions, DiagnosticSource, Buffers and Memory; DI.Abstractions net462 declares Bcl.AsyncInterfaces/Tasks.Extensions. Microsoft documentation confirms runtime binding policy belongs in the executable output config. No .NET 8 runtime install, OCR fallback or index-policy change is introduced.
- [x] QA: 101 previous structural/source checks, 39 OCR diagnostic checks, 90 recovery fault-injection reference checks and 58 new Worker XML/reference/dependency checks passed. Second review confirms exact Office redirect parity, 8.0.0.0 range coverage, unchanged startup/project reference and byte-identical product VB/.inky. Changed-only bytes/member set, original previous-delivery overlay, cumulative basis and exact edit replay verified externally.
- [ ] Native gate: restore existing solution packages; Clean/Rebuild Release in VS2022; inspect Logging.Abstractions DLL identity and generated redink-sa-worker.exe.config in output; repeat the same annual-report retry. Only after OCR reaches the model and coverage is validated should remaining failed extraction be retried. Actual compiled output, package DLL identity, runtime binding and OCR success have not been verified in this environment.

### CP-SA-R8-01 — correct nested ProjectReference metadata after native Clean failure (2026-10-05)

- [x] Input: CP-SA-R7-01 / RI_Gen2_261005_1401_cumulative_Code_inky.zip. User native Clean reported missing/invalid Project metadata for SharedLibrary.
- [x] Confirmed product defect: the R7 insertion script replaced every closing </Project> tag, including ProjectReference's GUID-valued Project metadata element. This left valid XML but invalid MSBuild metadata. R7 XML/reference-count checks did not validate metadata leaf structure; their PASS did not establish a valid native build.
- [x] Removed the nested target from the metadata; restored the exact SharedLibrary GUID leaf. The existing eight Copy Local runtime references, source version redirects, startup and single root-level validation target remain intact. Corrected insertion tooling to use only the outermost final closing tag.
- [x] QA: all 101 existing source checks, 39 OCR diagnostic checks, 90 recovery checks and 62 dependency/metadata checks passed. New assertions require GUID-shaped leaf metadata equal to the referenced ProjectGuid and all Targets directly under the project root. Independent replay from R6 with the corrected insertion script exactly reproduces the three current dependency/config/documentation files. This is a real product defect correction, not a test-only change.
- [x] Delivery: changed-only Worker vbproj and cumulative Tasks.md; previous-delivery overlay/exact bytes, cumulative Code/.inky basis and exact edit replay checked. All prior VB/OCR, dependency redirects and .inky content remain unchanged.
- [ ] Native acceptance: reload the Worker project in VS2022, Clean/Rebuild Release, then repeat the annual-report-only retry. No native Windows/MSBuild compile or successful OCR is claimed here.

### CP-SA-R9-01 — Worker ValueTuple inbox binding policy (2026-10-05)

- [x] Input: CP-SA-R8-01 / RI_Gen2_261005_1411_cumulative_Code_inky.zip; all 652 files reconciled before editing. User retry now reports System.ValueTuple 4.0.5.0 (inner requested 4.0.0.0), still at PdfReader.OpenFromStream / pdf_chunk_creation, before any OCR request. Logging failure is absent from this run; no successful OCR is inferred.
- [x] Audited SharedLibrary/Worker package references and runtime policy. SharedLibrary already imports System.ValueTuple.4.6.2/build/net471/System.ValueTuple.targets, but Worker has no equivalent target. Microsoft's net471+ target removes automatically suggested ValueTuple redirects because these frameworks provide ValueTuple inbox; a redirect to the higher-version package implementation can defeat that binding. Actual generated user exe.config was not supplied, so its contents remain a native verification point.
- [x] Worker now applies that same SuggestedBindingRedirects removal after ResolveAssemblyReferences and before GenerateBindingRedirects, scoped to its existing net48 target. This preserves the framework policy rather than deploying a second tuple implementation or guessing a facade version. Existing eight explicit logging dependency references/redirects, automatic redirects for all other assemblies, startup and valid ProjectReference leaf metadata remain intact.
- [x] README records full-output rebuild/deployment and the observable acceptance check: generated executable config contains no System.ValueTuple bindingRedirect, logging redirects remain, targeted PDF retry reaches beyond pdf_chunk_creation. No content/search coverage policy, extraction fallback, provider behavior or source permissions were changed. All product VB and .inky bytes match R8.
- [x] QA: 292 previous source/diagnostic/dependency checks and 12 focused tuple policy checks passed. Independent byte review removes only the new root target block and recovers R8 Worker project exactly. All Targets remain at project root and ProjectReference metadata remains the original leaf GUID. Exact changed ZIP bytes, previous-delivery overlay, cumulative 652-file Code/.inky basis and find/replace replay verified. These checks do not execute MSBuild or CLR.
- [ ] Native acceptance: reload project; Clean/Rebuild Release with existing packages restored; inspect the freshly generated redink-sa-worker.exe.config for the condition above; retry only the annual report. No Windows/MSBuild build, package binary inspection, CLR load or successful OCR was possible in this environment.

### CP-SA-R10-01 — usable bounded maintenance document selection (2026-10-05)

- [x] Input: CP-SA-R9-01 / RI_Gen2_261005_1435_cumulative_Code_inky.zip; all 652 files reconciled. User's targeted worker retry confirms one complete/searchable selected PDF and complete=true, with no dependency/OCR exception. Subsequent 11-document Admin extract publishes 494 searchable/complete, 16 incomplete and 9 empty; no failures or exclusions in that selected batch. This establishes those native workflows; it does not establish every remaining document or global permission discovery completion.
- [x] Existing invariants retained: same shared Word/Outlook form and dialog-owner/working-area helpers; current original-source identity/access before disclosing names, paths, cached processing state or diagnostics; immutable generation pin; no original-directory listing; explicit stable-ID command scope; empty selection never means all; 1,024-ID cap; unchanged maintenance/build/host-reader behavior. All 650 other product files, including every .inky and all R9 Worker policy files, remain byte-identical.
- [x] Larger initial form (1440x960, still clamped to the current working area) and a document layout that assigns remaining height to the list. Filename first, readable status, recommended action and original folder; opaque IDs and raw per-document diagnostics are opt-in, with manual ID input retained in the advanced panel. Status explanations distinguish incomplete/unknown extraction, empty text, processing failure, missing extraction/index, required Office reader, removal and unavailable access.
- [x] Status filters default to Needs attention. Separate incomplete/unknown, empty, failed, extraction/indexing, inaccessible-source, searchable, all and removed views. Coverage is classified even when partial-text search is allowed; Searchable uses the existing inventory predicate. Empty exporter results without a representation are correctly shown as No readable text. Inaccessible-source rows never reveal their cached status through issue-filter membership; they appear generically only in All or Source access unavailable.
- [x] Find documents scans published metadata asynchronously in bounded worker chunks (up to 1000 records / one second between UI updates), skips irrelevant records before source checks where safe, and continues past empty chunks until 200 matching rows or exhaustion. Only matching authorized rows are added. Pause/close cancellation remains live. This is a filtered metadata scan, not a new persisted issue index: finding sparse issues can still inspect the entire published metadata snapshot, and one shard/source check can exceed the cooperative time budget.
- [x] Load more matches appends results and retains current row selections. Header clicks toggle ascending/descending sorting with stable-ID tie breaks; sorting explicitly applies to loaded matches. A 5000-row display cap prompts filter refinement rather than unbounded UI growth. Select loaded matches observes the 1024-command cap; Clear selection and filter edits clear stale selection. Filters cannot reset/dispose an active background iterator because controls are disabled during the search.
- [x] QA: 105 general source checks, 39 OCR diagnostics, 90 exception recovery, 62 Worker dependency/metadata, 12 ValueTuple policy and 78 focused picker/source/reference checks passed (386 total). Scope assertions were updated solely for the separately verified Form change. The old structural counter missed multiline event lambdas; it now counts lambda starts/ends and correctly handles End Sub with Task.Run arguments. No native compiler success is inferred from these static/reference checks. Exact patch bytes, previous-basis overlay, 652-file cumulative basis, diff and exact find/replace replay verified externally.
- [ ] Native gate: VS2022 Rebuild; open the common console in Word and Outlook, inspect 100/150% DPI and a smaller work area; run Needs attention, incomplete, empty, searchable (including allowed partial text), inaccessible and All queries; verify no hidden cached labels; sort each column both ways; retain selection across Load more; cancel a sparse-metadata search; verify filter edits clear selection; paste explicit IDs; confirm selected/all commands retain their existing scopes. UI rendering, compiled VB, production ACL performance and Office runtime have not been executed in this environment.

### CP-SA-R11-01 — bounded footer repair, search batching and index-independent unregistration (2026-10-05)

- [x] Input: CP-SA-R10-01 / RI_Gen2_261005_R10_cumulative_Code_inky.zip; all 652 files reconciled. RI_Tooling_Log(2).txt shows 85.817 seconds total: 44.407 seconds tool execution, 39.005 seconds outer model requests and 1.883 seconds bootstrap. First search 38.651 seconds with 8 internal model calls, 2158 authorization checks; two continuations add 2 internal calls. One unavailable-generation search and a missing-task-status repair (12.978 seconds, draft changed language) are observed. These are baseline measurements, not post-change performance claims.
- [x] Shared footer-only recovery retains one substantive missing-footer draft per run, asks the model for its completion decision only, and rejoins only a strict terminal footer-only response. Empty, tool, invalid or corrected-prose replies consume the retained draft without joining it. Word and Outlook integrate the same prepare/restore hooks before tool detection. The host never invents complete/blocked, and the combined text passes the unchanged strict parser and all normal memory, mandatory-tool, mutation, deliverable, promise and user-presentability gates. Draft is excluded from state serialization. A nonconforming repair follows the existing bounded recovery path; no answer is silently accepted or translated.
- [x] Routing planner and metadata selector now share one canonical complete JSON-string wire record. Previously planner character counts used raw compact text while selector counts included JSON quoting/escaping, so planned groups could split during the single-call selector and leave a remainder for re-preparation/authorization. Character limits now agree without truncation, lexical narrowing or raising any configured model/prompt/candidate budget. Token-budget splitting and continuation retention remain in place.
- [x] Exact-metadata matching queues only private opaque candidate IDs/scores. Original sources are no longer opened just to rank cached matches; fresh source/identity/expiry checks still precede every routing/model input and every emitted hit/reference/name/path/summary. Access, validation, requester-grant and evidence-reader implementations are byte-identical. Denied candidates are removed at final validation; they may occupy bounded private capacity until that validation and can require a further continuation. No positive ACL cache, principal broadening or exhaustive negative claim is introduced.
- [x] Coverage now reports stable UnavailableArchiveIds for the selected snapshot, ExactMetadataCandidatesQueued, ExactMetadataElapsedMilliseconds and SemanticSelectionElapsedMilliseconds. Tool guidance explicitly distinguishes publication failure from query failure and discourages a new search solely in unavailable archives unless publication changed or the user explicitly asks. NextSearchArguments retain the original continuation contract. Selector clocks count real calls, not estimates; queued counts exclude existing/returned candidates.
- [x] User's old archive cannot be repaired because its published generation has no supported routing graph. Separately, local publisher removal was correctly blocked until central withdrawal. Admin Unregister now confirms the two effects and performs existing authorized withdrawal before guarded local unregistration, independently of index readability. Store/library publisher, ACL, directory, revision and subscriber opt-out guards remain unchanged. Failed withdrawal never proceeds to local removal; partial withdrawal/cancellation is reloaded and explained after reload. Original sources and generated artifacts are retained.
- [x] Old-index status is reported as an explicit unsupported format rather than repeatedly failing the entire status view. Only the existing unsupported_semantic_index InvalidDataException code is handled; unexpected corruption remains an error. Incompatible content repair buttons are disabled, while library withdrawal/unregistration remain available. No legacy migration or unsafe force-remove is added.
- [x] WORKER_HELP.txt updated: targeted retry/extract and permissions examples, final scoped verification versus published=true, operation-ID resumption, current-format repair limits, generated output/config dependencies, improved Admin selection and withdrawal/unregister semantics. Existing instructions retained/corrected rather than replaced.
- [x] QA: 117 general source/structural checks, 39 OCR diagnostics, 90 exception recovery, 62 Worker policy/metadata, 12 ValueTuple, 78 picker contracts and 67 new runtime/unregister/help source/reference checks passed (465 total). New checks cover symmetric hooks, unchanged strict gates/security code, stale-draft refusal, actual wire-size boundary overflow, fresh authorization after revocation, protected withdrawal ordering and unsupported-index handling. Prior byte-scope assertions explicitly allow the separately tested runtime files. Native host option declarations are preserved rather than imposing Option Strict on large existing host files. Exact ZIP/member bytes, previous-delivery overlay, full 652-file Code/.inky basis, diff and exact edit replay verified.
- [ ] Native acceptance: VS2022 rebuild both hosts; repeat the same query with the same archive scope and compare elapsed/internal calls/authorization checks/new phase clocks. Force missing-footer, footer-only complete/blocked, new tool work, malformed/footer-prose correction, incomplete memory/artifact and transport/cancellation cases; confirm unchanged final gates and language. Test live access revocation/expired requester grant and bounded candidates. Withdraw/unregister a current and an old published archive; check wrong publisher, missing/wrong library, failed central/local commit and subscriber opt-out; verify no source/artifact deletion. No native compile, production performance improvement, ACL or Office execution was run here.
- [ ] Requested Documentation and more/Semantic_Archives_User_and_Admin_Guide.md is absent from the available source tree. The user offered to supply it; the attachment is needed before updating its existing contents. No replacement guide was invented or existing content assumed.


## CP-SA-R12-01 — Guide update and native search comparison (2026-10-05)

- Added the supplied original user/admin guide to Documentation and more; preserved all 30 original sections and examples, updated document selection, scoped completion, worker resumption, dependencies and old-format unregister guidance, and added search coverage advice.
- Native comparison: same question/request hash and same 519-document generation. Total 85.817 -> 117.954 seconds (+37.4%); search tools 44.407 -> 81.377 seconds; outer model requests 39.005 -> 33.243 seconds. New fourth search changes query, reroutes 39 nodes, costs 33.279 seconds, adds six new documents and repeats two. Source authorization checks 2390 -> 4540; unique hits 28 -> 30. All result pages remain partial with zero source-text evidence bytes.
- New shared search metrics confirm deployment of search changes; no draft-retention/restore diagnostic appears in Outlook. Footer repair still costs 9.419 seconds with full response output. Full raw final prose is absent from the log; cannot distinguish an older Outlook host build from a shared eligibility rejection. Verify matching rebuilt Outlook and SharedLibrary binaries before claiming footer-only repair acceptance.
- No additional product-code change or fabricated measured speedup. Existing source, worker and .inky files preserved byte-for-byte from R11. Guide/source validation, diffs and changed-only overlay checked; cumulative Code/.inky updated. Windows compilation remains unverified here.


## CP-SA-R13-01 — Explicit search intent and continuation-first tools (2026-10-05)

- Log4 native baseline: 150.528 seconds; three reformulated searches repeat 39 routing nodes each, 27 selector calls and 6506 source authorization checks. No final footer repair needed. All 519 records still complete/searchable.
- Shared tool defaults search_mode=continue. One pending state with exactly matching requested archive-ID set and literal resumes the retained query even if the model reformulates; DeferredQuery and diagnostics explicitly say that reformulation was not searched. No semantic query-equivalence heuristic.
- search_mode=new without a reference permits deliberate independent questions. Multiple compatible pending states require an exact reference unless an exact original query uniquely identifies one. New mode cannot carry a continuation. Direct SearchAsync callers preserve prior exact-query auto-resume by default.
- Operation gate, scope narrowing, library validation, original literal/query checks, reference rotation, per-page budgets, pinned generations, current source authorization and final evidence gates remain active. Word/Outlook use the same shared tool implementation. No positive ACL cache or budget increase.
- Updated user/admin guide. Cumulative Code/.inky and changed-only package preserve all prior fixes; targeted static/reference and scope/intent regression checks precede delivery. No Windows compilation or post-change latency claim.


## CP-SA-R14-01 — Interoperability documentation and chat handoff (2026-10-05)

- Added and updated the supplied Current File Format Interoperability specification against R13 persisted schemas and retrieval implementation; current versions unchanged. Added section-index envelope/byte interpretation, continuation intent, runtime coverage, extraction/maintenance distinctions, and database-backend boundary. All existing sections retained.
- Latest native log5 validates same-query continuation without rerouting: 75.032 seconds, 46.910 seconds search, 11 internal selector calls, 2382 source authorization checks, 24 unique hits, no failures or footer repair. Query-reformulation deferral and ambiguity/new-mode branches remain static/reference-tested, not natively demonstrated by this run.
- Created standalone German handoff with release chain, all retained fixes, measured logs, architecture/security invariants, open verification and future scale/database decisions. No additional product code changes. Keep R13 behavior provisionally; benchmark larger corpora before new optimization or backend migration.
- R14 cumulative Code/.inky preserves all 653 R13 files except appended Tasks; adds the supplied specification. Changed-only documentation package, diff and checks prepared. No Windows compilation claimed.


## CP-G2-R15-01 — Passive Outlook phishing assessment (2026-10-05)

- Input: supplied RI_Gen2_261005_1759.zip / CP-SA-R14-01, plus mdeditor.html. Reconciled current WIP against the immutable input ZIP; CONTRIBUTING and prior cumulative Tasks retained. No prior source changes or .inky artifacts discarded.
- Cause: no dedicated passive phishing command or configurable prompt lifecycle existed.
- Change/files: added ThisAddIn.Commands.CheckPhishing.vb and its Outlook project registration; Outlook Commands/Ribbon1/Ribbon1.Designer; SP_CheckforPhishing in SharedContext, Constants, LoadConfig, all four Settings load/default/display/apply paths and Word/Outlook/Excel ThisAddIn.Properties wrappers. Button immediately above Mail Mover.
- Evidence: selected/current mail body, sender, Unicode/ANSI transport headers, literal href/src/action targets with display text, attachment filename/type/size metadata; no link fetching/resolution, attachment saving, document opening, execution or tool loop. CustomForms plain-text assessment uses INI_Language1, localized risk/confidence, explanation, limitations and next steps. Missing evidence and invalid/oversized model output fail explicitly; confidence is not presented as a calibrated probability.
- QA: source contracts and independent lifecycle/ribbon/call-path inspection passed. Strict fully qualified new system types; ToolExecution=False and absence of retrieval/execution APIs verified. Native model/Outlook/CustomForms execution remains open.
- Invariants: configurable prompt mirrors SP_CheckforII definitions; mail is untrusted evidence; attachments are metadata only. Next: complete delivered-candidate processing.

## CP-G2-R15-02 — Review every delivered M365 mail candidate (2026-10-05)

- Input: CP-G2-R15-01. Affected file: Outlook ThisAddIn.M365SearchForm.vb.
- Cause established in source: phase 1 restricted raw tool hits to the model's candidate_refs/selected identifiers before full-text review; a small returned subset could reduce 300 delivered hits to 25. No native trace was supplied to prove the exact observed 25/300 run. Graph full-text fetching already iterates all 20-request batches and was not a fixed 25-item cap.
- Change: raw successful search-tool payloads own the candidate inventory. Deduplicate by existing stable identities and retain all delivered distinct candidates, including delivered pools beyond the previous 300 merge cap. Review every candidate in existing 12-mail batches; retry individual missing batch bodies before internet-message-ID fallback, propagate cancellation. Summary also uses 12-mail batches and cannot remove phase-2 positives from the grid.
- Invariants: phase 2 remains the relevance filter, so 300 reviewed candidates do not necessarily mean 300 relevant grid rows. Existing search/retrieval limits and 6000/8000-character body budgets remain; unavailable full text and capped body evidence are explicitly identified. No claim of exhaustive attachment search or unlimited full-mail evidence.
- QA: source contracts, independent full pipeline inspection and boundary reference cases 0/1/12/13/25/26/299/300/301/500/1001 plus overlapping duplicate pools passed. Cancellation wiring and summary-positive retention verified. Native real 300-mail run remains open. Next: complete nonlocal export coverage.

## CP-G2-R15-03 — Force Outlook downloads and supplement M365 server-only exports (2026-10-05)

- Input: CP-G2-R15-02. Files: Outlook PSTConverter, new PSTServerExport, Outlook project; new SharedLibrary M365Service.MailExport and SharedLibrary project.
- Cause: original export counted/snapshotted only Outlook's local Items collection and did not explicitly request downloads. MarkForDownload/synchronization alone cannot enumerate messages outside the offline cache interval.
- Change: mark header-only items for full download, synchronize selected folders before count/snapshot, wait for completion/error/cancellation with timeout, restore original sync membership. One marking error does not prevent the remaining folder synchronization. Reject still-incomplete native mail bodies rather than writing misleading empty TXT.
- For account-backed Exchange/Microsoft-365 stores, enumerate the selected server folder with every @odata.nextLink, translate REST IDs to exact binary MAPI EntryIDs, match the selected Outlook delivery account and signed-in M365 mailbox, and supplement missing/header-only/failed native mails by downloading server bodies and attachment bytes. Exact parent-name traversal rejects ambiguous/unresolved server folders. Dedupe repeated server page identities. Keep successfully exported full local items on the original export path; no subject/InternetMessageId heuristic deduplication. Existing shared attachment extractor, folder mirroring, inline/separate text, placeholders, index/error logs and cancellation retained. Cloud-reference attachment targets are never fetched.
- Invariants: local PST stays offline and unchanged; no 1000-record inventory cutoff; identity conversion must be complete; incomplete coverage/download failures are visible, including when no local items exist. Server enumeration is not a transactional mailbox snapshot: concurrent moves/deletions can yield explicit item failures.
- Dependencies/open points: complete Graph coverage requires existing INI_M365ClientID configuration, appropriate delegated Mail.Read access and sign-in to the selected account. On-premises, unmatched shared/archive stores, missing Graph configuration/consent and unresolvable folders are explicitly reported as unverified coverage, not silently called complete; native synchronization still runs. Child recursion follows the selected Outlook folder tree. No new account selection is guessed.
- QA: source routing/invariants, independent identity/download/export review, 100-record pagination boundaries through 1001, all base64 padding classes (256 binary round trips), local/server union/dedup reference cases and cancellable stream wiring passed. Real Outlook/Graph/offline-cache export remains a required native test. Next: shared persistent Markdown widget.

## CP-G2-R15-04 — Shared modeless WebView2 Markdown Editor (2026-10-05)

- Input: CP-G2-R15-03 and supplied mdeditor.html. Files: SharedLibrary MarkdownEditorForm.vb, embedded MarkdownEditor.html and project; Word/Outlook Ribbon1 + designer, Word TalkToMe command registration, Outlook dispatch.
- Cause: supplied browser editor had no Office host, durable native path binding or on-screen compact window.
- Change: Word Helpers and Outlook open a modeless shared editor class, with independent per-host persistent WebView2 profiles and singleton/process lease. Red Ink icon in native title frame; native minimize/Compact creates a movable small on-screen editor window, including a repeated minimize of the already compact window. Position, topmost/compact state, selected folder and opened native filenames/hashes persist per host. Markdown/text native Open, Save/Save As, explicit Reload and drag/drop supported; paths require native user selection or WebView2 File objects. Atomic writes and external-change conflict checks; close waits for browser note storage/folder writes to confirm completion.
- Preserve original editor/parser, settings, themes, notes, undo, search, images, exports and folder workflow; bridge native opaque directory handles so WebView2 does not require unsupported browser folder permissions. Selected-root containment, reparse-point rejection, strict text decoding, recycle-bin deletion and write-conflict checks applied. Notes and imported file snapshots survive reopening without silently overwriting newer editor state. Linked files write on explicit Save/Ctrl+S; browser notes autosave; selected-folder notes use existing direct folder autosave.
- Safety/regression invariants: no Office window made a modal owner; every native modal path uses InspectDialogOwner + IfOwnerOnCurrentThread. External automatic resources/navigation blocked, permissions/host objects denied, only explicit user-initiated supported links can open externally after confirmation. Untrusted raw note HTML uses a DOM allowlist/decoded-attribute sanitizer so it cannot invoke the privileged bridge. Browser failure/save failure remains visible. Existing selftest-only code was excluded from the product HTML.
- QA: all executable script syntax passed; actual Markdown core compared against input across 28 samples/16 highlighting languages and safe-URL cases (108 checks), and parser source outside sanitizer was byte-equivalent after normalization. Actual native-bridge JavaScript tested with RPC stubs for Unicode/UTF-16/invalid UTF-8, binary boundaries and directory messages (34 checks). Project/resource/profile/dialog/persistence contracts and independent source review passed. RPC stubs are not native filesystem/browser tests.
- Failed QA attempts: Playwright launch was unavailable (initial absent /tmp, then absent Chromium executable); no browser rendering, screenshot or UI PASS claimed. Temporary syntax harness initially included text/markdown script data; Linux case-sensitive project inventory initially flagged the input's ThisAddIn/ThisAddin filename casing; broad prompt test mistakenly included Word's existing II command invocation; initial external-resource assertion included an embedded data favicon. These harness assumptions were corrected and rerun; they were not product defects. All harnesses remain outside product source.
- Native open: VS2022 rebuild, WebView2 runtime installation/profile creation, real DOM sanitizer, window positioning/DPI, Word/Outlook simultaneous use, drop/import/save/reload/conflicts/folder operations/exports, minimized icon visibility, storage errors and host shutdown. Next: complete third-party notices.

## CP-G2-R15-05 — Permissive icon/palette notices (2026-10-05)

- Input: CP-G2-R15-04. Files: root and all four host/shared license.txt files; Word/Outlook/Excel ThisAddIn license headers; SharedLibrary Source Code Overview and Constants license list; embedded editor HTML.
- Cause: supplied short note omitted Lucide's Feather-derived MIT portions and individual palette copyright notices.
- Change: added Lucide ISC + Feather MIT and adapted Nord/Catppuccin/Tokyo Night/Dracula/Gruvbox/Solarized MIT entries immediately before Microsoft in every existing list; full permission/disclaimer texts and upstream copyright/source notices in distributed license.txt files and embedded/exported HTML. Verified Tokyo Night's 2018-present Enkia notice and Dracula's 2023 notice against actual upstream LICENSE files; corrected intermediate generic/current-year attribution before packaging.
- All introduced icon/palette upstream dependencies are permissively licensed; no additional parser, font, package, CDN or theme implementation bundled. Replaced two Obsidian-labelled palettes with independently authored Red Ink palettes and stable preference keys; no proprietary Obsidian/Dracula PRO assets. Existing WebView2/Office/Microsoft licensing retained.
- QA: primary upstream license/source inspection, source dependency scan and independent notice/list-placement checks passed. No legal certification of all historic dependencies or the unknown original authorship of supplied self-contained editor code is claimed. Next: package verification and native acceptance.

## CP-G2-R15-06 — Cumulative source packaging and outstanding native acceptance (2026-10-05)

- Input: CP-G2-R15-05. Tasks appended without removing any prior checkpoint; original code and .inky outside the declared delta unchanged. Input ZIP contains 652 files; cumulative output contains 657. Current scope: 31 changed/new product files including Tasks, with five new source/resource files.
- Completed local QA: 1652 source/inventory/contract/reference assertions, 108 actual Markdown-core checks and 34 actual bridge-code checks; independent changed-file/project/resource/parser preservation inspection. These counts include source-preservation and inventory checks and are not 1794 end-to-end Office tests. Exact baseline-to-final find/replace replay, changed-only ZIP overlay and cumulative ZIP member/byte verification performed before delivery; QA/diffs/reports are separate deliverables, not project inputs.
- [ ] Native acceptance: rebuild SharedLibrary and all three Office hosts in VS2022; run Word/Outlook new commands and verify no new dialog-owner/watchdog anomalies. No Windows/VSTO/COM build was executed here.
- [ ] Phishing: benign, phishing, payment fraud, sender/reply-to/authentication mismatch, deceptive literal URLs, punycode, dangerous attachment names, missing headers/body, invalid model response, user config prompt override and multiple configured output languages. Confirm no URL fetch, DNS resolution, attachment open/download/execution by this command.
- [ ] M365 search: real 300 delivered distinct hits, model returns only 25 candidate refs, all 300 inspected through final batch; all phase-2 positives retained even if summary returns fewer refs. Also cancellation, >300 delivered, duplicates, missing/throttled full text and zero matches.
- [ ] Export: selected M365 account with narrow OST cache, 300+ server-only messages, all Graph continuation pages, header-only native mail, attachments and placeholders, cancellation/timeout/offline mode, wrong signed-in account, insufficient consent, shared/archive/on-premises failures, native-only nonmail items, PST unchanged and index/output accounting. Ensure full local native exports retain original metadata/extraction behavior.
- [ ] Markdown: real WebView2 rendering/sanitizer, drag onto page and compact window, persistent open paths/folder/profile across restart, external file edits/conflicts, strict encoding failure, browser storage/write failures, folder rename/delete/image workflow, HTML/ZIP/print exports, all DPI/screens and simultaneous Word/Outlook host shutdown. UTF-8/UTF-16 BOM text supported; unsupported encodings fail explicitly rather than silently replacing bytes.
- Delivery is a source implementation with completed local QA and explicitly pending native acceptance, not a claimed native-tested release. Next intended step: run this acceptance matrix on the Windows development installation and reconcile any failures into the cumulative Tasks checkpoint.


## CP-G2-R16-01 — Independent delivery review and corrective implementation (2026-10-05)

- Input: delivered CP-G2-R15-06 cumulative ZIP; all 657 working files verified byte-identical to that delivery before review. Re-read CONTRIBUTING; preserve all prior Tasks checkpoints. Files: Outlook ThisAddIn.M365SearchForm.vb; SharedLibrary M365Service.MailExport.vb, MarkdownEditorForm.vb and MarkdownEditor.html.
- Found causes: cached header records could bypass full-body retrieval; malformed/empty model replies silently counted as no matches; server export accepted null/malformed bodies. Browser storage fallback reported successful persistence despite blocked/quota failures and could read stale persistent text; async Reload could overwrite a switched note or new typing. Metadata polling advanced native conflict hashes without loading the corresponding text. Case-only folder rename fell through copy/delete on Windows. Folder Save As could bind an export copy to the active folder note. Read/hash used different file snapshots, and disconnect cleared the UI before native completion.
- Corrections: validate raw body.content as a string on batch/cache/individual retry and server export; valid empty body remains valid evidence without preview substitution. Reject invalid model arrays and out-of-batch references; valid [] remains no relevant matches. Keep failed browser writes/deletions in memory with retryable pending state and report unsuccessful durable flush. Bind Reload to original note/text and acknowledge its native snapshot only after application. Native folder reads acquire a separate snapshot hash; only accepted text advances the loaded conflict baseline. Native rename checks the loaded baseline, uses a temporary move with rollback for case-only names, and never falls through copy/delete after native failure. Reject reserved/trailing-dot-or-space Windows names. Save folder exports as unbound copies. Decode strict text and hash from one byte buffer. Keep failed-write folders connected, await native disconnect before clearing, preserve unsaved/current notes during scans and show read/decode failures.
- Invariants: passive phishing restrictions, prompt lifecycle, all-delivered search pool and retained relevant results, native/Graph export accounting, existing parser/theme/licenses, host ribbon/resource registration and modeless editor ownership remain unchanged. No new product dependency or source file. The R15 native acceptance matrix remains open.
- Completed local QA: original full suite rerun: 1652 source/inventory/contract/reference assertions, 108 actual Markdown-core comparisons and 34 actual JavaScript bridge checks. Added 50 actual JavaScript failure/race checks for blocked/quota storage, pending recovery, switched/edited Reload, passive scan, same-timestamp hash changes, failed/native case-only rename, failed disconnect/invalid decoding and unbound folder export; added 15 separately inspected source contracts for body evidence, native snapshots, strict encoding and rename/disconnect boundaries. RPC/file/storage stubs are not native Office, filesystem or DOM tests. No Windows compiler/runtime available; no native build, Graph request or model execution PASS claimed.
- Next: byte-verified R16 correction-only package against R15, cumulative source and exact Find/Replace review; then Windows acceptance from CP-G2-R15-06 including the new failure/race cases.

## CP-G2-R16-02 — Corrected source delivery (2026-10-05)

- Input: CP-G2-R16-01 after associated QA passed. Delta from R15: four implementation files plus cumulative Tasks (five files); total remains 657. Original CONTRIBUTING and .inky preserved. No QA scripts, tests, patches or reports added to product projects.
- Packaging harness attempt: an initial baseline filename substitution referenced a nonexistent R16 archive; failed before packaging/writes. Corrected the harness to the immutable R15 baseline and reran successfully. No product defect or native PASS inferred.
- Packaging validation: exact unique Find/Replace reconstruction with BOM/line endings preserved; correction ZIP overlay over immutable R15 equals final source; cumulative ZIP exact inventory/member bytes and archive integrity checked. Review artifacts include declared baseline hash and per-file SHA-256. R15 delivery archives retained unchanged.
- Open: all native Windows/Office/WebView2/Graph/model acceptance from R15, plus actual quota/blocked profile, external write conflict while polling, case-only rename, note-switch/typing during Reload, invalid text and failed folder disconnect. Delivery contains reviewed and corrected source, not a native-tested release.


## CP-G2-R17-01 — User screenshots and passive structured phishing report (2026-10-05)

- Input: delivered CP-G2-R16-02 and three supplied screenshots. Files: Outlook ThisAddIn.Commands.CheckPhishing.vb. Evidence: compact controls visibly cropped; phishing JSON parse fails on leading backtick. No actual model/Office execution was available here.
- Cause/change: direct JObject.Parse rejected fenced/encoded JSON. Normalize surrounding text/fences and JSON string envelopes, validate all risk/confidence/text/list/localized heading fields, then perform at most one format retry with the original evidence and tooling disabled. Invalid final results stay explicit failures. Explorer selection must contain exactly one item; multiple selection exits before evidence/model calls. Snapshot subject, sender and received date before await and render them with the assessment.
- Display: existing CustomForms HTML viewer; deterministic color/symbol hero card, localized risk/confidence and explanation, findings, practical actions and limitations. Encode every mail/model value as plain HTML text. No generated links, images, scripts or external styles; passive URL/attachment restrictions preserved. Model content never becomes markup.
- QA: updated complete static/passive/prompt/project suite plus separately inspected source contracts passed. Native model-format recovery, localized reports, HTML rendering and selection behavior still require Outlook acceptance. Next: floating editor and host recovery.

## CP-G2-R17-02 — Floating Markdown document, menu glyph and recovery snapshots (2026-10-05)

- Input: CP-G2-R17-01 and both compact-size screenshots. Files: shared MarkdownEditorForm.vb/MarkdownEditor.html, Word/Outlook Ribbon1.Designer.vb and ThisAddIn.vb.
- Cause/change: old 240x75 decorated compact window left too little client space for the toolbar. Replace compact mode with a DPI-aware 72 logical pixel transparent, borderless floating document glyph; click restores, drag moves, file drop imports/restores. The glyph still technically uses a WinForms window. Restore normal frame/WebView/auto-sized toolbar and clamp to the working area. Preserve Red Ink logo in normal title frame; Word Helpers/Outlook menus use an independently drawn .md document image with no new third-party assets/licenses.
- Host-close behavior: browser autosaved notes remain in the per-host persistent profile; linked source files still require Save/Ctrl+S. Every text change sends a native recovery snapshot. Collect separate pending snapshots per note; atomically persist with a 200ms native timer, and synchronously checkpoint latest acknowledged snapshots plus window/path state at Application.Quit and VSTO Shutdown. Outlook Shutdown alone is insufficient, as documented by Microsoft. No blocking async wait/pump, no implicit overwrite of linked files.
- Recovery reopens differing snapshots as separate notes, never overwriting newer profile/source content. Matching native snapshot clears only after corresponding browser/folder storage succeeds; restored snapshots clear only after durable browser confirmation. Keep snapshots on blocked storage and retain unread/corrupt recovery files. Report/log recovery storage failures. Switching among multiple failed/unsaved notes keeps all pending snapshots.
- Invariants: existing Markdown core, settings/themes/images/folder workflow, R16 storage/race/conflict/rename fixes and dialog-owner rules retained. Forced kill/crash can lose edits not yet delivered to the native host or not yet committed by the recovery timer. Quit checkpoint cannot guarantee those undelivered keystrokes; no claim that Office exit is cancelable here.
- QA: 21 actual JavaScript recovery checks with storage/native message stubs, 50 prior storage/race checks, 34 bridge-code checks, 108 Markdown-core comparisons plus independent source inspection passed. Native transparent window/DPI/drag/drop, quit timing, atomic filesystem writes and simultaneous hosts not executed. Next: enforce total M365 retrieval target and page Graph correctly.

## CP-G2-R17-03 — Total candidate target, +100 paging and Graph page contract (2026-10-05)

- Input: CP-G2-R17-02. Files: Outlook ThisAddIn.M365SearchForm.vb; shared M365Models.vb, M365Service.vb and M365ToolService.vb.
- Additional cause: the shared Graph search/query implementation performed one request with size up to 500, but message/event result pages are limited to 25. Fix the shared per-source page loop, not just the Outlook result display. Preserve existing fields/query/date filtering, cancellation and canonical metadata paths.
- Contract: UI default 50, maximum 2500 distinct candidate mails in total; all candidates in that bounded pool are reviewed and all accepted positives remain in the grid. This explicit new user request supersedes R15's unlimited delivered-pool rule. Candidate target is not a guaranteed result count: fewer available candidates or relevance filtering yield fewer shown mails. Full-text budgets and missing-evidence diagnostics remain.
- Scoped host-neutral single-source AsyncLocal budget isolates this interactive harvest from unrelated Word/Outlook tools and serializes concurrent calls. Enforce the caller's source/budget and actual per-query cursor, first page at zero; generic tool default stays 25. Tool response exposes actual continuation/exhaustion. Host pages recorded queries to fill the distinct target even when model harvesting stops early. Get more increases preset by 100 up to 2500 and retains existing candidates; cancellation/errors remain visible. All relevant results survive final summary. UI target captured before asynchronous work to avoid background NumericUpDown access.
- Shared Graph mail/event pages at most 25 (other search types keep prior 500-page ceiling); total request bounded to 2500. Deduplicate page identities, advance cursor by raw count and retain continuation separately from unique count. Reject missing/invalid continuation, oversized/repeated pages; never count a scoped search error as successful zero matches. Actual runtime requests and Graph behavior remain pending native tests.
- Primary references verified: https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview?view=graph-rest-1.0 (from/size, message/event page25) and https://learn.microsoft.com/en-us/visualstudio/vsto/events-in-office-projects?view=visualstudio (Outlook Shutdown vs Application.Quit).
- QA: 59 independent source/contract/reference checks including boundaries 25/26/50/150/300/500/2499/2500 and raw-cursor dedup cases; full existing 1652 source/inventory/contract/reference suite and editor checks rerun successfully. The previous unlimited-pool harness assertion initially failed after the intentional target change; updated to the new explicit target invariant, then passed. This was an obsolete harness expectation, not an unexplained product failure.
- Open native acceptance: real 2500-mail Graph paging, repeated queries/dedup/short pages, cancellation/throttling/error coverage, 50 -> 150 -> 250 growth and all relevant rows. Scope flow through the real tooling loop requires runtime confirmation. Next: correction-only packaging and cumulative basis.

## CP-G2-R17-04 — Reviewed source package and cumulative basis (2026-10-05)

- Input: CP-G2-R17-03 after associated verification passed. Product delta from R16: eleven implementation files plus cumulative Tasks (12 files); total still 657, no added product file or dependency. R15/R16 fixes preserved except explicitly superseded candidate-limit policy. CONTRIBUTING and .inky bytes unchanged; QA/diffs/reports outside product projects.
- QA/delivery: exact unique Find/Replace replay against R16, correction-only ZIP overlay, cumulative source inventory and member bytes, archive integrity and per-file SHA-256 verified. Changed source files are the primary ZIP; a full tar.gz source/.inky basis is provided separately. Review ZIP has Find/Replace (separate blocks with exact Tasks placement anchor), diff, manifest and notes. Prior deliverables remain immutable.
- All native gates in CP-G2-R15-06 remain, supplemented by R17 scenarios above. No Windows/VS2022/VSTO/Office/WebView2/Graph/model build or runtime PASS claimed. Local QA counts describe source/reference and JavaScript stub checks, not end-to-end tests. Next: native rebuild and acceptance with the supplied user screenshots/behaviors.


## CP-G2-R17-05 — Final adapter integration review and repack verification (2026-10-05)

- Input: CP-G2-R17-04 before release. Files: M365ToolService.vb, Outlook ThisAddIn.M365SearchForm.vb and cumulative Tasks.
- Additional actual wiring cause: existing Outlook M365 adapter passes ct:=Nothing to the shared tool, so a paging loop could ignore the UI cancellation token. Keep the scope caller-neutral: retain its token and use it whenever the adapter supplies no cancellable token. The same shared scope API is usable by Word. Validate gathered search failures at the form boundary instead of interpreting the tooling loop's skipped failed search as successful no-match coverage.
- Invariants: no broad host adapter rewrite, unchanged unrelated tool calls, scope/context isolation, strict total candidate target and symmetrical shared behavior. No Await in Catch; cancellation reaches Graph/page semaphore from the scoped caller, and pipeline checks token before accepting completed snapshots.
- QA after these final edits: full 1652 source/inventory/contract/reference suite, 108 Markdown-core comparisons, 34 bridge checks, 50 prior failure/race checks, 21 recovery checks and 61 independent source/reference checks passed. Added explicit adapter-token and failure-coverage contracts. This is source/stub verification; real ExecutionContext/scope/cancellation behavior remains a native gate.
- Rebuilt and byte-verified all final packages after this correction, including cumulative tar.gz source/.inky basis. Next: native acceptance; no unverified source changes included.


## CP-G2-R18-01 — Linked-file autosave with persistent controls (2026-10-05)

- Input: user follow-up to R17. Files: shared MarkdownEditor.html and MarkdownEditorForm.vb. Existing Autosave switch persisted editor/folder notes but did not write native linked originals.
- Change: retain that switch as Autosave notes; add a separate Autosave linked files switch, off by default, in the existing per-host browser settings. Both values round-trip through the existing persistent Store; report failed persistence. Word and Outlook use the same editor implementation and separate existing profiles.
- Debounced, serialized writes capture note id/text, preserve pending writes across note switches, coalesce newer edits, stop a conflicted note, and resume after explicit successful Save/Reload. Native autosave requires an existing linked original, never opens a picker or recreates a deleted original, and retains the existing external-fingerprint check and atomic writer. Explicit Save requeues typing received while it awaited older writes; Reload awaits an in-flight write and rejects changed note/text. Disabling stops pending writes; an already-started write may finish. Close awaits the same queue; host Quit still uses the synchronous acknowledged recovery checkpoint rather than waiting for asynchronous browser work.
- Invariants: folder-native handles remain on the existing folder workflow; parser, themes, license notices and profile/recovery durability rules remain. A recovery snapshot can clear after successful browser note persistence; no claim that both stores always retain identical copies. Recovery text/profile remains available after a source write conflict. Neither autosave nor shutdown guarantees keystrokes not delivered to the native host before a forced process kill.
- QA: actual shipped JavaScript queue/settings code executed against storage/RPC stubs, including on/off, persisted true/false, note switching, conflicts, coalescing, concurrent typing and close-drain behavior (29 checks). Prior bridge/storage/recovery/core suites passed. Initially isolated legacy harnesses lacked the new RILinked dependency and the new defaults harness lacked FONT_PRESETS; these harness setup errors were corrected and rerun. No native disk/UI execution claimed. Next: report triage and shared transport/operation corrections.

## CP-G2-R18-02 — Current-code report triage and request-scoped transport recovery (2026-10-05)

- Input: frozen RedInk_Host_Issues_2026-10-05_inkl_Belege.zip (all five proof texts, report, README, simulation and minimal DOCX inspected). Report code references are September 29, while this basis is October 5/R17. The attachment remains unchanged; reported live observations are considered evidence, not proof of the proposed universal fix.
- Report 1 deferred: the script's variable named Word is a simulation keyed by abstractNumId, not execution of Microsoft Word. A global abstractNum counter change is not justified by this oracle alone. Keep the current DOCX numbering reader; require actual Word-rendered/reopened results for the minimal case and independent lists, style links, nested restarts and startOverride fixtures. This does not dismiss the reported incorrect clause references.
- Report 2 implemented in shared MainLLM/transport scope and both host tooling loops: isolate one unattended logical LLM request's failure budget from the mutable parent/other concurrent calls; preserve bounded retries, backoff and caller cancellation. Resolve the requested profile into that scope, dispose both scopes on every exit, and report the last effective attempt timeout instead of only the configured timeout. Do not reset a shared stopwatch on another worker's success.
- Add scoped, host-neutral diagnostic sink. Both hosts send safe numeric status/attempt/backoff/remaining-budget diagnostics to their existing tooling logger for HttpClient POST, fallback POST and fallback GET. nextDelayMs=-1 means the failure has been observed before a retry-delay decision. Logging failure cannot alter transport behavior; no endpoint/header/prompt/body is included. Child scopes restore the previous sink.
- Report 2 delivery proposal deferred: failure alone cannot certify that a produced artifact is final, complete, authorized and ready to send. Existing finalization/delivery gates remain authoritative; no automatic partial-result sending or false completion introduced.
- QA: source contracts and independent scope/budget reasoning checked, existing transport limits/profile defaults preserved, effective timeout diagnostics verified in both branches. Actual AsyncLocal concurrency, 429/5xx/Retry-After, host timeout and log routing require Windows/live transport tests. Next: explicit operation result binding.

## CP-G2-R18-03 — Recorded-result replay without changing explicit identities (2026-10-05)

- Input: CP-G2-R18-02 and report 3. Files: shared ExplicitOperationRegistry.vb and both host tooling loops. Confirmed: old top-level succeeded guard could claim synthetic success for a different tool using the same operation_id+step_id and did not return the actual original output.
- Change: preserve the explicit identity key and monotonic terminal rules. Store the successful tool, canonical argument SHA-256 and original response as validation evidence, then replay only when every requested identity succeeded and the entire tool/call binding matches. Sort object properties recursively; retain array order and scalar values. Changed tool/arguments, incomplete batch or unbound historical record returns an explicit operation_replay_conflict, no execution or invented output. New distinct continuation uses a new step_id; a failed/completed step must not be re-executed by changing the registry key.
- Capture bindings only from host-confirmed successful logical outcomes; preserve top-level capability gating and existing per-task applied semantics. Export/import journal records retain the new fields; old records default empty and cannot authorize replay. Limit retained original responses to 65,536 characters per record; larger records retain success state but reject unverified replay rather than fabricate success. Replay represents the original execution, not fresh proof that the output still exists.
- Report's blanket mixed-batch claim is partly outdated: the current TryFilterTerminalTaskOperations already removes succeeded/terminal task items from ordinary independently identified batches while pending tasks continue. Preserve it. A top-level operation identity intentionally prevents splitting that logical call; mismatched/incomplete such batches are explicitly rejected with instructions to resubmit pending tasks. All-completed identical bound batches can replay the complete recorded result.
- QA: symmetric host guards, all-identity checks, canonicalization and bounded journal round-trip checked; independent reference cases cover reordered object keys, different tool/output, mixed statuses and historical records. No VB.NET/real tool execution here. Native acceptance must include the supplied pdf_to_word -> create_word_document collision, identical retry, restored journal, mixed per-task batches and top-level batch conflict. Next: opt-in Markdown export and Outlook restoration.

## CP-G2-R18-04 — Opt-in structured Markdown export (2026-10-05)

- Input: report 4. Files: shared TextExportService.vb, TextTools.SemanticAndExport.vb, SharedMethods.PdfMarkdownExtractor.vb. Add output_format=text (default) or markdown to the existing shared export tool/options.
- PDF/DOCX/DOCM/MD Markdown requests append .md to the source filename; other formats keep the existing .txt extraction with actual content_format metadata. Single files and recursive directories use the same snapshot/single-flight/provenance/atomic publication boundary. Do not call the direct convenience writer, which would bypass those protections. No model copying and no new dependency.
- DOCX uses the existing Markdown/coverage reader; numbering remains unchanged pending report-1 validation. PDF uses the existing layout renderer with a new explicit reader-status overload; preserve the original public two-argument API and its string outputs. Source text beginning with Error: is not interpreted as an error. The adapter consumes the separate status. Image-only/insufficient-text PDFs fail explicitly; PDF Markdown + ocr_pdf=true returns unsupported, never falls back silently or labels OCR/plain text as Markdown. Layout heuristic coverage/page association stays unverified.
- Preserve the exact legacy text options fingerprint; only opt-in Markdown adds its format discriminator, preventing cross-format session reuse without invalidating existing Semantic Archive derivatives. Retain overwrite=false precheck, source permission/path policy, captured-source hashing, cancellation-before-publication, existing counts and structured failure metadata.
- Report 5 deferred: empty paragraphs alone do not prove empty pages; removing them can affect spacing, styles, table requirements or pagination. No blanket cleanup or automatic trim added. A separate explicit structural operation needs Word layout fixtures and clear retained-paragraph rules.
- QA: format schema/defaults, reader status/API compatibility, filenames, format-isolated fingerprints and unchanged publication protections independently inspected (part of 66 source/reference checks). Real PDF/DOCX output, scanned PDF refusal, directory collisions/overwrite, ACLs and cancellation remain native acceptance. Next: restored editor state and final delivery verification.

## CP-G2-R18-05 — Restore Outlook editor at visible saved coordinates (2026-10-05)

- Input: further user instruction during R18. Files: shared MarkdownEditorForm.vb and Outlook ThisAddIn.vb. Shared restore helper is reusable by either host; only Outlook opts into automatic startup reopening, as requested.
- Persist Reopen=1 with the host-close/window/path checkpoint before separate recovery I/O. A successful explicit editor close saves state then clears Reopen; a failed close keeps the window/reopen intent. Host Quit marks the closing path so no late asynchronous modal/flush handshake is attempted. On deferred interactive Outlook startup, schedule restoration on the existing UI control once; avoid a second instance or enlarging an already-open compact editor. Startup errors are logged and do not break Outlook initialization. Headless/no-Explorer startup does not open an editor.
- Retain normal bounds separately from compact icon coordinates, plus current compact/maximized/topmost state. Existing per-host profile restores notes/settings and the selected note according to the existing Open last note preference. Legacy compact settings fall back to their raw stored X/Y. Read current screen WorkingArea and clamp size/position against it, including disconnected secondary monitors, negative coordinates, smaller resolutions and compact-icon placement. Normal minimum size is capped to the available working area.
- QA: independent geometric cases and source lifecycle/order checks passed; full 1652 static/source/reference suite plus 108 core, 34 bridge, 50 previous race, 21 recovery and 29 new linked/settings JavaScript checks passed. Additional R18 independent contracts/reference suite: 66 checks. JS execution used RPC/storage stubs; no DOM/native I/O/Office/.NET compiler was available.
- Open: Windows VS2022 rebuild and Word/Outlook acceptance for linked autosave switches/profile persistence, queued/conflicting writes, normal/compact/maximized editor startup, multiple monitors/DPI/resolution changes and host/manual close; unattended transport/scoped diagnostics/concurrency; registry retry/journal/collision/mixed batches; Markdown PDF/DOCX/scan/OCR/overwrite/resource reuse. Earlier R15/R16/R17 native gates remain. Next: correction-only R17-to-R18 package with exact Find/Replace and full source/.inky basis, verify archives before delivery.


## CP-G2-R18-06 — Explicit editor-button recovery and delivery verification (2026-10-05)

- Input: user steering during final review: clicking the editor button again should recover a lost open editor. File: shared MarkdownEditorForm.vb; the existing Word Helper and Outlook buttons both call this shared path.
- Change: for an existing instance, leave compact mode, restore a normal window, fit/center it in the current cursor monitor's working area, Show/BringToFront/Activate, and temporarily raise it above other windows while preserving its Always on top preference. The separate automatic-startup helper leaves an already-open instance untouched, so startup does not recenter/enlarge a compact editor. No modal Office ownership added.
- QA: button/placement source checks plus independent positive/negative-coordinate/small-screen centering cases; R18 independent contract/reference count now 71. Full prior source/core/bridge/race/recovery and 29 new linked/settings checks passed. First R18 package verified 12 changed files, 657 total files and 62 unique Find/Replace blocks. Final packages are rebuilt and byte-verified after this additional request; prior R17 archives and the report remain immutable.
- Open native gates from CP-G2-R18-05 remain, now including repeat-click rescue from compact, maximized and off-screen windows, across Word/Outlook and multiple monitors. No Windows/.NET build or real foreground activation claim. Exact corrections are against R17; cumulative source/.inky tar.gz is the new full basis. Next: native rebuild/acceptance with these artifacts.
