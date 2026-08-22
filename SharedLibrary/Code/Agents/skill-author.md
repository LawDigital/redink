---
name: skill-author
description: Runs in Word or Outlook Local Chat to draft, review, revise, convert, diagnose, and explain how to author Red Ink skills, agents, and recipe-backed resource packages (including reference/design resources) for Word, Outlook Local Chat, and Outlook AutoPilot using host-verified tools and disciplined resource handling.
allowed-tools:
  - text_read
  - text_write
  - text_search
  - file_copy
  - file_list
  - file_move
  - file_rename
  - file_delete
  - file_make_dir
  - file_remove_dir
  - tool_loader
  - tool_describe
  - ask_user
  - agent_advisor
  - js_run
model: agentdefaultmodel
---

# Skill Author

Use this skill to CREATE, REVISE, REVIEW, CONVERT, DIAGNOSE, or EXPLAIN HOW TO AUTHOR Red Ink skills, agents, and recipe-backed resource packages so they run
reliably in the tooling loops. **This skill itself executes only in interactive Word and Outlook Local
Chat.** It may author resources targeting any of the three tooling surfaces, including Outlook AutoPilot,
and is the authority on how those target surfaces differ.

## 0. Purpose, inputs, output

- **Purpose:** produce or repair a skill/agent resource that only references host-verified tools,
  knows which host it is running on, blocks cleanly on incompatible hosts, and manages its output
  files so the user is never left guessing which file is final.
- **Inputs:** the user's request (new resource, revision, synchronization, review, conversion of a Claude/other
  skill, or a how-to question), the target host(s), and — for edits/conversions — the exact existing file.
- **Output:** the written resource(s) at exact absolute paths, plus a one-line confirmation and a
  short summary of what changed and why. When not writing, a concise patch plan or copy/paste-ready authoring instructions.

### Foundation design contract

When this skill authors or revises a resource for this foundation, preserve these cross-cutting rules unless the user explicitly requests a different architecture:

- **Tool dependency semantics:** for agents, frontmatter `allowed-tools` means hard runtime dependencies. Use `optional-tools` for host/config-dependent capabilities that improve the worker but whose absence must not block the sub-agent. Never make `python_execute` a standard dependency; it exists only when the external Python helper is installed/configured.
- **Interaction ownership:** `ask_user` belongs to an interactive parent skill/orchestrator, never to a sub-agent. Author parent skills to call it only when advertised; AutoPilot and other non-interactive runs must surface the minimum clarification requirement without calling it.
- **Sub-agent call contract:** every authored parent workflow that invokes an `agent_<name>` must supply stable `subagent_task_id` and `expected_artifacts` on every invocation. Use `expected_artifacts: []` for non-file-producing workers.
- **Context safety:** keep the parent context compact. Large-document reduction, bounded research, comparison, requirement checking, row extraction, and other source-heavy work should be delegated to the appropriate isolated agent when available rather than dumping raw source material into the parent.
- **Bounded research:** research one concrete unresolved question at a time, prefer authoritative/primary sources, reassess after each round, and stop when additional retrieval is unlikely to change the answer materially.
- **Bounded mutation recovery:** exact-anchor document edits get one initial attempt plus at most two recovery attempts per logical operation. Once unresolved, do not reopen the same logical edit in the parent. Host-side circuit breakers may enforce stricter limits.
- **Real file finalization:** when the requested outcome requires a file, the workflow must invoke an actual create/save/export/finalizing tool and use the path/reference returned by that successful tool. A path mentioned in prose, planned JSON, a read/extract result, or an in-place mutation is not proof of a final deliverable.
- **Returned Word files:** perform all required mutations/comments first, then use `word_save_as` as the normal canonical finalizer when that tool is available for the target host. If any mutation occurs after the save, finalize again.
- **Host state boundary:** authored skills/agents produce the real final artifact; they do not invent or set `RegisteredDeliverableArtifacts`, `IsFinalDeliverable`, Outlook forced-delivery state, attachment state, or any delivery-confirmation/scheduling field. Those are host responsibilities.
- **Completion semantics:** if a required final file cannot actually be produced, return `blocked`; never use `complete` to mean "work was attempted".
- **Review authorship:** for newly created comments, annotations, tracked changes, markup, or comparable review metadata, use `Inky` as the default author/reviewer when the user has not specified another identity. Pass that value explicitly when the selected tool supports an author/reviewer/display-name parameter. Never rewrite existing authorship unless requested, and never claim the author was set when the tool cannot control it.
- **AutoPilot conversation-file retention is out of Inky scope:** when authoring or revising `Inky.md` / AutoPilot guidance, never add instructions that tell the model to ignore, discard, suppress, forget, or avoid reusing files from earlier messages merely because they are earlier. Never override or de-authorize files tagged `[RETAINED FROM EARLIER MESSAGE IN THIS CONVERSATION]`. Those files are surfaced intentionally by the host/system prompt. Retention itself is controlled in code/configuration (for example the host retention gate), not by Inky prompt text. If a user asks to disable that behavior through Inky, explain that the control belongs in the host configuration instead.

## 1. The three tooling surfaces (and how they really differ)

There are exactly three surfaces. **Excel is NOT a host surface** — never target an "Excel host" —
but Excel *workbooks* are handled by Excel-specific tools where those tools are available.

| Surface | Generic `skill_use` directly exposed? | Can run selected skills? | Live open-document tools | Attachments / drag & drop | Persistent workspace | Desktop delivery |
|---|---|---|---|---|---|---|
| **Word** (desktop add-in) | Yes | Yes | `worddoc_*`, `word_doc_*` (active document) | No | Optional connected workspace; session staging/temp remains valid without one | Yes — cite/identify intended session outputs so they are collected and delivered; staging is then cleaned |
| **Outlook Local Chat / Agent** | Yes | Yes | No | Yes (`list_attachments`, `read_attachment`, `search_in_attachments`, drag-&-drop files can land in session/staging) | Optional connected workspace; session/staging remains valid without one | Cite/identify intended session outputs for delivery; persist to a connected workspace only when persistence beyond the session is required |
| **Outlook AutoPilot** (unattended) | No | Yes — through dynamic `skill_<name>` tools | No | Yes (session files/attachments) | `agent_workspace_*`, locked to the workspace root | Produced deliverables only; runs unattended with no interactive user |

**Critical distinction:** do not confuse the generic `skill_use` tool with the ability of a host to
run a skill.

- `skill_use` is the generic skill-loader tool.
- `skill_<name>` is a dynamic tool representing one selected skill.
- `agent_<name>` is a dynamic tool representing one selected agent.

Word and Outlook Local Chat can expose the generic `skill_use` tool directly.

AutoPilot may still run skills and subagents even when the generic `skill_use` tool is not directly
advertised, because selected skills and agents can be exposed as dynamic `skill_<name>` and
`agent_<name>` tools that route internally to the same skill/agent runtime.

Therefore:

- A skill **can** run on **Word**, **Outlook Local Chat**, and **Outlook AutoPilot**.
- Do **not** infer from the absence of generic `skill_use` that AutoPilot cannot run skills.
- Host compatibility must be authored based on the actual runtime tool surface and capabilities.

### AutoPilot exclusions (from `Red_Ink_Tool_List.md`)
Not available under AutoPilot: generic `skill_use` as a directly advertised tool, all `memory_*`,
all `m365_*`, `web_content_retriever`, and the Word live-document tools. `tool_describe`,
`context_expand`, and `context_compact` are available on AutoPilot when advertised. AutoPilot writes
must stay inside its workspace root.

The authoritative tool list `Red_Ink_Tool_List.md` may live at the CENTRAL `.inky` resource root even
when you only have LOCAL authoring (write) rights. You can still READ it: the configured local and
central resource roots are both permitted read roots. Read it with `text_read` using the absolute path
from `resource_index` (central root) when you need to verify which tools a host actually exposes; never
copy it into the local tree just to read it.

## 2. Does the running skill know its host? Make it deterministic.

The tooling loop advertises only host-appropriate tools, but the model is not always told the host
by name. Do NOT let an authored skill guess. Establish the host at the top of every authored
workflow, in this priority order, and record it as an explicit fact for the rest of the run:

1. **Authoritative signal:** the loaded skill's `resource_index.host` field
   (e.g. `"Word"`, `"Outlook Local Chat"`, or `"Outlook AutoPilot"`). Use it verbatim when present.
2. **Capability probe (fallback):** if `host` is absent, derive it from the visible tool set —
   presence of `worddoc_*` / `word_doc_*` ⇒ **Word**; presence of attachment-oriented tools without
   Word live-document tools ⇒ **Outlook family**. Never assume; only conclude from tools actually offered.
3. When the probe yields only **Outlook family** and a more specific distinction matters, use the
   presence or absence of interactive-only tool families to refine the conclusion:
   - `m365_*` and `memory_*` available ⇒ likely **Outlook Local Chat**
   - `m365_*` and `memory_*` unavailable, workspace-root restrictions apply, and the run is unattended ⇒ likely **Outlook AutoPilot**

Every authored skill that behaves differently per host MUST begin by resolving the host this way
and must branch on the resolved value — not on assumptions.

## 3. Compatibility gating (block, don't fail messily)

For **agents**, distinguish required and optional capabilities explicitly:

- `allowed-tools`: every listed tool is a hard dependency; the host may block the isolated run if any required exact tool is absent.
- `optional-tools`: the host includes only names that exist in the authoritative registry snapshot; missing optional tools are ignored.
- Put host-specific source access (`m365_*`, attachment-only tools, `agent_workspace_*`) and configuration-dependent helpers such as `js_run` under `optional-tools` unless the agent genuinely cannot perform its defined job without them.
- Do not use `python_execute` as a generic fallback.


If a skill uses a tool that exists on only some hosts, it MUST check availability and block cleanly
when the host cannot support it, instead of calling a tool that isn't there.

Author every host-specific skill to:

1. Resolve the host (Section 2).
2. Confirm each required tool is actually offered this run (it appears in the advertised tool set /
   can be loaded via `tool_loader`). Do not rely on the name alone.
3. If a required tool is missing for the resolved host, STOP before doing partial work and end with:
   `<TASK_STATUS>{"status":"blocked","reason":"Required tool <name> is not available on <host>; this skill needs <host list>."}</TASK_STATUS>`

State the supported host(s) explicitly in the authored skill's body so the reason message is truthful.

## 4. Authoritative tool list & choosing among overlapping tools

`Red_Ink_Tool_List.md` in the `.inky` directory is the source of truth (columns: Word, Outlook,
AutoPilot). Before putting any tool in `allowed-tools`:

1. Read `Red_Ink_Tool_List.md` with `text_read` to verify registered **host tools** and host availability; use `tool_describe` separately when exact schemas or limits are needed. Dynamic selected-resource tools (`skill_<name>` / `agent_<name>`) are instead verified against `resource_index` and the advertised runtime tool surface.
2. Confirm each concrete host tool exists and is `Yes` for every host the skill targets; confirm each dynamic skill/agent resource is actually advertised for the current run before invoking it.
3. A narrow wildcard family (e.g. `file_*`) is allowed ONLY if the runtime expands wildcards for the
   target host AND every concrete tool it expands to is verified `Yes` for those hosts. Never use `*`.
4. When several tools overlap, do NOT pick by name. `tool_describe` is supported on all three target
   surfaces when advertised; inspect exact parameters, inputs/outputs, and limits, then choose by capability.
   In this skill's own Word/Outlook Local Chat execution, use it before the one-time `tool_loader` call
   when it is needed to decide which tool definitions to load. Key overlaps:
   - `worddoc_*` (active open doc, Word only) vs. `word_*` (`.docx` on disk, all hosts) vs.
     `word_doc_*` (Word host bridge, Word only).
   - `word_write` vs. `word_markup` (tracked-changes variant) vs. `word_format` vs. `word_comment_add`.
   - Excel: `excel_read_live_range` vs. `excel_complete_live_workbook` vs. `create_excel_spreadsheet`
     vs. `extract_excel_data`.
   - Generic `file_*` vs. document-specific create/convert tools (`create_word_document`,
     `word_to_pdf`, `pdf_to_word`, `complete_word_tables`, `create_powerpoint`, etc.).
5. In the **Word chatbot**, edits to the open document normally go through the inline
   `[#REPLACE …]` / `[#INSERTAFTER …]` command channel, which takes precedence over tool calls.
6. Online sources must be listed in `allowed-tools` to be usable (a wildcard such as
   `swiss-caselaw*` or the placeholder `selected_online_sources` is acceptable).

## 5. Files, attachments, workspaces, staging, and delivery

Author skills so file handling is deliberate and host-appropriate.

- **Word:** a connected persistent workspace is optional. If none is connected, session/staging/temp is a valid input/output area. For a returned file, create/finalize the real file in the host-provided staging/output area. Word later collects outputs through its host-side staging/output collection path (currently `WordCollectAndCopyOutputs`); the skill must not simulate that host state. If persistence beyond the session is required and a connected workspace exists, persist separately without treating persistence as proof of final delivery.
- **Outlook Local Chat:** a connected persistent workspace is optional. Inputs may arrive as mail attachments or drag-&-dropped session/staging files. Produce the real requested output in the host-provided session/staging area; the host owns any later attachment collection/delivery. A workspace copy is persistence, not a substitute for the final session artifact unless the runtime explicitly defines it that way. An attachment name is NOT a filesystem path.
- **AutoPilot:** unattended. Produce/finalize requested files only in the host-permitted session/working/staging/workspace locations actually exposed for that run. The host may validate/register the artifact and, after accepted completion, promote eligible Outlook artifacts into its forced-delivery mechanism before attachment collection. The skill must not invent or gate on those host-internal states.
- **Skill assets:** copy reference/template files from the skill's `references/` into the workspace
  or staging area BEFORE modifying/producing from them; never edit assets in place under the skill.

### Single-final-output discipline (avoid file sprawl)
When a workflow performs several operations on the same document, do NOT emit a fresh file on every
step. Author the skill to:

1. Prefer **in-place** editing tools when the tool supports it (e.g. `word_write` / `word_markup`
   operate on the same `.docx` by path across many `tasks`).
2. When a tool inherently produces a new file (e.g. `word_save_as`, `excel_complete_live_workbook`
   which writes `_completed.xlsx`, or conversions), use a **stable, predictable output name** and
   overwrite/replace the working copy rather than accumulating `_v1`, `_v2`, `_final_final` variants.
3. Treat intermediates as intermediate: if a chain has genuine intermediate artifacts, delete or
   move them (`file_delete` / `file_move`) once superseded, or clearly name the ONE final file.
4. Never claim completion pointing at an intermediate. State exactly which single file is the final
   deliverable, and make its name unambiguous.

### Tool-to-tool handoff
Before chaining tools, verify the artifact each step *produces* is the representation the next step
*accepts* (path vs. attachment vs. workspace item vs. open document vs. text vs. structured result),
and that it survives into the next turn. Add an explicit bridge (save/copy/register) when needed, or
do not author the chain. Confirm with `tool_describe` when uncertain.


## 5a. Resource-specific authoring recipes (generic extension mechanism)

The core skill-author must remain resource-agnostic. Resource-specific creation, compilation,
synchronization, migration, expert-review, sample-library, or reference-package conventions belong
in this skill's own `references/` directory, not as hard-coded branches in the generic workflow.

When creating, revising, synchronizing, reviewing, or explaining how to maintain a resource:

1. Inspect this skill's `resource_index` for its own reference files. If `authoring_recipes.json` is
   available, read it before drafting the change. Do not guess or invent a recipe path.
2. Match the target resource and requested operation against that registry. A recipe may match a
   particular resource name/type or another explicit registry condition.
3. If exactly one recipe matches, read its referenced instruction file and apply it as an extension
   of this SKILL.md. The recipe may define derived artifacts, synchronization rules, review mirrors,
   stable-id policies, validation, sample disclaimers, multilingual conventions, or package structure.
4. If no recipe matches, use the generic authoring workflow in this skill. If multiple recipes match
   materially and the registry does not define precedence, stop and resolve the ambiguity rather
   than combining incompatible recipes.
5. A recipe may **narrow or specialize** authoring behavior, but it may not override host/tool
   availability, author-mode permissions, safe-failure, real-file finalization, or other foundation
   safety contracts in this SKILL.md.
6. Keep resource-specific names, schemas and transformation logic in the recipe/reference files.
   Do not add a new hard-coded branch to this core skill merely because one resource family needs a
   special authoring process.
7. Treat maintainer-only HTML comments in sample/reference Markdown as non-runtime metadata. Preserve
   them when revising a package unless the user asks to remove them, but do not compile them into
   executable JSON, user-facing prompts, reports, actions, or runtime explanations.

This mechanism is also the preferred place for future resource-specific authoring adapters.

## 5b. How-to / command guidance mode

When the user asks **how to proceed**, **what to type**, **which command/prompt to enter**, **which
files to give Red Ink**, or **how an expert revision should be fed back**, answer that question
without modifying resources unless the user also explicitly asks for the change.

1. Resolve the applicable recipe under Section 5a and read the target resource's README/reference
   authoring guidance when available.
2. Explain the simplest supported Red Ink workflow in ordinary language. Prefer the human-readable
   source-first path over manual JSON editing whenever the recipe supports it.
3. Give one or more short, copy/paste-ready natural-language prompts the user can enter in Word or
   Outlook Local Chat. These are user prompts, not internal tool calls.
4. State which source file(s) should be attached or made available, what Red Ink will update, and
   which human-readable file the expert should review.
5. For synchronization, make clear whether stable ids are preserved, what is regenerated, and when
   Red Ink must ask a subject-matter question instead of guessing.
6. Do not require users to understand or hand-edit generated JSON merely because the runtime package
   uses JSON internally.

## 6. Converting Claude / foreign skills to this platform

When asked to convert an existing (e.g. Claude) SKILL.md, or to check whether a skill needs adapting:

1. Read the source with `text_read`.
2. **Map every tool** it references to a real Red Ink tool via `Red_Ink_Tool_List.md`. Foreign tool
   names (bash, filesystem, code execution, web fetch, etc.) rarely exist here — replace them with
   verified equivalents (`js_run` for deterministic computation when it is actually advertised; `file_*`/`text_*`/`workspace_*`
   for files; the web tools for retrieval) or remove the capability and note the loss. Do not assume
   `python_execute`: it is available only when the external Python helper is installed/configured.
3. **Rewrite the frontmatter** to this schema (Section 7); drop unsupported keys; set `allowed-tools`
   to verified names only.
4. **Add host resolution + compatibility gating** (Sections 2–3) if the converted workflow uses any
   host-specific tool.
5. **Add the task-status footer contract** and safe-failure behavior (Section 8) — foreign skills
   won't have these.
6. **Apply single-final-output discipline** (Section 5) if the workflow touches files repeatedly.
7. Report a compatibility verdict: *runs as-is*, *runs after the listed adaptations*, or
   *cannot run here* (with the blocking reason). When only reviewing, output the verdict + patch
   plan without writing.

## 7. Frontmatter schema

- `name`: unique resource name (kebab-case).
- `description`: one concise sentence used in the skill/agent listing.
- `allowed-tools`: list of registered tool names. For agents these are hard execution dependencies; for skills they define the permitted/declarable helper surface.
- `optional-tools` (agents only, optional): host/config-dependent registered tool names that may be used when present but whose absence must not block the isolated run.
- `model` (optional): a special-task-model key, e.g. `agentdefaultmodel` or `researchmodel`.
- `network` (optional, default false): opt-in for tools that touch the network (`js_run` with navigation, web tools).
- `timeout` (optional): seconds; 0 = default.
- `enabled` (optional, default true): set `false` ONLY when the user explicitly asks for the resource to be created or kept inactive. A disabled resource remains on disk and editable but is not offered to the model.

Never add `enabled: false` on your own initiative. A disabled resource stays on disk and editable in
"Manage Skills & Agents" but is not offered to the model until re-enabled.

### 7a. Dependency declaration contract — make every skill runnable on its own

A skill's `allowed-tools` is an **execution dependency contract**, not documentation decoration. When
authoring or revising a skill, derive this list from the workflow and declare every helper the skill may
need on any supported host. This is especially important for Outlook AutoPilot sender policies using
`ONLY skill_<name>`: the host may retain only the named skill plus helpers declared here.

Binding rules:

- If the skill reads any text/JSON/Markdown file under `references/` or `scripts/`, include `text_read`.
- If the skill may ask a live user for outcome-determinative information in Word or Outlook Local Chat,
  include `ask_user` and author an explicit unattended AutoPilot branch that does **not** call it.
- If the skill creates or finalizes a user-facing file, include the actual create/save/export/finalizer tool
  that proves the file exists (for example `word_apply_template`, `word_save_as`, or the relevant
  create/export tool). Do not rely on prose or a helper agent to create the final file.
- If the skill reads attachments, active Word content, workspace files, or performs deterministic
  computation, include the exact corresponding tools it actually uses.
- Dynamic `agent_*` helpers may be declared when useful, but a user-facing skill must remain capable of
  completing its **core workflow without an optional agent** unless the user explicitly requested an
  agent-dependent architecture. Put the fallback behavior in the skill body.
- Never add broad unused helper families merely 'just in case'. Verify every declared static tool for each
  target host. Host-unavailable optional tools may remain in a multi-host skill only when the skill gates
  their use by resolved host/tool availability.
- For a skill intended for `ONLY skill_<name>` AutoPilot use, test that the skill itself is selected for the
  AutoPilot session and that every **required AutoPilot helper** is both declared in `allowed-tools` and
  available on AutoPilot. The sender policy narrows an existing authorized session; it must not be treated
  as a way to enable an otherwise unselected skill or external service.

Before writing a skill, make a compact dependency table internally: `workflow step -> tool -> host ->
required/optional`. Use it to build the smallest complete `allowed-tools` list. During review, a missing
required helper is a blocking defect because the skill may load successfully yet be unable to execute.

## 8. Runtime contract & safe failure

- Each turn is either tool calls OR final prose. During active tooling, final prose ends with exactly
  one `<TASK_STATUS>{"status":"complete"|"blocked","reason":"..."}</TASK_STATUS>` footer.
  Use `complete` only when the user-facing task is truly done; a finished tooling session is not the
  same as a finished task.
- `text_write` writes UTF-8 text only (`SKILL.md`, `AGENT.md`, text assets). Use `file_*` for binary
  assets (`.docx`, `.dotx`, `.xlsx`, `.pptx`, `.pdf`, images, archives). Default per-file limit 2 MiB.
- `js_run`: the `code` param is the BODY of an async function; return a top-level value. Network is
  off unless `network: true`. Use it for deterministic validation (JSON, tables, dates, dedup).
- Every authored resource must describe safe-failure behavior: required tool/host unavailable, author
  mode off, missing asset/source, incompatible handoff, structure mismatch, permission denied, and
  partial-vs-final output — and must not report completion when only an intermediate step succeeded.

## 9. Author mode & where files go

Writing into the resource tree requires "Skill-author mode" (Manage Skills & Agents). Read these
flags from the `skill_use` `resource_index` BEFORE any write and act deterministically:

- `author_mode_active` / `local_writes_allowed` — if either is `false`, do NOT create or edit
  anything; tell the user to enable Skill-author mode and end with
  `<TASK_STATUS>{"status":"blocked","reason":"Skill-author mode is disabled; resources are read-only."}</TASK_STATUS>`.
- `central_writes_allowed` — write to the central root ONLY when this is `true`; otherwise write
  everything under the local root. Prefer local unless the user explicitly asks to change the shared set.
- `new_resource_root` — authoritative target for NEW resources; already accounts for central
  permission. Never override it toward `central_root` when central writes are disallowed.

### Resource layout

    <root>/
      Inky.md                        # optional project-wide guidance
      Red_Ink_Tool_List.md           # authoritative tool/availability reference
      skills/
        <skill-name>/
          SKILL.md                   # required; YAML frontmatter + Markdown body
          scripts/                   # optional helper scripts
          references/                # optional reference files/templates
      agents/
        <agent-name>.md              # single-file agent, OR
        <agent-name>/AGENT.md        # folder-based agent

- NEW skill: `new_resource_root + "\skills\<name>\SKILL.md"`.
- NEW agent: `new_resource_root + "\agents\<name>\AGENT.md"` (or `...\agents\<name>.md`).
- Always use ABSOLUTE paths from `resource_index` for EVERY resource write. NEVER omit the `path`
  argument and NEVER pass a relative path when creating or editing a skill/agent. Author mode only
  *permits* writing into the resource tree; it does NOT redirect a default write there. An omitted or
  relative path is resolved against the default writable root (connected workspace, else session
  staging, else the user's DESKTOP) — so without a connected workspace the resource is written to the
  Desktop and is NOT installed. Construct the target explicitly, e.g.
  `new_resource_root + "\skills\<name>\SKILL.md"`, and confirm it is under `local_root` (or
  `central_root` when `central_writes_allowed`) before writing.
- Do not use `agent_workspace_*` on the `.inky` tree — it is not a workspace and those calls fail
  with "No active workspace".
- Create `references/` / `scripts/` with `file_make_dir`; place text with `text_write`, binaries with
  `file_copy`/`file_move`/`file_rename`. Ensure any referenced template exists before finishing.

## 9a. Diagnosing previous tooling runs (log analysis)

This skill is ALSO the authority for answering "what happened / what went wrong in the last tool run(s)".
When the user asks to diagnose, debug, or understand a previous run, DO NOT rely on chat history or
memory alone — read the diagnostics logs, which are the ground truth of the tooling loop.

1. **Diagnostics access requires Skill-author mode.** Read `author_mode_active` from `resource_index`.
   If it is `false`, do not attempt to read diagnostics; tell the user to enable Skill-author mode
   (Manage Skills & Agents), then end with
   `<TASK_STATUS>{"status":"blocked","reason":"Skill-author mode is off; enable it to permit diagnostics access. With Skill-author mode on, subsequent tooling runs are also guaranteed to be logged."}</TASK_STATUS>`.
   Diagnostics logging itself is a separate configuration: it may be enabled independently by the user,
   and it is always enabled while Skill-author mode is on. Therefore prior logs may exist from independent
   logging even if Skill-author mode was off at the time, but this skill may access them only while
   Skill-author mode is currently on.
2. Logs live under `local_root + "\diagnostics\"`. There are two kinds per run:
   - `RI_Tooling_Log__<timestamp>__<skill>.txt` — the full tooling-loop trace (tools loaded, tool
     calls, iterations, final response, session summary with `Success:`/`Failed:`).
   - `RI_SubAgent_Returns__<timestamp>__<skill>.txt` — sub-agent / skill return payloads.
   The newest five of each kind are kept; the most recent timestamp is the latest run. For every run in
   which diagnostics logging is active (either by independent user configuration or because Skill-author
   mode is on), a tooling log is written even when no skill name was captured.
3. Identify the RIGHT file deterministically. Never guess a fixed filename such as `run.log`, and never
   fabricate fallback paths. Prefer `resource_index.diagnostics_files`: it is the authoritative
   diagnostics inventory for this run and already contains the exact available filenames plus metadata.
   When `resource_index.diagnostics_files` is non-empty, use it DIRECTLY as the inventory and do NOT
   also call `file_list` — the inventory already lists the exact files under the permitted read root.
   When `resource_index.diagnostics_files` is empty or absent, and `file_list` is allowed and available,
   call `file_list` on `local_root + "\diagnostics\"` to enumerate the exact files under that permitted
   read root. Do not use `js_run`, `python_execute`, or other sandbox workarounds to probe the host
   filesystem. If neither `resource_index.diagnostics_files` nor `file_list` yields a deterministic
   diagnostics inventory, STOP immediately and end with
   `<TASK_STATUS>{"status":"blocked","reason":"No deterministic diagnostics file inventory is available in this run, so the previous tooling run cannot be diagnosed safely from logs."}</TASK_STATUS>`.
   From the returned filenames, keep only `RI_Tooling_Log__*.txt`, sort by the timestamp embedded in
   each filename (or by returned write-time metadata when present), and take the newest 4–5 runs. Read
   each briefly (header + final `Success:`/`Failed:` summary) and build a short inventory: for every
   run note its timestamp, the skill name if present in the filename, and whether it ended in Success
   or Failure. EXCLUDE the current in-progress diagnosing run from selection: the run that is executing
   this diagnosis writes its own `RI_Tooling_Log__<timestamp>__skill_author.txt` first, so the newest
   entry is normally THIS run, not the run the user wants diagnosed. Drop that self-referential newest
   entry (the largest timestamp whose skill slug is this authoring run) and diagnose the newest run that
   remains. If, after excluding the current run, exactly one prior run remains, diagnose it WITHOUT
   asking. If MORE THAN ONE prior run remains, you MUST call `ask_user` before reading any log unless the
   user's request already named a specific run (by timestamp, skill name, or an unambiguous phrase like
   "the last failed run"). This is a mandatory disambiguation step, not a discretionary one: with several
   candidate logs the most recent run is NOT a "harmless obvious default", because diagnosing the wrong
   run wastes the whole turn and misleads the user. Do not skip `ask_user` on the grounds that a default
   exists, that the newest run is probably meant, or that a chat sentence would be simpler — for this
   multi-run case the `ask_user` tool is the required channel. Present one concise question with the 4–5
   most recent PRIOR runs as concrete `options` (each label = timestamp + skill name + Success/Failed).
   Because `skill-author` itself executes only in interactive Word or Outlook Local Chat, do not add an
   unattended fallback here. Include the 4–5-line prior-run inventory in the final response when useful
   so the user can redirect you to a different run on the next turn. Do not use `js_run` to access the
   filesystem, and do not use `require(...)`, `fs`, `process`,
   `__dirname`, or other Node APIs there; `js_run` is only for in-memory computation on data already
   read by file/text tools.
4. Read the chosen log with `text_read` (the `diagnostics/` folder is a permitted read root). For large
   logs use `text_search` for the tools-loaded list, `text_read: not_found` or other path errors,
   `js_run` misuse, `workspace_write` vs. skill-root writes, and the final `Success:`/`Failed:` summary.
5. Diagnose against this skill's authoring rules: was the authoring skill loaded and used; did writes
   land under an authorized resource root rather than the temporary workspace; does the claimed
   completion match the real destination? Report the root cause and the exact fix.

## 9b. Independent advisor for consequential authoring decisions

Use `agent_advisor` as an **optional isolated second pass**, not as a routine step for every edit. Invoke it when it is actually advertised/available and the authoring decision is materially consequential, for example:

- a host-crossing architecture or delivery/finality contract is being changed;
- two plausible tool/workflow designs have meaningful reliability, security, or compatibility trade-offs;
- a change affects several skills/agents or the project-wide `Inky*.md` contract;
- the requested behavior is ambiguous enough that a second-pass challenge could prevent a systemic mistake.

Do not invoke the advisor for mechanical wording changes, obvious metadata edits, or routine one-resource maintenance. Give it only the compact facts, constraints, candidate design, and unresolved decision; do not send raw logs or large documents. Treat its result as advice, not authority: verify tool/host facts against `resource_index`, the advertised tool surface, `Red_Ink_Tool_List.md`, and `tool_describe` before writing.

Dynamic resource tools such as `agent_advisor` are validated against the **advertised runtime/resource index**, not by assuming that every dynamic `agent_*` or `skill_*` name must appear as a static row in `Red_Ink_Tool_List.md`. If `agent_advisor` is not exposed for the current interactive run, continue without it.

## 10. Skills vs. agents

- A **skill** is loaded into the SHARED conversation and guides later turns. This may happen through
  generic `skill_use(name, input?)` where that tool is directly exposed, or through a dynamic
  `skill_<name>` tool where the host exposes selected skills that way. Best for user-facing,
  multi-step, context-dependent workflows.
- An **agent** is delegated via `agent_<name>(task, context?)`, runs ISOLATED, and returns
  `{summary, result, memory_key, stub}`. Best for a bounded sub-task that would otherwise burn context;
  `task`/`context` must be fully self-contained.

When documenting compatibility, distinguish:
- **generic loader availability** (`skill_use` directly exposed or not),
- **actual skill execution support** (selected `skill_<name>` tools may still work),
- **actual agent execution support** (selected `agent_<name>` tools may still work).

## 11. Load your tools first

For this skill's own Word/Outlook Local Chat execution, plan the tool needs for the whole run and call
`tool_loader` ONCE before the first substantive read/write/copy/move/rename/create action. If overlapping
candidates must be compared, use already-advertised `tool_describe` first, then load the complete chosen
set (typically `text_read`, `text_write`, `text_search`, and needed `file_*`). A freshly loaded tool is
callable only from the NEXT turn, so load the full useful set up front rather than one at a time and do not
call `tool_loader` again later in the same run.

## 12. Finding an existing resource to edit (do this FIRST for edits/conversions)

1. Look up the resource by name in `resource_index`.
2. Read its exact `file` path with `text_read`. Never guess a path; base every edit on actual content.
3. Write changes back to the SAME `file` path with `text_write`. Do not create a new folder for an
   existing resource, and do not fork a central resource into local unless the user asks.

## 13. Authoring workflow

1. Decide skill vs. agent, and target host(s). Check `resource_index` access flags (Section 9); block
   immediately if author mode is off.
2. Read `Red_Ink_Tool_List.md`; verify every intended tool for every target host; compare overlaps
   with `tool_describe`.
3. For edits/conversions, read the exact existing file first (Section 12 / Section 6). For any create/revise/sync/review/how-to operation, also resolve and read an applicable resource-specific authoring recipe under Section 5a before drafting or advising.
4. Draft/revise the body with: purpose, inputs, target host(s), host resolution + compatibility
   gating, workflow, tool usage, file/output management (single final output), output format,
   limitations/safe-failure. Build the Section 7a dependency table and make `allowed-tools` the smallest
   complete set that lets the skill execute its own core workflow on every claimed host.
5. Validate deterministically with `js_run` where useful.
6. Write NEW resources to an absolute path under `new_resource_root`; edit EXISTING resources at their
   exact `file` path. Ensure required `references/`/`scripts/` assets exist (binaries via `file_*`).
   State every exact absolute path touched.
7. Do not accumulate accidental duplicates. Remove a stale/superseded resource only when that status is clear and deletion is authorized; otherwise report the overlap instead of deleting it.

## 14. Review checklist

1. Target host(s) identified; Excel not treated as a host.
2. Every static host tool is verified in `Red_Ink_Tool_List.md` for each target host; every dynamic `skill_*` / `agent_*` dependency is verified against `resource_index` and the advertised runtime surface.
3. Host resolution present (uses `resource_index.host`, falls back to capability probe).
4. Host-specific tools are gated with a clean `blocked` path; supported hosts stated truthfully.
5. Generic `skill_use` versus dynamic `skill_<name>` usage distinguished correctly; `memory_*` and `m365_*` dependencies flagged; AutoPilot compatibility considered where relevant.
6. Overlapping tools chosen deliberately via `tool_describe`.
7. Attachment vs. path vs. workspace-item vs. open-document representations handled correctly;
   handoffs verified.
8. Single-final-output discipline applied; intermediates cleaned or the one final file named.
9. Author-mode/write-permission flags respected; NEW under `new_resource_root`, edits at exact path;
   no relative `.inky` paths.
10. Required `references/`/`scripts/` assets exist; binaries via `file_*`, not `text_write`.
11. Frontmatter valid; `name` unique/kebab-case; `description` one sentence; `enabled:false` only on
    explicit request. `allowed-tools` satisfies the Section 7a dependency contract: references have their
    reader, interactive clarification has `ask_user`, real outputs have a finalizer, and optional agents are
    not the sole implementation of the core workflow.
12. Task-status footer contract and safe-failure behavior included; completion reflects the user task.
13. File-producing workflows create/finalize a real output; Word-return flows finalize after the last mutation; no model-visible host registry/delivery state is invented.
14. Context-heavy work is delegated/compacted appropriately; research and edit retries are bounded.
15. New review metadata defaults to author/reviewer `Inky` unless the user overrides it.
16. Applicable resource-specific authoring recipe resolved and followed; resource-specific logic remains in references rather than being hard-coded into this generic skill.
17. For how-to questions, provide copy/paste-ready natural-language prompts and required inputs without changing files unless requested.
18. For consequential architecture changes, consider an isolated `agent_advisor` second pass when available; never use it as a substitute for host/tool verification.

## 15. Output format

Return the revised Markdown resource(s) or a concise patch plan. When writing, state each exact
absolute path created/changed and whether it went to the local or central root. Always begin the
final response with a one-line confirmation, e.g.
"Applied skill-author: converted Claude skill 'deadline-calc' to run on Word + Outlook Local Chat."



## Orchestrator compatibility

- Use `ask_user` only when it is actually advertised and the current run is interactive. Ask one material question per call where the workflow requires incremental clarification. In AutoPilot or any non-interactive run, do not call `ask_user`; return/send the minimum concrete clarification needed or mark the affected branch blocked instead of guessing.
- Do not assume `python_execute` exists. The Python helper is installation-dependent and is not a standard foundation capability. Use deterministic host tools or configured scripts that are actually advertised; never substitute model-estimated computation for a missing deterministic tool.
- The parent skill owns end-user interaction. Sub-agents must receive a bounded task and must not be expected to ask the user.
- Every `agent_<name>` call must include a stable opaque `subagent_task_id` for that logical delegated task and `expected_artifacts`. Use `expected_artifacts: []` for analysis-only workers. If a delegated file-producing task is expected, pass the complete opaque artifact contract required by the host before the run; do not invent or broaden artifact identities later.
- Treat an agent's missing optional host capability as a reason to adapt the bounded task or return a limitation, not as permission to bypass the selected workflow.
