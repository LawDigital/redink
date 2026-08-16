# Red Ink Tool List

This file lists the built-in internal tools that can be advertised by Word, Outlook, and Outlook AutoPilot.

Notes:

- Availability can still depend on configuration, feature flags, model support, and current execution mode.
- `js_run` may be disabled by the user by setting the `JsRunDisable` parameter in the configuration file.
- `browser_open`, `browser_snapshot`, and `browser_interact` may be disabled by the user by setting the `BrowserToolsDisable` parameter in the configuration file.
- `python_execute` requires the `redink-pythonagent` helper to be installed and configured; Python execution is unavailable without it.
- **A path is not a deliverable.** For a file-required task, actually create/save/export the real final file in the host-provided session/working/staging location. For Word-return workflows, finish edits/comments first and use `word_save_as` as the normal finalizing step. Do not claim attachment or delivery from a path alone; the host/orchestrator validates eligible artifacts and performs Outlook attachment delivery.
- **Host registry terminology:** the structured host registry uses `RegisteredDeliverableArtifacts` and `IsFinalDeliverable`. Delivery scheduling is not a model-visible registry state. Outlook currently bridges eligible final artifacts through its forced-delivery mechanism before attachment collection; Word uses `WordCollectAndCopyOutputs` from staging. These are host responsibilities, not model-settable states.
- The `Outlook` column refers to Outlook tooling outside AutoPilot, primarily Local Chat / Agent mode.
- `AutoPilot` is listed separately because it does not expose the full Outlook tool surface.

## Shared tools

| Tool | Description | Word | Outlook | AutoPilot |
|---|---|---:|---:|---:|
| `retrieve_web_content` | Retrieves readable text, and optionally links, from one or more public web URLs. | Yes | Yes | Yes |
| `web_content_retriever` | Alias for the public web content retrieval tool. | Yes | Yes | No |
| `download_web_files` | Downloads remote files and saves the original binary files locally. | Yes | Yes | Yes |
| `internet_search` | Searches the public internet and returns readable content from top results. | Yes | Yes | Yes |
| `web_grounding` | Uses a web-enabled model to perform cited live-web research. | Yes | Yes | Yes |
| `knowledge_search` | Searches the user's local knowledge store for relevant internal content. | Yes | Yes | Yes |
| `tool_loader` | Lazily loads full tool definitions only when a specific tool is needed. | Yes | Yes | Yes |
| `report_progress` | Announces a short user-facing major-step progress update. Use it before the first substantive tool action and again whenever a new major phase begins. | Yes | Yes | No |
| `tool_describe` | Returns the full parameter schema and usage instructions for one or more tools (by name or name prefix) without making them callable, so overlapping tools can be compared before loading. | Yes | Yes | Yes |
| `context_expand` | Retrieves a character window from a large tool result that was stored by reference. Large results are replaced in context by a short `result_ref` plus a preview; this tool reads more of that stored content on demand. | Yes | Yes | Yes |
| `ask_user` | Asks the user one concise clarifying question (with optional concrete choices, single- or multi-select, or a requested text/number value) when required information is missing, several materially different interpretations exist, or a skill/workflow needs an explicit decision. The user may always answer freely. Interactive contexts only (Word chat and Outlook Local Chat); non-blocking no-op in unattended e-mail Scheduler / AutoPilot runs. | Yes | Yes | No |
| `context_compact` | Voluntarily compacts older tool results out of the active context to free space when they are no longer needed in full. The full text stays retrievable via `context_expand`; an optional `keep_recent` controls how many recent results stay fully visible. | Yes | Yes | Yes |
| `memory_put` | Stores a key/value memory entry with summary, tags, and metadata. | Yes | Yes | No |
| `memory_get` | Retrieves a stored memory entry by key. | Yes | Yes | No |
| `memory_list` | Lists stored memory entries and their summaries. | Yes | Yes | No |
| `memory_delete` | Deletes a stored memory entry by key. | Yes | Yes | No |
| `text_read` | Reads a UTF-8 text file within the allowed workspace boundary. | Yes | Yes | Yes |
| `text_write` | Writes, replaces, or appends a UTF-8 text file. | Yes | Yes | Yes |
| `text_search` | Searches text files for substring or regex matches. | Yes | Yes | Yes |
| `text_export_to_text` | Silently extracts readable text from a supported file or from supported files in a directory and writes UTF-8 `.txt` outputs without host UI prompts. | Yes | Yes | Yes |
| `semantic_index_create_from_file` | Creates a self-indexed semantic-search text file from an existing text file. | Yes | Yes | Yes |
| `semantic_index_create_from_text` | Creates a self-indexed semantic-search text file from supplied in-memory text. | Yes | Yes | Yes |
| `semantic_index_validate` | Validates whether a file is a readable semantic-search index and returns basic counts. | Yes | Yes | Yes |
| `semantic_index_search` | Runs an initial semantic retrieval against an indexed text file and returns grounded source excerpts plus internal handles for continuation and verification. | Yes | Yes | Yes |
| `semantic_index_search_continuation` | Continues a semantic retrieval conversation using a prior `conversation_handle`. | Yes | Yes | Yes |
| `semantic_index_load_entries` | Loads exact source ranges for trusted semantic-index entry IDs without re-running semantic selection. | Yes | Yes | Yes |
| `semantic_index_verify_answer` | Verifies whether a drafted answer is supported by the exact excerpts returned by a prior retrieval handle. | Yes | Yes | Yes |
| `semantic_index_retrieve_after_verification` | Retrieves additional semantic-index evidence after verification reports that more sources are needed. | Yes | Yes | Yes |
| `semantic_index_reset_conversation` | Resets and removes a stored semantic-search conversation handle. | Yes | Yes | Yes |
| `semantic_index_invalidate_cache` | Invalidates one indexed-file cache entry or the full semantic-search cache. | Yes | Yes | Yes |
| `js_run` | Executes sandboxed JavaScript in a hidden WebView2 environment. Availability is subject to the user's `JsRunDisable` configuration setting. | Yes | Yes | Yes |
| `browser_open` | Opens or navigates the shared Playwright browser session to an absolute HTTP/HTTPS URL. After success, call `browser_snapshot` before attempting interaction. | Yes | Yes | No |
| `browser_snapshot` | Captures the current page as an AI-optimized Playwright ARIA snapshot and returns short-lived refs such as `[ref=e7]` for later interaction. | Yes | Yes | No |
| `browser_interact` | Performs exactly one Playwright action against one ref from the most recent `browser_snapshot`. After every attempted interaction, take a new snapshot before another interaction. | Yes | Yes | No |
| `python_execute` | Executes sandboxed Python code through the configured secure Python agent and may return structured results or published output files.1) | Yes | Yes | Yes |
| `skill_use` | Loads a skill's instructions and file inventory for guided execution. Word and Outlook Local Chat expose this generic loader directly. AutoPilot does not advertise the generic `skill_use` tool, but it can still run selected skills through dynamic `skill_<name>` tools that route internally to the same skill loader. | Yes | Yes | No |
| `m365_search` | Searches Microsoft 365 content such as mail, files, chats, events, and notes. | Yes | Yes | No |
| `m365_get_mail` | Retrieves a mail message and its attachment text. | Yes | Yes | No |
| `m365_get_mail_thread` | Retrieves an entire mail conversation as one transcript. | Yes | Yes | No |
| `m365_get_file` | Retrieves a Microsoft 365 file and extracts its readable text. | Yes | Yes | No |
| `m365_get_event` | Retrieves calendar event details. | Yes | Yes | No |
| `m365_get_chat_thread` | Retrieves a Teams chat or channel thread. | Yes | Yes | No |
| `m365_get_onenote_page` | Retrieves a OneNote page and returns readable content. | Yes | Yes | No |
| `word_extract_text` | Extracts plain text from a `.docx` file on disk. | Yes | Yes | Yes |
| `word_search` | Searches a `.docx` file on disk for text or regex matches. | Yes | Yes | Yes |
| `word_write` | Inserts, replaces, appends, or deletes a paragraph in a `.docx` file on disk while preserving footnotes, endnotes, fields, images, comments, existing revision content, and run formatting. Supports batch `tasks`; matching is resilient within a single paragraph and results report `complete` / `partial` / `none` plus unmatched-anchor suggestions. | Yes | Yes | Yes |
| `word_markup` | Edits a `.docx` file on disk with tracked changes while preserving footnotes, endnotes, fields, images, comments, existing revision content, and run formatting. Uses word-level markup by default, collapses large rewrites, supports batch `tasks`, supports `delete_paragraph`, and reports `complete` / `partial` / `none` plus unmatched-anchor suggestions. | Yes | Yes | Yes |
| `word_comment_add` | Adds one or more Word comments to matched spans in a `.docx` file on disk while preserving footnotes, endnotes, fields, images, existing revision content, and formatting. Supports batch `tasks`, resilient single-paragraph matching, and reports `complete` / `partial` / `none` plus unmatched-anchor suggestions. | Yes | Yes | Yes |
| `word_comment_list` | Lists comments in a `.docx` file on disk. | Yes | Yes | Yes |
| `word_comment_remove` | Removes comments from a `.docx` file on disk. | Yes | Yes | Yes |
| `word_format` | Applies paragraph or run formatting to matched text in a `.docx` file on disk. | Yes | Yes | Yes |
| `word_apply_template` | Creates a document from a template with substitutions. | Yes | Yes | Yes |
| `word_save_as` | Saves a `.docx` file to a new path; use as the canonical finalizing step when a revised Word file must be returned. | Yes | Yes | Yes |

1) Only available when the `redink-pythonagent` helper is installed, configured, and available.

## Shared file tools

These binary-safe tools operate across the PathPolicy-governed roots (the agent workspace, the session staging/temp area, and skill `scripts`/`references`). Reading skill files is always allowed; writing into a skill folder requires skill-author mode; workspace operations honor the user's configured read/write/move/delete permissions.

| Tool | Description | Word | Outlook | AutoPilot |
|---|---|---:|---:|---:|
| `file_copy` | Copies a file or directory (binary-safe) between the workspace, staging area, and skill references/scripts. Use for representation/persistence handoffs, not as a substitute for host-confirmed delivery. | Yes | Yes | Yes |
| `file_list` | Lists files and/or directories under an allowed root, with optional glob filtering and recursion. | Yes | Yes | Yes |
| `file_move` | Moves a file or directory (binary-safe) between the allowed roots. | Yes | Yes | Yes |
| `file_rename` | Renames a file or directory in place. | Yes | Yes | Yes |
| `file_delete` | Deletes a single file (Recycle Bin by default). | Yes | Yes | Yes |
| `file_make_dir` | Creates a directory, including intermediate directories. | Yes | Yes | Yes |
| `file_remove_dir` | Removes a directory and its contents (Recycle Bin by default). | Yes | Yes | Yes |

## Workspace tools

| Tool | Description | Word | Outlook | AutoPilot |
|---|---|---:|---:|---:|
| `workspace_get` | Returns the current workspace path and permissions. | Yes | Yes | Yes |
| `workspace_inventory` | Lists files in the workspace with optional recursion and filtering. | Yes | Yes | Yes |
| `workspace_read` | Reads a UTF-8 text file from the workspace. | Yes | Yes | Yes |
| `workspace_read_many` | Reads multiple text files from the workspace in one call. | Yes | Yes | Yes |
| `workspace_write` | Writes, appends, or creates a text file in the workspace. | Yes | Yes | Yes |
| `workspace_search` | Searches across workspace file contents. | Yes | Yes | Yes |
| `workspace_copy` | Copies a file or folder within the workspace. | Yes | Yes | Yes |
| `workspace_move` | Moves a file or folder within the workspace. | Yes | Yes | Yes |
| `workspace_rename` | Renames a file or folder within the workspace. | Yes | Yes | Yes |
| `workspace_delete` | Deletes a file or folder from the workspace. | Yes | Yes | Yes |
| `workspace_make_dir` | Creates a folder in the workspace. | Yes | Yes | Yes |
| `workspace_extract_text` | Extracts readable text from a supported workspace file such as PDF, Word, or Excel. | Yes | Yes | Yes |
| `workspace_extract_text_many` | Extracts readable text from multiple supported workspace files. | Yes | Yes | Yes |

## Outlook and AutoPilot tools

| Tool | Description | Word | Outlook | AutoPilot |
|---|---|---:|---:|---:|
| `process_word_document` | Processes a Word document or attachment through the document-processing pipeline. | No | Yes | Yes |
| `comment_word_document` | Adds comment bubbles to a Word document. | No | Yes | Yes |
| `extract_pdf_text` | Extracts readable text from a PDF file. | No | Yes | Yes |
| `merge_pdfs` | Merges multiple PDFs into one output PDF. | No | Yes | Yes |
| `read_attachment` | Reads or extracts text from an email attachment. | No | Yes | Yes |
| `list_attachments` | Lists the attachments available in the current session context. | No | Yes | Yes |
| `describe_binary_attachment` | Produces a description of a non-text attachment. | No | Yes | Yes |
| `compare_word_documents` | Compares two Word documents and reports differences. | No | Yes | Yes |
| `read_word_document_details` | Returns metadata or structural details about a Word document. | No | Yes | Yes |
| `create_pdf_from_text` | Generates a PDF from supplied text content. | No | Yes | Yes |
| `extract_excel_data` | Extracts readable or structured data from an Excel file using the normal reader path. | No | Yes | Yes |
| `excel_list_live_worksheets` | Lists worksheet names and basic live worksheet metadata from an existing Excel file through Excel Interop. | No | Yes | Yes |
| `excel_read_live_range` | Reads a worksheet or range from an existing Excel file through live Excel Interop, including current values and workbook-state-dependent details. | No | Yes | Yes |
| `excel_complete_live_workbook` | Updates an existing Excel workbook through live Excel Interop and saves a new `_completed.xlsx` copy, including LiftLock handling where present. | No | Yes | Yes |
| `split_pdf` | Splits a PDF into multiple output files. | No | Yes | Yes |
| `add_pdf_watermark` | Applies a watermark to a PDF. | No | Yes | Yes |
| `word_to_pdf` | Converts a Word document to PDF. | No | Yes | Yes |
| `search_in_attachments` | Searches across attachment content for relevant matches. | No | Yes | Yes |
| `summarize_thread` | Summarizes an email thread. | No | No | Yes |
| `pdf_to_word` | Converts a PDF into a Word document. | No | Yes | Yes |
| `create_word_document` | Creates a new Word document output. | No | Yes | Yes |
| `complete_word_tables` | Completes existing Word tables, placeholders, and form fields in place. | No | Yes | Yes |
| `create_excel_spreadsheet` | Creates a new Excel workbook output. | No | Yes | Yes |
| `create_powerpoint` | Creates a new PowerPoint presentation output. | No | Yes | Yes |
| `create_code_file` | Creates a source code or text-based file output. | No | Yes | Yes |
| `comment_pdf_document` | Adds annotation comments to a PDF. | No | Yes | Yes |
| `extract_data_from_attachments` | Pulls structured information from one or more attachments. | No | Yes | Yes |
| `redact_pdf` | Redacts content in a PDF. | No | Yes | Yes |
| `overlay_pdf` | Overlays one PDF onto another PDF. | No | Yes | Yes |
| `create_audio_file` | Generates an audio file output. | No | Yes | Yes |
| `generate_image` | Generates an image file output. | No | Yes | Yes |
| `manage_scheduled_tasks` | Creates, lists, updates, pauses, resumes, or deletes scheduled tasks. | No | Yes | Yes |
| `manage_user_memory` | Manages per-user persistent memory storage. | No | Yes | Yes |
| `manage_user_files` | Manages files in per-user storage. | No | Yes | Yes |
| `report_inability` | Returns a structured inability report when the requested action cannot be completed. | No | Yes | Yes |
| `agent_workspace_list` | Lists files and folders in the agent workspace. | No | Yes | Yes |
| `agent_workspace_read` | Reads or extracts text from a workspace file. | No | Yes | Yes |
| `agent_workspace_write` | Writes a text or code file into the workspace. | No | Yes | Yes |
| `agent_workspace_file_op` | Performs safe file operations such as copy, move, rename, create folder, or delete inside the workspace. | No | Yes | Yes |
| `agent_workspace_save_session_file` | Copies a session-produced file into the workspace for persistence. This does not by itself prove Outlook attachment/delivery. | No | Yes | Yes |
| `agent_workspace_search` | Searches workspace filenames and text-like content. | No | Yes | Yes |
| `agent_workspace_find_files` | Finds workspace files by name, extension, size, or modified date. | No | Yes | Yes |
| `agent_workspace_move_to` | Moves one or more workspace items into another folder. | No | Yes | Yes |
| `agent_workspace_copy_to` | Copies one or more workspace items into another folder. | No | Yes | Yes |
| `agent_workspace_rename` | Renames a workspace file or folder. | No | Yes | Yes |
| `agent_workspace_bulk_rename` | Renames many workspace files using batch rules. | No | Yes | Yes |
| `agent_workspace_file_details` | Returns detailed metadata for a workspace file or folder. | No | Yes | Yes |
| `agent_workspace_recent_files` | Lists recently changed workspace files. | No | Yes | Yes |
| `agent_workspace_create_folder_structure` | Creates multiple folders under a workspace path in one operation. | No | Yes | Yes |
| `agent_workspace_trash` | Moves workspace files or folders to the Recycle Bin. | No | Yes | Yes |
| `agent_workspace_inventory_report` | Creates a Word or Excel inventory report for workspace files. | No | Yes | Yes |

## Word live-document tools

| Tool | Description | Word | Outlook | AutoPilot |
|---|---|---:|---:|---:|
| `worddoc_list_open` | Lists the documents currently open in Word. | Yes | No | No |
| `worddoc_get_active` | Returns metadata for the active Word document. | Yes | No | No |
| `worddoc_extract_text` | Extracts plain text from the active or a named open Word document. | Yes | No | No |
| `worddoc_search` | Searches the active or a named open Word document. | Yes | No | No |
| `worddoc_list_comments` | Lists comments in the active or a named open Word document. | Yes | No | No |
| `worddoc_insert_text` | Inserts text into the active or a named open Word document. Supports `location` (`start`, `end`, `cursor`) and `track_changes` (defaults to `true`). | Yes | No | No |
| `worddoc_replace` | Replaces a resilient-located span in the active or a named open Word document. Supports `only_first`, `track_changes` (defaults to `true`), and `match_scope` (`sentence`, `paragraph`, or exact span). | Yes | No | No |
| `worddoc_delete` | Deletes a resilient-located span in the active or a named open Word document. Supports `track_changes` (defaults to `true`) and `match_scope` (`sentence`, `paragraph`, or exact span). | Yes | No | No |
| `worddoc_comment_add` | Adds a comment to resilient-located text in the active or a named open Word document. | Yes | No | No |
| `worddoc_format` | Applies formatting to resilient-located text in the active or a named open Word document. | Yes | No | No |
| `word_doc_read` | Reads content from the active Word document through the Word host bridge. | Yes | No | No |
| `word_doc_edit` | Edits the active Word document through the Word host bridge. | Yes | No | No |
| `word_doc_create` | Creates a new Word document through the Word host bridge. | Yes | No | No |
| `word_doc_export_pdf` | Exports a Word document to PDF through the Word host bridge. | Yes | No | No |

## Additional Outlook and AutoPilot data-collection tools

| Tool | Description | Word | Outlook | AutoPilot |
|---|---|---:|---:|---:|
| `list_collection_use_cases` | Lists configured collection use cases for structured extraction workflows. | No | Yes | Yes |
| `collect_data` | Runs a configured data-collection workflow against the current session files. | No | Yes | Yes |
| `preview_collection` | Previews how a configured data-collection workflow would interpret the current request and files. | No | Yes | Yes |

## Choosing among overlapping tools

Several tool families overlap. Use `tool_describe` to read each candidate's exact
parameters before selecting one. Key distinctions:

- **Open-document vs. file-on-disk vs. host-bridge Word tools.**
  - `worddoc_*` operate on a document already OPEN in Word (Word host only), addressing the active or a named open document. They support resilient live-document matching for text operations; `worddoc_insert_text`, `worddoc_replace`, and `worddoc_delete` support `track_changes` (default `true`), and `worddoc_replace`/`worddoc_delete` also support `match_scope` for exact-span, sentence-level, or paragraph-level edits.
  - `word_*` (for example `word_extract_text`, `word_write`, `word_markup`, `word_format`, `word_comment_add`) operate on a `.docx` FILE by path and are available in Word, Outlook, and AutoPilot.
  - `word_doc_*` (`word_doc_read`/`word_doc_edit`/`word_doc_create`/`word_doc_export_pdf`) go through the Word host bridge for the active document (Word host only).
- **Editing the open Word document in the Word chatbot** is normally done through the inline `[#REPLACE …]`/`[#INSERTAFTER …]` command channel, which takes precedence over tool calls.
- **Word content edits on a file.** `word_write` and `word_markup` now preserve inline Word content such as footnotes, endnotes, fields, images, comments, existing revision content, and run formatting while editing only the matched text. Both support batch `tasks`; `word_markup` is the tracked-changes variant, while `word_write` performs the same kind of edit without revision marks. Use `delete_paragraph` to remove a paragraph (and, in markup mode, to mark its paragraph mark deleted so two paragraphs can merge when the change is accepted). `word_comment_add` uses the same resilient single-paragraph matching and bulk-task model for comment placement. These file-level tools report `complete`, `partial`, or `none`; when a task is not found, inspect the per-task unmatched-anchor suggestions, re-read the current document text, and retry only the missed tasks rather than treating the result as blocked.
- **Workspace vs. text tools.** `workspace_*` act inside the connected workspace boundary; `text_*` act on allowed-boundary files by path. Both are available on all three surfaces.
- **Skills vs. agents.** `skill_use` loads instructions into the SHARED conversation; `agent_<name>` delegates a self-contained sub-task in an ISOLATED context.

## Online Sources

The selected online sources must also be included as "allowed tools" if they shall be available to a skill or agent. Wildcards (for example `swiss-caselaw*`) can be used, as well as the universal placeholder `selected_online_sources`.
