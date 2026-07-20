# Red Ink Tool List

This file lists the built-in internal tools that can be advertised by Word, Outlook, and Outlook AutoPilot.

Notes:

- Availability can still depend on configuration, feature flags, model support, and current execution mode.
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
| `memory_put` | Stores a key/value memory entry with summary, tags, and metadata. | Yes | Yes | No |
| `memory_get` | Retrieves a stored memory entry by key. | Yes | Yes | No |
| `memory_list` | Lists stored memory entries and their summaries. | Yes | Yes | No |
| `memory_delete` | Deletes a stored memory entry by key. | Yes | Yes | No |
| `text_read` | Reads a UTF-8 text file within the allowed workspace boundary. | Yes | Yes | Yes |
| `text_write` | Writes, replaces, or appends a UTF-8 text file. | Yes | Yes | Yes |
| `text_search` | Searches text files for substring or regex matches. | Yes | Yes | Yes |
| `js_run` | Executes sandboxed JavaScript in a hidden WebView2 environment. | Yes | Yes | Yes |
| `skill_use` | Loads a skill's instructions and file inventory for guided execution. | Yes | Yes | No |
| `m365_search` | Searches Microsoft 365 content such as mail, files, chats, events, and notes. | Yes | Yes | No |
| `m365_get_mail` | Retrieves a mail message and its attachment text. | Yes | Yes | No |
| `m365_get_mail_thread` | Retrieves an entire mail conversation as one transcript. | Yes | Yes | No |
| `m365_get_file` | Retrieves a Microsoft 365 file and extracts its readable text. | Yes | Yes | No |
| `m365_get_event` | Retrieves calendar event details. | Yes | Yes | No |
| `m365_get_chat_thread` | Retrieves a Teams chat or channel thread. | Yes | Yes | No |
| `m365_get_onenote_page` | Retrieves a OneNote page and returns readable content. | Yes | Yes | No |
| `word_extract_text` | Extracts plain text from a `.docx` file on disk. | Yes | Yes | Yes |
| `word_search` | Searches a `.docx` file on disk for text or regex matches. | Yes | Yes | Yes |
| `word_write` | Inserts, replaces, or appends plain text in a `.docx` file on disk. | Yes | Yes | Yes |
| `word_markup` | Edits a `.docx` file on disk using tracked-change style markup. | Yes | Yes | Yes |
| `word_comment_add` | Adds a Word comment to a matched span in a `.docx` file on disk. | Yes | Yes | Yes |
| `word_comment_list` | Lists comments in a `.docx` file on disk. | Yes | Yes | Yes |
| `word_comment_remove` | Removes comments from a `.docx` file on disk. | Yes | Yes | Yes |
| `word_format` | Applies paragraph or run formatting to matched text in a `.docx` file on disk. | Yes | Yes | Yes |
| `word_apply_template` | Creates a document from a template with substitutions. | Yes | Yes | Yes |
| `word_save_as` | Saves a `.docx` file to a new path. | Yes | Yes | Yes |

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
| `agent_workspace_save_session_file` | Copies a session-produced file into the workspace. | No | Yes | Yes |
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
| `worddoc_insert_text` | Inserts text into the active or a named open Word document. | Yes | No | No |
| `worddoc_replace` | Replaces text in the active or a named open Word document. | Yes | No | No |
| `worddoc_comment_add` | Adds a comment to matched text in the active or a named open Word document. | Yes | No | No |
| `worddoc_format` | Applies formatting to matched text in the active or a named open Word document. | Yes | No | No |
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

## Online Sources

The selected online sources must also be included as "allowed tools" if they shall be available to a skill or agent. Wildcards (for example `swiss-caselaw*`) can be used, as well as the universal placeholder `selected_online_sources`.
