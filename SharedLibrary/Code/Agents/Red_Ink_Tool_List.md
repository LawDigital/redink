# Red Ink Tool List

This file lists the built-in internal tools registered in `SharedLibrary\Code\Agents\HostToolRegistration.vb`.

## Shared tools

| Tool | Description | Host(s) |
|---|---|---|
| `retrieve_web_content` | Retrieves readable text, and optionally links, from one or more public web URLs. | Word, Outlook |
| `web_content_retriever` | Alias for the public web content retrieval tool. | Word, Outlook |
| `download_web_files` | Downloads remote files and saves the original binary files locally. | Word, Outlook |
| `internet_search` | Searches the public internet and returns readable content from top results. | Word, Outlook |
| `web_grounding` | Uses a web-enabled model to perform cited live-web research. | Word, Outlook |
| `knowledge_search` | Searches the user's local knowledge store for relevant internal content. | Word, Outlook |
| `tool_loader` | Lazily loads full tool definitions only when a specific tool is needed. | Word, Outlook |
| `memory_put` | Stores a key/value memory entry with summary, tags, and metadata. | Word, Outlook |
| `memory_get` | Retrieves a stored memory entry by key. | Word, Outlook |
| `memory_list` | Lists stored memory entries and their summaries. | Word, Outlook |
| `memory_delete` | Deletes a stored memory entry by key. | Word, Outlook |
| `text_read` | Reads a UTF-8 text file within the allowed workspace boundary. | Word, Outlook |
| `text_write` | Writes, replaces, or appends a UTF-8 text file. | Word, Outlook |
| `text_search` | Searches text files for substring or regex matches. | Word, Outlook |
| `js_run` | Executes sandboxed JavaScript in a hidden WebView2 environment. | Word, Outlook |
| `skill_use` | Loads a skill's instructions and file inventory for guided execution. | Word, Outlook |
| `m365_search` | Searches Microsoft 365 content such as mail, files, chats, events, and notes. | Word, Outlook |
| `m365_get_mail` | Retrieves a mail message and its attachment text. | Word, Outlook |
| `m365_get_mail_thread` | Retrieves an entire mail conversation as one transcript. | Word, Outlook |
| `m365_get_file` | Retrieves a Microsoft 365 file and extracts its readable text. | Word, Outlook |
| `m365_get_event` | Retrieves calendar event details. | Word, Outlook |
| `m365_get_chat_thread` | Retrieves a Teams chat or channel thread. | Word, Outlook |
| `m365_get_onenote_page` | Retrieves a OneNote page and returns readable content. | Word, Outlook |
| `word_extract_text` | Extracts plain text from a `.docx` file on disk. | Word, Outlook |
| `word_search` | Searches a `.docx` file on disk for text or regex matches. | Word, Outlook |
| `word_write` | Inserts, replaces, or appends plain text in a `.docx` file on disk. | Word, Outlook |
| `word_markup` | Edits a `.docx` file on disk using tracked-change style markup. | Word, Outlook |
| `word_comment_add` | Adds a Word comment to a matched span in a `.docx` file on disk. | Word, Outlook |
| `word_comment_list` | Lists comments in a `.docx` file on disk. | Word, Outlook |
| `word_comment_remove` | Removes comments from a `.docx` file on disk. | Word, Outlook |
| `word_format` | Applies paragraph or run formatting to matched text in a `.docx` file on disk. | Word, Outlook |
| `word_apply_template` | Creates a document from a template with substitutions. | Word, Outlook |
| `word_save_as` | Saves a `.docx` file to a new path. | Word, Outlook |

## Outlook-only tools

| Tool | Description | Host(s) |
|---|---|---|
| `process_word_document` | Processes a Word document or attachment through the document-processing pipeline. | Outlook |
| `comment_word_document` | Adds comment bubbles to a Word document. | Outlook |
| `extract_pdf_text` | Extracts readable text from a PDF file. | Outlook |
| `merge_pdfs` | Merges multiple PDFs into one output PDF. | Outlook |
| `read_attachment` | Reads or extracts text from an email attachment. | Outlook |
| `list_attachments` | Lists the attachments available in the current email context. | Outlook |
| `describe_binary_attachment` | Produces a description of a non-text attachment. | Outlook |
| `compare_word_documents` | Compares two Word documents and reports differences. | Outlook |
| `read_word_document_details` | Returns metadata or structural details about a Word document. | Outlook |
| `create_pdf_from_text` | Generates a PDF from supplied text content. | Outlook |
| `extract_excel_data` | Extracts readable or structured data from an Excel file using the normal reader path. For live form-filling scenarios with formulas, dropdowns, recalculation, or protected sheets, prefer the live Excel Interop tools below. | Outlook |
| `excel_list_live_worksheets` | Lists worksheet names and basic live worksheet metadata from an existing Excel file through Excel Interop. | Outlook |
| `excel_read_live_range` | Reads a worksheet or range from an existing Excel file through live Excel Interop, including current values and workbook-state-dependent details. | Outlook |
| `excel_complete_live_workbook` | Updates an existing Excel workbook through live Excel Interop and saves a new `_completed.xlsx` copy, including LiftLock handling where present. | Outlook |
| `split_pdf` | Splits a PDF into multiple output files. | Outlook |
| `add_pdf_watermark` | Applies a watermark to a PDF. | Outlook |
| `word_to_pdf` | Converts a Word document to PDF. | Outlook |
| `search_in_attachments` | Searches across attachment content for relevant matches. | Outlook |
| `summarize_thread` | Summarizes an email thread. | Outlook |
| `pdf_to_word` | Converts a PDF into a Word document. | Outlook |
| `create_word_document` | Creates a new Word document output. | Outlook |
| `create_excel_spreadsheet` | Creates a new Excel workbook output. | Outlook |
| `create_powerpoint` | Creates a new PowerPoint presentation output. | Outlook |
| `create_code_file` | Creates a source code or text-based file output. | Outlook |
| `comment_pdf_document` | Adds annotation comments to a PDF. | Outlook |
| `extract_data_from_attachments` | Pulls structured information from one or more attachments. | Outlook |
| `redact_pdf` | Redacts content in a PDF. | Outlook |
| `overlay_pdf` | Overlays one PDF onto another PDF. | Outlook |
| `create_audio_file` | Generates an audio file output. | Outlook |
| `generate_image` | Generates an image file output. | Outlook |
| `manage_scheduled_tasks` | Creates, lists, updates, pauses, resumes, or deletes scheduled tasks. | Outlook |
| `manage_user_memory` | Manages per-user persistent memory storage. | Outlook |
| `manage_user_files` | Manages files in per-user storage. | Outlook |
| `report_inability` | Returns a structured inability report when the requested action cannot be completed. | Outlook |
| `agent_workspace_list` | Lists files and folders in the chat-agent workspace. | Outlook |
| `agent_workspace_read` | Reads or extracts text from a workspace file. | Outlook |
| `agent_workspace_write` | Writes a text or code file into the workspace. | Outlook |
| `agent_workspace_file_op` | Performs safe file operations such as copy, move, rename, create folder, or delete inside the workspace. | Outlook |
| `agent_workspace_save_session_file` | Copies a session-produced file into the workspace. | Outlook |
| `agent_workspace_search` | Searches workspace filenames and text-like content. | Outlook |
| `agent_workspace_find_files` | Finds workspace files by name, extension, size, or modified date. | Outlook |
| `agent_workspace_move_to` | Moves one or more workspace items into another folder. | Outlook |
| `agent_workspace_copy_to` | Copies one or more workspace items into another folder. | Outlook |
| `agent_workspace_rename` | Renames a workspace file or folder. | Outlook |
| `agent_workspace_bulk_rename` | Renames many workspace files using batch rules. | Outlook |
| `agent_workspace_file_details` | Returns detailed metadata for a workspace file or folder. | Outlook |
| `agent_workspace_recent_files` | Lists recently changed workspace files. | Outlook |
| `agent_workspace_create_folder_structure` | Creates multiple folders under a workspace path in one operation. | Outlook |
| `agent_workspace_trash` | Moves workspace files or folders to the Recycle Bin. | Outlook |
| `agent_workspace_inventory_report` | Creates a Word or Excel inventory report for workspace files. | Outlook |

## Word-only tools

| Tool | Description | Host(s) |
|---|---|---|
| `workspace_get` | Returns the current workspace path and permissions. | Word |
| `workspace_inventory` | Lists files in the workspace with optional recursion and filtering. | Word |
| `workspace_read` | Reads a UTF-8 text file from the workspace. | Word |
| `workspace_read_many` | Reads multiple text files from the workspace in one call. | Word |
| `workspace_write` | Writes, appends, or creates a text file in the workspace. | Word |
| `workspace_search` | Searches across workspace file contents. | Word |
| `workspace_copy` | Copies a file or folder within the workspace. | Word |
| `workspace_move` | Moves a file or folder within the workspace. | Word |
| `workspace_rename` | Renames a file or folder within the workspace. | Word |
| `workspace_delete` | Deletes a file or folder from the workspace. | Word |
| `workspace_make_dir` | Creates a folder in the workspace. | Word |
| `workspace_extract_text` | Extracts readable text from a supported workspace file such as PDF, Word, or Excel. | Word |
| `workspace_extract_text_many` | Extracts readable text from multiple supported workspace files. | Word |
| `worddoc_list_open` | Lists the documents currently open in Word. | Word |
| `worddoc_get_active` | Returns metadata for the active Word document. | Word |
| `worddoc_extract_text` | Extracts plain text from the active or a named open Word document. | Word |
| `worddoc_search` | Searches the active or a named open Word document. | Word |
| `worddoc_list_comments` | Lists comments in the active or a named open Word document. | Word |
| `worddoc_insert_text` | Inserts text into the active or a named open Word document. | Word |
| `worddoc_replace` | Replaces text in the active or a named open Word document. | Word |
| `worddoc_comment_add` | Adds a comment to matched text in the active or a named open Word document. | Word |
| `worddoc_format` | Applies formatting to matched text in the active or a named open Word document. | Word |
| `word_doc_read` | Reads content from the active Word document through the Word host bridge. | Word |
| `word_doc_edit` | Edits the active Word document through the Word host bridge. | Word |
| `word_doc_create` | Creates a new Word document through the Word host bridge. | Word |
| `word_doc_export_pdf` | Exports a Word document to PDF through the Word host bridge. | Word |

## Online Sources

The selected online sources must also be included as "allowed tools" is they shall be available to a skill or agent.
Wild cards (e.g., swiss-caselaw*) can be used as well as the universal placeholder "selected_online_sources".
