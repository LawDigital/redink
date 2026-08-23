' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Tools.vb
' Purpose:
'   AutoPilot tool-surface aggregator for Outlook: assembles built-in tool definitions,
'   normalizes results, tracks produced files and delegates execution to modular Office,
'   PDF, data-collection and miscellaneous tool implementations.
'
' Architecture / Function:
'   - Registration combines host tools with the shared registry/lazy-loading model; tool
'     names and schemas remain stable model-facing contracts.
'   - Execution is delegated to the owning module and returns normalized ToolResponse
'     payloads into the same sequencing/finality pipeline used by Local Chat.
'   - Produced artifacts are registered with explicit artifact/deliverable semantics so
'     later tool calls and final delivery do not rely on filename heuristics.
'   - Uses per-run AutoPilot attachment/temp/session state but does not own the Office
'     rendering algorithms implemented in Tools.Office* or the shared agent tools.
' =============================================================================



Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Xml
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL NAMES (constants for matching)
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Const AP_Tool_ProcessWordDoc As String = "process_word_document"
    Private Const AP_Tool_CommentWordDoc As String = "comment_word_document"
    Private Const AP_Tool_ExtractPdfText As String = "extract_pdf_text"
    Private Const AP_Tool_MergePdfs As String = "merge_pdfs"
    Private Const AP_Tool_ReadAttachment As String = "read_attachment"
    Private Const AP_Tool_ListAttachments As String = "list_attachments"
    Private Const AP_ToolPrefix As String = "autopilot_"
    Private Const AP_Tool_CompareWordDocs As String = "compare_word_documents"
    Private Const AP_Tool_ReadWordDocDetails As String = "read_word_document_details"
    Private Const AP_Tool_CreatePdfFromText As String = "create_pdf_from_text"
    Private Const AP_Tool_DescribeBinary As String = "describe_binary_attachment"
    Private Const AP_Tool_ExtractExcelData As String = "extract_excel_data"
    Private Const AP_Tool_ExcelListLiveWorksheets As String = "excel_list_live_worksheets"
    Private Const AP_Tool_ExcelReadLiveRange As String = "excel_read_live_range"
    Private Const AP_Tool_ExcelCompleteLiveWorkbook As String = "excel_complete_live_workbook"
    Private Const AP_Tool_SplitPdf As String = "split_pdf"
    Private Const AP_Tool_AddPdfWatermark As String = "add_pdf_watermark"
    Private Const AP_Tool_WordToPdf As String = "word_to_pdf"
    Private Const AP_Tool_SearchInAttachments As String = "search_in_attachments"
    Private Const AP_Tool_SummarizeThread As String = "summarize_thread"
    Private Const AP_Tool_PdfToWord As String = "pdf_to_word"
    Private Const AP_Tool_CreateWordDoc As String = "create_word_document"
    Private Const AP_Tool_CreateExcel As String = "create_excel_spreadsheet"
    Private Const AP_Tool_CreatePowerPoint As String = "create_powerpoint"
    Private Const AP_Tool_CreateCodeFile As String = "create_code_file"
    Private Const AP_Tool_CommentPdf As String = "comment_pdf_document"
    Private Const AP_Tool_ExtractDataFromAttachments As String = "extract_data_from_attachments"
    Private Const AP_Tool_RedactPdf As String = "redact_pdf"
    Private Const AP_Tool_OverlayPdf As String = "overlay_pdf"
    Private Const AP_Tool_CreateAudioFile As String = "create_audio_file"
    Private Const AP_Tool_GenerateImage As String = "generate_image"
    Private Const AP_Tool_WebGrounding As String = "web_grounding"
    Private Const AP_Tool_ManageScheduledTasks As String = "manage_scheduled_tasks"
    Private Const AP_Tool_ManageUserMemory As String = "manage_user_memory"
    Private Const AP_Tool_ManageUserFiles As String = "manage_user_files"
    Private Const AP_Tool_ListCollectionUseCases As String = "list_collection_use_cases"
    Private Const AP_Tool_CollectData As String = "collect_data"
    Private Const AP_Tool_PreviewCollection As String = "preview_collection"
    Private Const AP_Tool_ReportInability As String = "report_inability"
    Private Const AP_Tool_CompleteWordTables As String = "complete_word_tables"


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL REGISTRATION
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Function BuildManageScheduledTasksTool() As ModelConfig
        Return New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ManageScheduledTasks,
            .ModelDescription = "Manage Scheduled Tasks (built-in)",
                       .ToolInstructionsPrompt =
                AP_Tool_ManageScheduledTasks & ": Manages the AutoPilot task scheduler. " &
                "Users can schedule tasks to be executed automatically at specific times (one-time or recurring) " &
                "with results delivered by e-mail or, in Local Chat mode, through the Local Agent browser workflow after user confirmation. " &
                "Supports creating, listing, querying, updating, and deleting scheduled tasks. " &
                "If the user asks to schedule, remind, repeat, recur, or do something every X minutes/hours/days, you SHOULD use this tool instead of claiming that scheduling is unavailable. " &
                "The user gives natural-language scheduling instructions like 'every Monday at 8am', " &
                "'every first Sunday of the month at 10:00', 'three times starting tomorrow at 14:00 every 2 days', " &
                "'between now and end of October every second day at 09:00', or 'tell me a joke every five minutes'. " &
                "You MUST translate these into structured fields: " &
                "- schedule_description: the human-readable schedule text as stated by the user " &
                "- rrule: an iCalendar RRULE string (e.g. FREQ=WEEKLY;INTERVAL=1;BYDAY=MO or FREQ=MONTHLY;INTERVAL=1;BYDAY=1SU or FREQ=DAILY;INTERVAL=2) " &
                "- time_of_day_local: the local time in HH:mm format (e.g. '08:00', '14:30') " &
                "- next_due_utc: the ISO 8601 UTC timestamp of the FIRST execution " &
                "- end_date_utc: the ISO 8601 UTC end date if specified (otherwise omit) " &
                "- remaining_occurrences: number of times to execute if count-limited (e.g. 'three times' → 3), otherwise 0 for unlimited " &
                "Interpret all schedule phrases relative to the CURRENT LOCAL time on this machine, not UTC. " &
                "Use UTC only for the internal next_due_utc and end_date_utc fields; keep all user-facing wording and reasoning in local time only. " &
                "When the user says 'every Monday at 8am' and today is Wednesday, the next_due_utc should be next Monday at 08:00 local time converted to UTC. " &
                "The machine timezone offset is used for UTC conversion (current local time: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm") & ", " &
                "UTC offset: " & DateTimeOffset.Now.Offset.ToString() & "). " &
                "For the 'list' action, return ALL tasks including their IDs, instructions, schedules, and status. " &
                "For 'delete' or 'update', match by task ID prefix or instruction text. " &
                "The deliver_to field is the e-mail address(es) for result delivery. " &
                "When invoked from an e-mail, use the sender's address as deliver_to unless the user specifies otherwise. " &
                "When invoked from Local Chat, " &
                If(INI_AutoPilotSchedulerLocalChat,
                   "deliver_to may be omitted — default it to the current mailbox address and use e-mail delivery rather than the Local Agent browser prompt workflow. ",
                   "deliver_to may be omitted — the task will run in the Local Agent browser workflow after user confirmation rather than by sending the result by e-mail. ") &
                "Tasks can reference attached files — use store_attachment_names to copy the current e-mail's attachments " &
                "into the task's permanent storage for use during execution.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ManageScheduledTasks & """," &
                """description"":""Manages the AutoPilot task scheduler. Supports creating, listing, querying, updating, and deleting " &
                "scheduled tasks that execute automatically and deliver results by e-mail or, in Local Chat mode, via the Local Agent browser workflow after user confirmation. " &
                "Use this when the user asks to schedule, remind, repeat, recur, or run something every X minutes/hours/days. " &
                "Translate natural-language schedules into structured rrule/time_of_day_local/next_due_utc fields.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """action"":{""type"":""string"",""enum"":[""create"",""list"",""get"",""update"",""delete""]," &
                """description"":""The operation to perform""}," &
                """task_id"":{""type"":""string"",""description"":""Task ID (or prefix) for get/update/delete actions""}," &
                """instruction"":{""type"":""string"",""description"":""The task instruction/prompt to execute (for create/update)""}," &
                """subject"":{""type"":""string"",""description"":""Subject line for the result e-mail (for create/update)""}," &
                """deliver_to"":{""type"":""array"",""items"":{""type"":""string""},""description"":""E-mail addresses for result delivery (for create/update)""}," &
                """schedule_description"":{""type"":""string"",""description"":""Human-readable schedule description (e.g. 'every Monday at 08:00')""}," &
                """rrule"":{""type"":""string"",""description"":""iCalendar RRULE string (e.g. 'FREQ=WEEKLY;INTERVAL=1;BYDAY=MO'). Empty for one-time tasks.""}," &
                """time_of_day_local"":{""type"":""string"",""description"":""Local time of execution in HH:mm format (e.g. '08:00')""}," &
                """next_due_utc"":{""type"":""string"",""description"":""ISO 8601 UTC timestamp for the first/next execution (e.g. '2026-03-30T06:00:00Z')""}," &
                """end_date_utc"":{""type"":""string"",""description"":""ISO 8601 UTC end date for recurrence (omit for no end date)""}," &
                """remaining_occurrences"":{""type"":""integer"",""description"":""Number of remaining executions for count-limited tasks (0 = unlimited)""}," &
                """status"":{""type"":""string"",""enum"":[""active"",""paused""],""description"":""Task status (for update)""}," &
                """store_attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of current e-mail attachments to copy into the task's permanent storage""}," &
                """status_filter"":{""type"":""string"",""description"":""Filter for list action (e.g. 'active', 'completed', 'paused'). Omit to list all.""}" &
                "},""required"":[""action""]}}"
        }
    End Function

    ''' <summary>
    ''' Builds and returns the full set of AutoPilot internal tool definitions.
    ''' </summary>
    ''' <returns>
    ''' A list of <see cref="ModelConfig"/> items representing built-in tools
    ''' registered for the current AutoPilot run.
    ''' </returns>
    Friend Function GetAutoPilotInternalTools() As List(Of ModelConfig)
        Dim tools As New List(Of ModelConfig)()

        ' Deterministic helper for exact text/data computation inside AutoPilot.
        Dim jsRunTool As ModelConfig = SharedLibrary.Agents.JsRunTool.Build(_context)
        If jsRunTool IsNot Nothing Then
            tools.Add(jsRunTool)
        End If

        ' Shared file/workspace tools available inside the current AutoPilot workspace.
        tools.AddRange(SharedLibrary.Agents.TextTools.BuildAll())
        tools.AddRange(SharedLibrary.Agents.FileTools.BuildAll())
        tools.AddRange(SharedLibrary.Agents.WordTools.BuildAll())
        tools.AddRange(SharedLibrary.Agents.BrowserTools.BuildAll(_context))
        tools.AddRange(SharedLibrary.Agents.WorkspaceTools.BuildAll())
        tools.AddRange(GetAutoPilotAgentWorkspaceTools())

        ' python_execute: secure sandboxed Python execution.
        ' Only advertised when INI_PythonAgentPath is set, the exe is available, and
        ' (when requested) its authenticity has been verified.
        Dim pythonExecuteTool As ModelConfig = Nothing
        If TryConfigureAndBuildPythonExecuteTool(pythonExecuteTool) Then
            tools.Add(pythonExecuteTool)
        End If

        ' Shared helper tools used by the tooling runtime.
        tools.Add(SharedLibrary.Agents.ToolDescribeTool.Build())
        tools.Add(SharedLibrary.Agents.ContextExpandTool.Build())
        tools.Add(SharedLibrary.Agents.ContextCompactTool.Build())

        ' ── process_word_document ──
        tools.Add(New ModelConfig() With {
        .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ProcessWordDoc,
        .ModelDescription = "Process Word/PowerPoint/Excel Document (built-in)",
        .ToolInstructionsPrompt =
            AP_Tool_ProcessWordDoc & ": PREFERRED for transforming an EXISTING Word (.docx), PowerPoint (.pptx), or Excel (.xlsx) attachment while preserving the source document's structure and formatting. " &
            "Use this for translation, correction, proofreading, anonymization, replacements, data updates, formula changes, or any other transformation of an existing Office file. " &
            "If the user asks for a translated/corrected/revised/new-language VERSION of an existing Word document, use this tool rather than create_word_document. " &
            "For Word documents, returns both a clean version and a compare document showing changes. " &
            "For PowerPoint and Excel files, returns the processed version (no compare document). " &
            "For Excel files, you can optionally restrict processing to specific sheet names using the sheet_names parameter. " &
            "CHUNKING / GLOBAL CONSISTENCY (IMPORTANT): The processor model does NOT receive the whole document at once. " &
            "It processes the document sequentially in bounded chunks with only a small nearby context window, and the individual processor calls do not share hidden memory. " &
            "Therefore YOU, the calling model, are responsible for document-wide consistency. When the task depends on information or choices that must remain stable across the whole document, " &
            "put those facts and decisions in global_context so the same compact guidance is repeated to every chunk. Examples: defined-term translations, names, terminology, " &
            "target-language variant, tone/register, abbreviation policy, recurring replacements, or terms that must not be translated. " &
            "For very long documents, keep global_context bounded to reusable cross-document rules and decisions; do NOT paste the document or a long document summary into it. " &
            "Do not assume the processor can infer information that appears only elsewhere in the document, and do not invent global facts that are not supported by the user request or known source context. " &
            "global_context is context, not a second operation. " &
            "For a successful Word transformation, the tool returns TWO user-relevant files: the clean processed document and the tracked-changes compare document. In the final user-facing answer, identify BOTH generated files by name; do not collapse them into a generic reference to one processed document. " &
            "CRITICAL — ONE OPERATION PER CALL: This tool applies exactly ONE instruction per call. " &
            "If the user requests multiple distinct operations (e.g., 'correct and translate', 'anonymize and summarize', 'fix grammar then make more concise'), " &
            "you MUST split them into separate sequential calls. First call: apply the first operation to the original file. " &
            "Wait for the result. Second call: apply the second operation to the output file from the first call (the '_processed' file). " &
            "Example for 'correct and translate to German': " &
            "(1) Call process_word_document with task_type='correct', instruction='Correct spelling, grammar and style' on 'Contract.docx'. Result: 'Contract_processed.docx'. " &
            "(2) Call process_word_document with task_type='translate', instruction='Translate to German' on attachment_names=['Contract_processed.docx']. Result: 'Contract_processed_processed.docx'. " &
            "NEVER combine two distinct operations into a single instruction string. " &
            "However, a single coherent task counts as one operation (e.g., 'Translate to German' is one operation even though it involves reading and rewriting). " &
            "Output files are named '<original>_processed.<ext>' and can be referenced in subsequent tool calls by that name.",
        .ToolDefinition =
            "{""name"":""" & AP_Tool_ProcessWordDoc & """," &
            """description"":""Preferred for transforming an EXISTING Word (.docx), PowerPoint (.pptx), or Excel (.xlsx) attachment while preserving its structure and formatting. Applies exactly ONE processing instruction per call. " &
            "Use for translation, correction, proofreading, anonymization, replacements, data updates, formula modifications, and other transformations of an existing Office file. " &
            "For Word documents, produces clean output plus a compare document with tracked changes. " &
            "For PowerPoint and Excel, produces the processed file only. " &
            "The processor works chunk-by-chunk and does not see the full document in one model call or retain hidden memory between chunks. The calling model is responsible for document-wide consistency: " &
            "when stable terminology, names, translation choices, tone, abbreviations, recurring replacements, or other cross-document facts matter, provide a compact reusable global_context. " &
            "For successful Word processing, surface both returned files (clean processed document and compare document) to the user. " &
            "IMPORTANT: Apply only ONE operation per call. For multi-step requests (e.g. 'correct and translate'), " &
            "make separate sequential calls — first correct, then translate the corrected output file. " &
            "Output files are named '<original>_processed.<ext>' and can be used as input for the next call via attachment_names.""," &
            """parameters"":{""type"":""object"",""properties"":{" &
            """instruction"":{""type"":""string"",""description"":""A single, specific instruction to apply to the document. Must be ONE operation only — " &
            "e.g. 'Translate to German' or 'Correct spelling and grammar' or 'Anonymize all personal names'. " &
            "Do NOT combine multiple operations like 'Correct and translate'. Split those into separate calls.""}," &
            """global_context"":{""type"":""string"",""description"":""Optional document-wide consistency context repeated to EVERY processing chunk. " &
            "Use this when all chunks must know stable facts or decisions that may not appear in their local text window, such as defined-term translations, names, terminology, " &
            "target-language variant, tone/register, abbreviation policy, recurring replacements, or do-not-translate terms. " &
            "The processor never sees the complete document in one model call, so the calling model is responsible for supplying material cross-document guidance here. " &
            "Keep this compact even for very long documents: include reusable rules/decisions, not the document itself or a long summary. Do not invent unsupported facts. " &
            "This is context only and does not count as a second operation.""}," &
            """task_type"":{""type"":""string"",""enum"":[""translate"",""correct"",""other""]," &
            """description"":""Classifies the operation: 'translate' for language translation, 'correct' for spelling/grammar/style correction or proofreading, " &
            "'other' for everything else (anonymization, data transformation, restructuring, summarization, etc.). Default: 'other'""}," &
            """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of the attachments to process. " &
            "Can include output files from previous tool calls (e.g. 'Contract_processed.docx'). " &
            "If empty or omitted, processes all .docx, .pptx, and .xlsx attachments.""}," &
            """sheet_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Optional: for Excel files only, restrict processing to these sheet names. If omitted, all sheets are processed.""}" &
            "},""required"":[""instruction""]}}"
    })

        ' ── extract_pdf_text ──
        tools.Add(New ModelConfig() With {
        .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ExtractPdfText,
        .ModelDescription = "Extract PDF Text (built-in)",
        .ToolInstructionsPrompt =
            AP_Tool_ExtractPdfText & ": Extracts content from one or more PDF attachments. " &
            "By default it returns plain text. If the user wants Markdown, pass output_format='markdown'. " &
            "For text-based PDFs, Markdown uses the direct PDF-to-Markdown extractor without OCR. " &
            "If OCR is needed and available, the OCR path is retained and should return Markdown when Markdown was requested.",
        .ToolDefinition =
            "{""name"":""" & AP_Tool_ExtractPdfText & """," &
            """description"":""Extracts content from PDF file attachments. " &
            "Optional output_format='markdown' returns Markdown instead of plain text when possible.""," &
            """parameters"":{""type"":""object"",""properties"":{" &
            """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of the PDF attachments to extract text from. If empty, processes all PDFs.""}," &
            """output_format"":{""type"":""string"",""enum"":[""text"",""markdown""],""description"":""Optional output format. Default is 'text'. Use 'markdown' to request Markdown output.""}" &
            "},""required"":[]}}"
    })

        ' ── merge_pdfs ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_MergePdfs,
            .CapabilityTags = "explicit_operation",
            .ModelDescription = "Merge PDFs (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_MergePdfs & ": Merges multiple PDF attachments into a single PDF file.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_MergePdfs & """," &
                """description"":""Merges multiple PDF file attachments into a single combined PDF""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of the PDF attachments to merge, in order. If empty, merges all PDFs.""}," &
                """output_filename"":{""type"":""string"",""description"":""Filename for the merged output PDF (default: merged.pdf)""}" &
                "},""required"":[]}}"
        })

        ' ── read_attachment ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ReadAttachment,
            .ModelDescription = "Read Attachment Content (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ReadAttachment & ": Reads and returns the content of one or more supported attachments. " &
                "By default it returns plain text. If the user wants Markdown, pass output_format='markdown'. " &
                "For Word documents (.docx), Markdown uses the sandboxed DOCX-to-Markdown path. " &
                "For PDFs (.pdf), Markdown uses the direct PDF-to-Markdown extractor without OCR unless OCR is explicitly requested by another tool. " &
                "Other supported formats continue to return plain text even when Markdown is requested. " &
                "Embedded mail files (.msg, .eml) are automatically unpacked — their body text and nested attachments " &
                "are extracted recursively and appear as separate attachments that you can reference by name. " &
                "Use attachment_name for a single file or attachment_names for batch reading.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ReadAttachment & """," &
                """description"":""Reads and returns the content of one or more attachment files. " &
                "Supports Word documents (.docx), PDFs (.pdf), Excel spreadsheets (.xlsx, .xls), " &
                "PowerPoint presentations (.pptx), and text-based files (.txt, .csv, .html, .xml, .json, .md, .log). " &
                "Optional output_format='markdown' enables Markdown output for DOCX and PDF files; other formats remain plain text. " &
                "Embedded mail files (.msg, .eml) are automatically unpacked at intake — their text content " &
                "and nested attachments appear as separate files in the attachment list. " &
                "For Word documents, also reports if comments or tracked changes are present.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of a single attachment to read""}," &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of multiple attachments to read in batch. Use this instead of attachment_name when reading several files.""}," &
                """output_format"":{""type"":""string"",""enum"":[""text"",""markdown""],""description"":""Optional output format. Default is 'text'. Use 'markdown' to request Markdown output for DOCX and PDF attachments.""}" &
                "},""required"":[]}}"
        })

        ' ── list_attachments ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ListAttachments,
            .ModelDescription = "List Attachments (built-in)",
            .ToolInstructionsPrompt =
                If(_chatAgentActive,
                   AP_Tool_ListAttachments & ": Lists files currently loaded into the active Local Agent session. " &
                   "IMPORTANT: This does NOT list files that merely exist in the connected workspace. " &
                   "Use agent_workspace_list to inspect workspace files and agent_workspace_stage to load them first.",
                   AP_Tool_ListAttachments & ": Lists all attachments of the current email with name, type, and size."),
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ListAttachments & """," &
                """description"":""" &
                If(_chatAgentActive,
                   "Lists files currently loaded into the Local Agent session. Does not list unstaged workspace files; use agent_workspace_list and agent_workspace_stage for workspace access.",
                   "Lists all email attachments with their filename, type, size, and processing status") &
                """," &
                """parameters"":{""type"":""object"",""properties"":{},""required"":[]}}"
        })

        ' ── describe_binary_attachment ──
        Dim apiCallObj As String = If(_apConfig IsNot Nothing AndAlso _apConfig.UseSecondApi,
                                      INI_APICall_Object_2, INI_APICall_Object)
        If Not String.IsNullOrWhiteSpace(apiCallObj) Then
            tools.Add(New ModelConfig() With {
                .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_DescribeBinary,
                .ModelDescription = "Describe or transcribe a binary attachment (built-in)",
                .ToolInstructionsPrompt =
                    AP_Tool_DescribeBinary & ": Sends a binary attachment (image, audio, voicemail, video) to the AI for description or transcription. " &
                    "Do NOT use this for every image — only when the user explicitly asks about an attachment or when it appears to be substantive content. " &
                    "Common footer/signature images should be ignored.",
                .ToolDefinition =
                    "{""name"":""" & AP_Tool_DescribeBinary & """," &
                    """description"":""Sends a binary attachment (image, audio file, voicemail, video) directly to the AI for description, transcription, or analysis. " &
                    "Use for .png, .jpg, .mp3, .wav, .m4a, .mp4, etc. Do NOT use for footer/signature images.""," &
                    """parameters"":{""type"":""object"",""properties"":{" &
                    """attachment_name"":{""type"":""string"",""description"":""Filename of the binary attachment to analyze""}," &
                    """prompt"":{""type"":""string"",""description"":""Instructions for the AI (e.g. 'describe this image', 'transcribe this voicemail')""}" &
                    "},""required"":[""attachment_name"",""prompt""]}}"
            })
        End If

        ' ── comment_word_document ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CommentWordDoc,
            .ModelDescription = "Comment Word Document (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CommentWordDoc & ": Adds review comments (Word comment bubbles) to a Word (.docx) attachment. " &
                "Use this when the user wants the document annotated, commented, reviewed with margin notes, " &
                "or marked up with feedback directly inside the document as Word comment bubbles. " &
                "Do NOT use this when the user wants a textual summary or analysis — only when comments " &
                "should appear as annotations within the document itself. " &
                "Supports an optional author parameter: if the user asks for comments under a specific name " &
                "(e.g. the sender's name), pass it as author. If not specified, comments are authored as 'Inky'.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CommentWordDoc & """," &
                """description"":""Adds review comments as Word comment bubbles to a .docx attachment. " &
                "Use when the user wants in-document annotations, margin comments, or review feedback placed directly " &
                "inside the Word file. Do NOT use for plain textual summaries or analyses. " &
                "Supports an optional author name for the comments.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """instruction"":{""type"":""string"",""description"":""The review instruction (e.g., 'Review for legal risks', 'Check for inconsistencies', 'Suggest improvements')""}," &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of the .docx attachments to annotate. If empty or omitted, annotates all .docx attachments.""}," &
                """author"":{""type"":""string"",""description"":""Optional author name for the comment bubbles. Use this when the user requests a specific name (e.g. the sender's name). If omitted, defaults to 'Inky'.""}" &
                "},""required"":[""instruction""]}}"
        })

        ' ── compare_word_documents ── 
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CompareWordDocs,
            .CapabilityTags = "explicit_operation",
            .ModelDescription = "Compare two Word documents (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CompareWordDocs & ": Compares exactly two Word document (.doc/.docx) attachments using Word's " &
                "built-in comparison engine (track changes). Use 'original_filename' for the BASE/earlier/reference " &
                "version and 'revised_filename' for the MODIFIED/newer/changed version. Returns a textual summary of " &
                "revisions found and produces a comparison document with tracked changes attached to the reply. " &
                "This tool requires exactly two attachments to be specified — it cannot compare more than two at once. " &
                "If the sender provides more than two Word documents, ask which two should be compared, or run " &
                "multiple comparisons. " &
                "IMPORTANT: This tool can also accept output files produced by other tools (e.g. a '_processed.docx' " &
                "from process_word_document). Note that process_word_document already produces its own compare " &
                "document automatically — only use compare_word_documents separately when comparing two independently " &
                "provided attachments or when the user explicitly asks for a comparison between specific files.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CompareWordDocs & """," &
                """description"":""Compares exactly two Word documents (.doc/.docx) using Word's built-in comparison (track changes). " &
                "The 'original_filename' is the BASE document (the earlier or reference version). " &
                "The 'revised_filename' is the MODIFIED document (the newer or changed version). " &
                "Returns a textual summary of the differences found and produces a comparison document with tracked changes as a result attachment. " &
                "IMPORTANT: 'original_filename' = the source/baseline; 'revised_filename' = the version that was changed or updated. " &
                "This tool compares exactly two documents per call. " &
                "Can also reference output files from earlier tools (e.g. '_processed.docx' from process_word_document).""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """original_filename"":{""type"":""string"",""description"":""Exact filename of the original/baseline/source Word attachment (the earlier version)""}," &
                """revised_filename"":{""type"":""string"",""description"":""Exact filename of the revised/modified/updated Word attachment (the newer version)""}" &
                "},""required"":[""original_filename"",""revised_filename""]}}"
        })

        ' ── read_word_document_details ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ReadWordDocDetails,
            .ModelDescription = "Read Word Document Details (built-in)",
                  .ToolInstructionsPrompt =
                AP_Tool_ReadWordDocDetails & ": Deep-reads a Word document (.docx) including body text with inline tracked changes, " &
                "comment bubbles (with author, date, and anchored text), headers, footers, footnotes, and endnotes. " &
                "This is a heavier tool — only use it when the user explicitly asks about comments, tracked changes, " &
                "revisions, review history, headers/footers, or footnotes/endnotes. For general content questions, use " & AP_Tool_ReadAttachment & " instead. " &
                "NOTE: The body text returned by this tool is optimized for review markup and does NOT include automatic list/heading numbering, " &
                "multilevel list restarts, resolved fields, or bookmark/cross-reference (REF) resolution. When the user needs the fully rendered " &
                "content — numbered lists, heading numbers, field results, or cross-references — use " & AP_Tool_ReadAttachment & " instead. " &
                "Tracked changes are shown inline using «INS|author|date»...«/INS» and «DEL|author|date»...«/DEL» markers. " &
                "Use tracked_changes_author and tracked_changes_since to filter changes by a specific author or date.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ReadWordDocDetails & """," &
                """description"":""Deep-reads a Word document (.docx) with comments, tracked changes, headers/footers, and footnotes/endnotes. " &
                "Only use when the user asks about comments, revisions, changes, review history, headers, footers, footnotes, or endnotes. " &
                "For general content, use " & AP_Tool_ReadAttachment & " instead. " &
                "The body text here omits automatic list/heading numbering, multilevel list restarts, resolved fields, and cross-references (REF); " &
                "use " & AP_Tool_ReadAttachment & " when those are needed.""," & """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the .docx attachment to read""}," &
                """include_comments"":{""type"":""boolean"",""description"":""Include comment bubbles with author, date, and anchored text (default: true)""}," &
                """include_headers_footers"":{""type"":""boolean"",""description"":""Include headers and footers (default: false)""}," &
                """include_footnotes_endnotes"":{""type"":""boolean"",""description"":""Include footnotes and endnotes (default: false)""}," &
                """include_tracked_changes"":{""type"":""boolean"",""description"":""Include tracked changes as inline markers in the body text (default: true)""}," &
                """tracked_changes_author"":{""type"":""string"",""description"":""Optional: only show tracked changes by this author""}," &
                """tracked_changes_since"":{""type"":""string"",""description"":""Optional: only show tracked changes on or after this date (ISO 8601, e.g. '2026-01-15')""}" &
                "},""required"":[""attachment_name""]}}"
        })

        ' ── create_pdf_from_text ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CreatePdfFromText,
            .CapabilityTags = "explicit_operation",
            .ModelDescription = "Create PDF from Text (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CreatePdfFromText & ": Creates a PDF document from provided text content. " &
                "Use this when the user wants a new PDF created with specific content.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CreatePdfFromText & """," &
                """description"":""Creates a PDF file from provided text content.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """content"":{""type"":""string"",""description"":""The text content for the PDF""}," &
                """output_filename"":{""type"":""string"",""description"":""Filename for the output PDF (default: output.pdf)""}," &
                """title"":{""type"":""string"",""description"":""Optional title displayed at the top of the PDF""}" &
                "},""required"":[""content""]}}"
        })

        ' ── extract_excel_data ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ExtractExcelData,
            .ModelDescription = "Extract Excel Data (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ExtractExcelData & ": Reads data from an Excel attachment (.xlsx/.xls) with control over which sheet to read. " &
                "Returns data in CSV-like tabular format. " &
                "Use this as the normal Excel reader. " &
                "For existing workbooks that must be understood or completed live through Excel Interop " &
                "(for example when formulas, dropdowns, validations, comments, recalculation, or current workbook state matter), " &
                "prefer excel_list_live_worksheets first, then excel_read_live_range, and then excel_complete_live_workbook.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ExtractExcelData & """," &
                """description"":""Reads data from an Excel spreadsheet attachment with sheet selection. Returns tabular data. " &
                "This is the normal Excel reader. For live Excel form completion or inspection of evaluated workbook state, " &
                "prefer excel_list_live_worksheets, excel_read_live_range, and excel_complete_live_workbook.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the Excel attachment""}," &
                """sheet_name"":{""type"":""string"",""description"":""Optional: name of the specific sheet to read. If omitted, reads all sheets.""}" &
                "},""required"":[""attachment_name""]}}"
        })

        ' ── excel_list_live_worksheets ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ExcelListLiveWorksheets,
            .ModelDescription = "List Live Excel Worksheets (built-in)",
            .ToolPriority = 970,
            .ToolInstructionsPrompt =
                AP_Tool_ExcelListLiveWorksheets & ": Lists the worksheets in an existing Excel attachment by opening it through live Excel Interop. " &
                "Use this before reading or completing an existing workbook when the current workbook structure matters. " &
                "This is especially important for Excel forms, templates, or workbooks with formulas, dropdowns, validations, or protected sheets. " &
                "If an agentic loop needs to complete an existing Excel workbook, first call this tool, then excel_read_live_range, then excel_complete_live_workbook.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ExcelListLiveWorksheets & """," &
                """description"":""Lists worksheet names and basic sheet metadata from an existing Excel attachment by opening the workbook through live Excel Interop. " &
                "Prefer this before filling an existing workbook so the agent understands which worksheet to read and complete.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the Excel attachment to inspect""}" &
                "},""required"":[""attachment_name""]}}"
        })

        ' ── excel_read_live_range ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ExcelReadLiveRange,
            .ModelDescription = "Read Live Excel Range (built-in)",
            .ToolPriority = 980,
            .ToolInstructionsPrompt =
                AP_Tool_ExcelReadLiveRange & ": Reads the live contents of an existing Excel attachment through Excel Interop, not through XML/OpenXML parsing. " &
                "Use this for existing workbooks whenever formulas, dropdowns, validations, comments, threaded comments, recalculated values, colors, font properties, common formatting, structure, protection details, conditional formatting, data bars, icon sets, or current workbook state matter. " &
                "For form completion or updating an existing workbook, call excel_list_live_worksheets first, then this tool to understand the live sheet content and options, and only then call excel_complete_live_workbook. " &
                "If worksheet_name is omitted, the first worksheet is used. If range_address is omitted, the worksheet's used range is read. " &
                "By default include_formulas is false, include_color is true, include_font_properties is false, include_formatting is false, include_structure is false, include_protection_details is false, and include_conditional_formatting is false.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ExcelReadLiveRange & """," &
                """description"":""Reads a worksheet or range from an existing Excel attachment through live Excel Interop. " &
                "This is the preferred reader for existing workbooks that may contain active formulas, dropdowns, validations, recalculated values, comments, font properties, colors, common formatting, structure, protection details, conditional formatting, data bars, icon sets, or protected sheets. " &
                "Defaults: first worksheet if worksheet_name is omitted, used range if range_address is omitted, include_formulas=false, include_color=true, include_font_properties=false, include_formatting=false, include_structure=false, include_protection_details=false, include_conditional_formatting=false.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the Excel attachment to read""}," &
                """worksheet_name"":{""type"":""string"",""description"":""Optional worksheet name. If omitted, the first worksheet is used.""}," &
                """range_address"":{""type"":""string"",""description"":""Optional Excel range in A1 notation, for example 'A1:D20'. If omitted, the worksheet's used range is read.""}," &
                """include_formulas"":{""type"":""boolean"",""description"":""Optional. Default false. Include cell formulas in the output.""}," &
                """include_color"":{""type"":""boolean"",""description"":""Optional. Default true. Include font and background color information in the output.""}," &
                """include_font_properties"":{""type"":""boolean"",""description"":""Optional. Default false. Include font properties such as name, size, strikethrough, bold, italic, and underline in the output.""}," &
                """include_formatting"":{""type"":""boolean"",""description"":""Optional. Default false. Include common formatting such as number format, alignment, wrapping, shrink-to-fit, and borders in the output.""}," &
                """include_structure"":{""type"":""boolean"",""description"":""Optional. Default false. Include structural details such as row height, column width, and merged-area information.""}," &
                """include_protection_details"":{""type"":""boolean"",""description"":""Optional. Default false. Include cell protection flags such as locked and formula_hidden.""}," &
                """include_conditional_formatting"":{""type"":""boolean"",""description"":""Optional. Default false. Include conditional-formatting summaries, including data bars and icon sets.""}" &
                "},""required"":[""attachment_name""]}}"
        })

        ' ── excel_complete_live_workbook ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ExcelCompleteLiveWorkbook,
            .ModelDescription = "Complete Live Excel Workbook (built-in)",
            .ToolPriority = 990,
            .ToolInstructionsPrompt =
                AP_Tool_ExcelCompleteLiveWorkbook & ": Completes or updates an existing Excel attachment through Excel Interop. By default it edits the existing workbook in place, so repeated updates accumulate in one file rather than creating a new file per round. Supply output_filename only when a separate copy is wanted, leaving the source workbook untouched. " &
                "Use this for any Excel form completion or update task when formulas, dropdowns, validations, recalculation, protection, current workbook state, or cell and structure formatting matter. " &
                "Do not use normal XML-based Excel editing for such tasks. " &
                "Before filling an existing workbook, first call excel_list_live_worksheets and then excel_read_live_range so the tool loop understands the current workbook structure, live values, available options, and any required formatting. " &
                "Updates are provided as JSON. Each update targets a single cell and can set a value, a formula, a comment, font formatting, fill formatting, common formatting, borders, row height, column width, merge state, and cell protection flags. " &
                "Always batch every cell change for a workbook into a single call by passing all updates in one 'updates' array. Do not make one call per cell; the workbook is opened, recalculated, and saved once per call, so batching many updates together is far more efficient. " &
                "If worksheet_name is omitted, the first worksheet is used. A per-update worksheet_name may override the default worksheet. " &
                "For formulas, prefer English Excel formulas with comma separators when possible. The tool includes locale-safe fallbacks for localized Excel installations and different list separators.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ExcelCompleteLiveWorkbook & """," &
                """description"":""Updates an existing Excel attachment through live Excel Interop. By default the existing workbook is edited in place so repeated updates accumulate in a single file; provide output_filename to instead save the result as a separate copy and leave the source untouched. " &
                "This is the preferred tool for completing existing Excel workbooks that may contain active formulas, dropdowns, validations, recalculation, protected sheets with LiftLock markers, or required formatting such as strikethrough, font color, fill color, number format, alignment, borders, row height, column width, merge state, and cell protection flags. " &
                "Use JSON-based cell updates. Prefer English Excel formulas with comma separators; locale fallbacks are built in.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the Excel attachment to complete or update""}," &
                                """worksheet_name"":{""type"":""string"",""description"":""Optional default worksheet name. If omitted, the first worksheet is used.""}," &
                """output_filename"":{""type"":""string"",""description"":""Optional. If omitted, the existing workbook is edited in place (preferred for iterative edits). If provided, the result is saved as a separate .xlsx copy under this name and the source workbook is left unchanged.""}," &
                """updates"":{""type"":""array"",""description"":""Array of cell updates to apply. Each update targets one cell."",""items"":{""type"":""object"",""properties"":{" &
                """worksheet_name"":{""type"":""string"",""description"":""Optional worksheet override for this update. If omitted, the tool-level worksheet_name or the first worksheet is used.""}," &
                """cell"":{""type"":""string"",""description"":""Target cell address in A1 notation, for example 'B12'""}," &
                """value"":{""description"":""Optional cell value. Use JSON string, number, boolean, or null.""}," &
                """formula"":{""type"":""string"",""description"":""Optional Excel formula for the cell. Prefer English function names and comma separators, starting with '='. Locale-safe fallbacks are applied automatically.""}," &
                """comment"":{""type"":""string"",""description"":""Optional comment text to add as a threaded comment or reply where supported.""}," &
                """font"":{""type"":""object"",""description"":""Optional font properties to set on the target cell."",""properties"":{" &
                """name"":{""type"":""string"",""description"":""Optional font name.""}," &
                """size"":{""description"":""Optional font size as number.""}," &
                """strikethrough"":{""type"":""boolean"",""description"":""Optional. Set or clear strikethrough formatting.""}," &
                """bold"":{""type"":""boolean"",""description"":""Optional. Set or clear bold formatting.""}," &
                """italic"":{""type"":""boolean"",""description"":""Optional. Set or clear italic formatting.""}," &
                """color"":{""description"":""Optional font color. Use '#RRGGBB', an integer OLE color value, or 'automatic'.""}," &
                """underline"":{""description"":""Optional. Use true/false or one of: 'none', 'single', 'double', 'single_accounting', 'double_accounting'.""}" &
                "}}," &
                """fill"":{""type"":""object"",""description"":""Optional fill properties to set on the target cell."",""properties"":{" &
                """color"":{""description"":""Optional background color. Use '#RRGGBB', an integer OLE color value, or 'none' to clear the fill.""}" &
                "}}," &
                """number_format"":{""type"":""string"",""description"":""Optional Excel number format string.""}," &
                """horizontal_alignment"":{""type"":""string"",""description"":""Optional horizontal alignment: general, left, center, right, fill, justify, center_across_selection, distributed.""}," &
                """vertical_alignment"":{""type"":""string"",""description"":""Optional vertical alignment: top, center, bottom, justify, distributed.""}," &
                """wrap_text"":{""type"":""boolean"",""description"":""Optional. Enable or disable wrap text.""}," &
                """shrink_to_fit"":{""type"":""boolean"",""description"":""Optional. Enable or disable shrink to fit.""}," &
                """borders"":{""type"":""object"",""description"":""Optional border settings. Supported properties: top, bottom, left, right. Each side may be 'none' or an object with optional style, weight, and color.""}," &
                """row_height"":{""description"":""Optional row height for the target cell's row.""}," &
                """column_width"":{""description"":""Optional column width for the target cell's column.""}," &
                """merge_action"":{""type"":""string"",""description"":""Optional merge action: 'merge' or 'unmerge'.""}," &
                """merge_range"":{""type"":""string"",""description"":""Optional A1 range for the merge action. If omitted, the target cell or its current merged area is used.""}," &
                """protection"":{""type"":""object"",""description"":""Optional cell protection properties to set on the target cell or merged area."",""properties"":{" &
                """locked"":{""type"":""boolean"",""description"":""Optional. Set or clear the locked flag.""}," &
                """formula_hidden"":{""type"":""boolean"",""description"":""Optional. Set or clear the formula-hidden flag.""}" &
                "}}" &
                "},""required"":[""cell""]}}" &
                "},""required"":[""attachment_name"",""updates""]}}"
        })

        ' ── split_pdf ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_SplitPdf,
            .CapabilityTags = "explicit_operation",
            .ModelDescription = "Split PDF (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_SplitPdf & ": Extracts a range of pages from a PDF attachment into a new PDF.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_SplitPdf & """," &
                """description"":""Extracts a page range from a PDF into a new PDF file.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the PDF attachment""}," &
                """start_page"":{""type"":""integer"",""description"":""First page to extract (1-based)""}," &
                """end_page"":{""type"":""integer"",""description"":""Last page to extract (1-based, inclusive)""}," &
                """output_filename"":{""type"":""string"",""description"":""Filename for the output PDF (default: split.pdf)""}" &
                "},""required"":[""attachment_name"",""start_page"",""end_page""]}}"
        })

        ' ── add_pdf_watermark ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_AddPdfWatermark,
            .CapabilityTags = "explicit_operation",
            .ModelDescription = "Add PDF Watermark (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_AddPdfWatermark & ": Adds a diagonal text watermark to every page of a PDF attachment.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_AddPdfWatermark & """," &
                """description"":""Adds a diagonal text watermark to every page of a PDF.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the PDF attachment""}," &
                """watermark_text"":{""type"":""string"",""description"":""Text for the watermark (e.g., 'DRAFT', 'CONFIDENTIAL')""}," &
                """output_filename"":{""type"":""string"",""description"":""Filename for the output PDF (default: watermarked.pdf)""}" &
                "},""required"":[""attachment_name"",""watermark_text""]}}"
        })

        ' ── word_to_pdf ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_WordToPdf,
            .CapabilityTags = "explicit_operation",
            .ModelDescription = "Convert Word to PDF (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_WordToPdf & ": Converts a Word document (.doc/.docx) attachment to PDF format using Word.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_WordToPdf & """," &
                """description"":""Converts a Word document (.doc/.docx) attachment to PDF format.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the Word document attachment to convert""}" &
                "},""required"":[""attachment_name""]}}"
        })

        ' ── search_in_attachments ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_SearchInAttachments,
            .ModelDescription = "Search in Attachments (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_SearchInAttachments & ": Searches for a keyword or phrase across all readable attachments. " &
                "Returns matching lines with surrounding context.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_SearchInAttachments & """," &
                """description"":""Searches for a keyword or phrase across all readable attachments and returns matching excerpts with context.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """search_term"":{""type"":""string"",""description"":""The keyword or phrase to search for (case-insensitive)""}," &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Optional: limit search to these attachments. If omitted, searches all.""}" &
                "},""required"":[""search_term""]}}"
        })

        ' ── summarize_thread ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_SummarizeThread,
            .ModelDescription = "Summarize Email Thread (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_SummarizeThread & ": Extracts and structures the email conversation thread from the current mail, " &
                "excluding messages sent to/from the monitored AutoPilot mailbox. Returns each message with sender, date, and body.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_SummarizeThread & """," &
                """description"":""Extracts the full email conversation thread, excluding AutoPilot's own replies. Returns structured messages with sender, date, and content.""," &
                """parameters"":{""type"":""object"",""properties"":{},""required"":[]}}"
        })

        ' ── pdf_to_word ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_PdfToWord,
            .CapabilityTags = "explicit_operation",
            .ModelDescription = "Convert PDF to Word (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_PdfToWord & ": Converts a PDF attachment to a Word document (.docx) using Word's built-in PDF import. " &
                "The resulting .docx can then be used with compare_word_documents or other Word tools. " &
                "This is the PREFERRED method for PDF-to-Word conversion — use it FIRST. It works well for most PDFs " &
                "that contain real (selectable/searchable) text and preserves layout, tables, and formatting. " &
                "If the conversion result indicates the PDF is scanned/image-only (no extractable text), THEN " &
                "fall back to extract_pdf_text (which supports OCR) to obtain the text, and use create_word_document " &
                "to produce a .docx from that text. " &
                "ALTERNATIVE APPROACH — USER-REQUESTED OCR PIPELINE: If the user explicitly asks to OCR the PDF first, " &
                "or asks to 'rasterize and OCR' the PDF before converting to Word, use this pipeline instead: " &
                "(1) Call extract_pdf_text on the PDF — this will rasterize each page and run OCR via the LLM to extract text. " &
                "(2) Call create_word_document with the OCR-extracted text to produce a .docx. " &
                "This OCR pipeline is useful for scanned documents, image-heavy PDFs, or when Word's built-in conversion " &
                "produces poor results. However, it does NOT preserve the original layout/formatting — it produces a " &
                "clean text-based document. The standard Word-based conversion (this tool) remains the default.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_PdfToWord & """," &
                """description"":""Converts a PDF attachment to a Word document (.docx) using Word's built-in PDF reflow. " &
                "Use this as the PRIMARY method for PDF-to-Word conversion. Works well for text-based PDFs with layout preservation. " &
                "If the result indicates the PDF is scanned/image-only, fall back to extract_pdf_text (OCR) + create_word_document. " &
                "ALTERNATIVE: If the user explicitly requests OCR-based conversion, use extract_pdf_text (rasterize+OCR) followed by create_word_document instead.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the PDF attachment to convert""}," &
                """output_filename"":{""type"":""string"",""description"":""Filename for the output .docx (default: derived from PDF name)""}" &
                "},""required"":[""attachment_name""]}}"
        })

        ' ── create_word_document ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CreateWordDoc,
            .CapabilityTags = "explicit_operation",
            .ModelDescription = "Create Executive Word Document (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CreateWordDoc & ": Creates a NEW polished Word document (.docx) from Markdown. " &
                "Do NOT use this tool merely to translate, correct, proofread, replace text in, or otherwise transform an existing Word attachment when the user wants a version of that existing document with its structure/formatting retained; use process_word_document for that case. " &
                "Use create_word_document when the requested deliverable is genuinely a newly authored document assembled from Markdown/content rather than a transformed copy of an existing Office document. " &
                "Default to style_preset='consulting' and professional_layout=true unless the user explicitly requests a plain document. " &
                "Write lean, decision-oriented content with strong section hierarchy, short paragraphs, informative headings, concise bullets, and tables where comparison is useful. " &
                "The renderer applies consulting-style typography, page margins, heading hierarchy, restrained accent color, professional table formatting, optional cover page, headers/footers, and page numbers. " &
                "For native Word cross-references, put [[anchor:ID]] on its own Markdown line immediately before the target heading or paragraph. Refer to it inline with [[ref:ID:number]] for the target's native full-context Word numbering (for example I.A or Randziffer 17), [[ref:ID:text]] for the target text, or [[ref:ID:full]] for native number plus target text. IDs use letters, digits, underscore, dot, or hyphen. Never type or guess the target heading/Randziffer number manually. The host converts these markers to real Word bookmarks + REF fields. If at least one REF marker is present, the host automatically performs one isolated hidden Word field-refresh pass after the DOCX is otherwise complete, saves the resolved field caches, and removes automatic-update-on-open state; no separate refresh argument is needed. " &
                "Use include_cover=true plus cover_title/cover_subtitle for reports, proposals, board papers, investment memos, and other substantial documents. " &
                "Use page_orientation='landscape' for wide tables. Use the DESIGN REPOSITORY routing metadata for Word format selection: explicit requested design first, then organization/document type/language, then an applicable default. If no design_name is supplied and the catalog has a global Word default, the host applies it deterministically. If that design lists Word slots, markdown_content is the substantive body and template_fields must contain the listed keys exactly. The DESIGN REPOSITORY summary includes each slot's declared purpose; use that purpose rather than guessing from a placeholder id. Visible placeholders such as [[RI:Text]] or [[RI:Body]] are labels only and may be renamed without changing their meaning. Shared style_policy_file mappings and design-local body-style mappings are authoritative: use only the declared Markdown heading and list levels, and let the host apply exact native Word styles. Structured DOCX is the recommended carrier for new Word designs; slot-bound DOCX designs and generic Word creation are OOXML-first and normally do not start Word. The sole automatic exception is one final hidden Word field-refresh pass when native [[ref:...]] cross-references are present. Only an explicitly selected legacy non-slot Word carrier may otherwise use the Word/COM compatibility renderer. Their native structure, fields, styles, headers and footers remain authoritative. Do not put front-matter fields into markdown_content when the design exposes dedicated template_fields. Explicit tool arguments override repository defaults. Use accent_color only when supplied by the user or a concrete authorized design source. Never claim a named organization's design unless such a source is available; otherwise use a neutral professional preset and state that the specific corporate design was not available. " &
                "For true Word footnotes, use the footnotes array and put exactly one inline [[footnote:ID]] marker at the intended reference position in markdown_content for each footnote. The host replaces these markers deterministically with native Word footnote references after document creation. Place the marker AFTER sentence/clause punctuation when the footnote supports the whole preceding sentence or clause (for example: '... Aussage.[[footnote:fn1]]'). Place it immediately after a specific word or phrase, before following punctuation, only when the footnote refers specifically to that word or phrase rather than to the whole sentence or clause. Do not manually number footnotes and do not simulate them with superscript text or a notes section. " &
                "NEVER simulate diagrams, process maps, timelines, or charts with ASCII/Unicode box characters, Mermaid source, fenced code blocks, monospace/Courier pseudo-graphics, or manual spacing. " &
                "Whenever the user requests, names, or the document itself labels something as a chart, diagram, organigram/organization chart, graphical representation, visualization, timeline, process map, or other graphic, the native 'visuals' parameter is REQUIRED and each requested graphic must have an exact [[visual:ID]] token on its own line in markdown_content. If the requested report calls for multiple graphics, EVERY requested graphic must have its own visuals entry and placeholder. Never omit a requested visual on a retry, and never substitute one with a table, repeated block/square characters, fenced character diagram, or arrow chain. Use type='process' for flows/value chains/workflows, type='timeline' for timelines, type='org_chart' for reporting hierarchies, bar_chart/column_chart/line_chart/area_chart/pie_chart/doughnut_chart for quantitative charts, and editable diagram types list/cycle/relationship/matrix/pyramid/hierarchy where they fit. Use type='smartart' only as a generic compatibility alias when a more specific diagram type does not fit. Every visual is EDITABLE by default. The host writes native Office Open XML directly into the saved DOCX: diagram families become editable DrawingML canvases/shapes and quantitative graphics become native Office chart parts with embedded XLSX data. No Word/Excel Interop is used to create visuals. Set insertion_mode='inline' or 'floating' only when placement is specifically required; otherwise use 'auto' (inline by default). Raster output is not used for create_word_document visuals. The Word-document visual renderer never substitutes PNG/raster output. For an organizational hierarchy use structured nodes and parent_id relationships and keep node labels concise, with names/subtitles in detail. " &
                "The renderer verifies after saving that every requested chart or DrawingML diagram was actually persisted inside the DOCX. It must fail rather than claim success if an editable visual disappeared. The resulting file is attached to the reply.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CreateWordDoc & """," &
                """description"":""Creates a NEW polished Word document (.docx) from Markdown. Not for translating/correcting/revising an existing Word attachment when source structure and formatting should be retained; use process_word_document for that transformation.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """markdown_content"":{""type"":""string"",""description"":""Full document content in Markdown. Use headings, concise bullets, emphasis, and Markdown tables where useful. For native cross-references, put [[anchor:ID]] alone on the line immediately before the target heading/paragraph and use inline [[ref:ID:number]], [[ref:ID:text]], or [[ref:ID:full]] markers; do not manually type the target's native number. When the selected Word design exposes a native body-style contract, use only the heading and list levels declared by that contract; unsupported levels are rejected.""}," &
                """file_name"":{""type"":""string"",""description"":""Desired output filename without .docx. Defaults to Document.""}," &
                """design_name"":{""type"":""string"",""description"":""Exact design id/name from the configured AgentResources design repository. Use this when the user explicitly named a particular design. Otherwise prefer document_type plus document_language so the host can enforce document-type-first routing.""}," &
                """document_type"":{""type"":""string"",""description"":""Requested Word artifact type such as memo, letter, note, or generic. When the user clearly names a document type, pass it. The host treats this as the primary routing key before language or global defaults.""}," &
                """document_language"":{""type"":""string"",""description"":""Requested document language, preferably a short language code such as de or en. Used only after document_type to choose among same-type design variants.""}," &
                """organization"":{""type"":""string"",""description"":""Optional organization/design-family routing hint when more than one organization is visible in the active repository. Do not invent one.""}," &
                """template_fields"":{""type"":""object"",""description"":""Values for deterministic Word-template slots exposed by the selected design. Use only the exact keys listed for that design in the DESIGN REPOSITORY context and follow each key's declared purpose. The companion .md maps these keys to arbitrary speaking placeholders such as [[RI:To]] or [[RI:Recipient]]; the placeholder id itself is not semantic."",""additionalProperties"":{""type"":""string""}}," &
                """document_title"":{""type"":""string"",""description"":""Document title metadata.""}," &
                """document_author"":{""type"":""string"",""description"":""Optional author metadata.""}," &
                """base_font_name"":{""type"":""string"",""description"":""Base font. Default Aptos.""}," &
                """base_font_size"":{""type"":""number"",""description"":""Base body size in points. Default 10.5.""}," &
                """page_orientation"":{""type"":""string"",""enum"":[""portrait"",""landscape""],""description"":""Page orientation.""}," &
                """professional_layout"":{""type"":""boolean"",""description"":""Default true. Applies executive typography, spacing, margins and professional tables.""}," &
                """style_preset"":{""type"":""string"",""enum"":[""consulting"",""executive"",""minimal"",""plain""],""description"":""Visual preset. Default consulting.""}," &
                """accent_color"":{""type"":""string"",""description"":""Primary accent as hex RGB, e.g. #17365D.""}," &
                """include_cover"":{""type"":""boolean"",""description"":""Add an executive cover page. Default false unless a substantive report/memo benefits from one.""}," &
                """cover_title"":{""type"":""string"",""description"":""Cover page title. If omitted and include_cover=true, document_title is used.""}," &
                """cover_subtitle"":{""type"":""string"",""description"":""Optional cover subtitle.""}," &
                """cover_kicker"":{""type"":""string"",""description"":""Optional small uppercase cover label such as CONFIDENTIAL or STRATEGY REVIEW.""}," &
                """header_text"":{""type"":""string"",""description"":""Optional running header text.""}," &
                """footer_text"":{""type"":""string"",""description"":""Optional running footer text.""}," &
                """show_page_numbers"":{""type"":""boolean"",""description"":""Show page numbers in the footer. Default true.""}," &
                """footnotes"":{""type"":""array"",""description"":""Native Word footnotes. For each entry, put exactly one inline [[footnote:ID]] marker at the intended reference position in markdown_content. If the note supports the whole preceding sentence or clause, put the marker after its punctuation; put it directly after a word/phrase before punctuation only when the note specifically refers to that word/phrase. IDs are caller-defined opaque labels used only to bind the marker to its footnote; the host assigns Word footnote numbers automatically."",""items"":{""type"":""object"",""properties"":{""id"":{""type"":""string"",""description"":""Stable marker id using letters, digits, underscore, dot, or hyphen.""},""text"":{""type"":""string"",""description"":""Footnote text. Plain text with optional line breaks; do not include a manual footnote number.""}},""required"":[""id"",""text""]}}," &
                """visuals"":{ ""type"":""array"",""description"":""Native embedded graphics. This array is REQUIRED whenever the user requests or the document labels any chart/diagram/organigram/visualization. Put exactly one [[visual:ID]] placeholder in markdown_content for each visual. Every requested graphic must be represented here on every retry; never replace or drop one in favor of a table, repeated block/square characters, fenced character diagram, bracket/arrow chain, or monospace text."",""items"":{ ""type"":""object"",""properties"":{" &
                """id"":{ ""type"":""string"",""description"":""Stable placeholder id used by [[visual:ID]].""}," &
                """type"":{ ""type"":""string"",""enum"":[""process"",""timeline"",""org_chart"",""hierarchy"",""list"",""cycle"",""relationship"",""matrix"",""pyramid"",""smartart"",""bar_chart"",""column_chart"",""line_chart"",""area_chart"",""pie_chart"",""doughnut_chart""],""description"":""Graphic type. Diagram families are written as native editable DrawingML canvases/shapes; quantitative graphics are native Office charts with embedded workbook data.""}," &
                """editable"":{ ""type"":""boolean"",""description"":""Defaults to true. create_word_document visuals are always emitted as native editable OOXML objects; raster substitution is not used.""}," &
                """insertion_mode"":{ ""type"":""string"",""enum"":[""auto"",""inline"",""floating""],""description"":""Native OOXML placement mode. Default auto writes inline; floating writes a top/bottom anchored object.""}," &
                """smartart_layout"":{ ""type"":""string"",""description"":""Compatibility hint retained for callers. The OOXML renderer uses deterministic editable DrawingML geometry and does not depend on an installed SmartArt layout.""}," &
                """title"":{ ""type"":""string""}," &
                """caption"":{ ""type"":""string""}," &
                """items"":{ ""type"":""array"",""description"":""For process/timeline/list/cycle/relationship/matrix/pyramid/smartart: strings or objects with label and optional detail."",""items"":{}}," &
                """nodes"":{ ""type"":""array"",""description"":""For org_chart/hierarchy: hierarchy nodes. Each node has a stable id, label, optional detail, and optional parent_id. Root nodes omit parent_id. Keep labels concise; use detail for role/name."",""items"":{ ""type"":""object"",""properties"":{ ""id"":{ ""type"":""string""},""label"":{ ""type"":""string""},""detail"":{ ""type"":""string""},""parent_id"":{ ""type"":""string""}},""required"":[""id"",""label""]}}," &
                """categories"":{ ""type"":""array"",""items"":{ ""type"":""string""}}," &
                """series"":{ ""type"":""array"",""description"":""For charts: [{name, values:[numbers]}]."",""items"":{ ""type"":""object"",""properties"":{ ""name"":{ ""type"":""string""},""values"":{ ""type"":""array"",""items"":{ ""type"":""number""}}},""required"":[""values""]}}," &
                """width_inches"":{ ""type"":""number"",""description"":""Optional display width. The native renderer constrains it to the printable page area and uses visual-specific defaults.""}," &
                """height_inches"":{ ""type"":""number"",""description"":""Optional display height. The native renderer constrains it to the printable page area and uses visual-specific defaults.""}},""required"":[""id"",""type""]}}," &
                """table_style_name"":{""type"":""string"",""description"":""Optional Word table style name. Renderer still applies professional header/banding treatment.""}" &
                "},""required"":[""markdown_content""]}}"
        })

        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CompleteWordTables,
            .ModelDescription = "Complete Word Tables (built-in)",
            .ToolPriority = 950,
            .ToolInstructionsPrompt =
                AP_Tool_CompleteWordTables & ": Use this to fill in an existing Word form or table-based document in place. " &
                "Prefer this tool over process_word_document when the goal is to complete empty fields, content controls, checkboxes, dropdowns, or table cells. " &
                "Do not use process_word_document for form filling when this tool applies.",
                        .ToolDefinition =
                "{""name"":""" & AP_Tool_CompleteWordTables & """," &
                """description"":""Completes empty or incomplete tables, body placeholders, and form fields in Word documents " &
                "(.docx or .doc) using AI while preserving the original structure and formatting as much as possible. " &
                "Produces a completed .docx and optionally a compare document.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """instruction"":{""type"":""string"",""description"":""Instructions describing how the tables and placeholders should be completed.""}," &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Optional list of Word files to process. If omitted, all .docx/.doc files in the current session are processed.""}," &
                """use_second_api"":{""type"":""boolean"",""description"":""Optional. Use the secondary API/model for the completion run.""}," &
                """create_compare_document"":{""type"":""boolean"",""description"":""Optional. Default true. If true, also creates a compare document with tracked changes.""}" &
                "},""required"":[""instruction""]}}"
        })

        ' ── create_excel_spreadsheet ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CreateExcel,
            .CapabilityTags = "explicit_operation",
            .ModelDescription = "Create Professional Excel Workbook (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CreateExcel & ": Creates a product-quality Excel workbook (.xlsx/.xlsm). Use this for tables, trackers, budgets, models, dashboards, analyses, schedules, and structured datasets. " &
                "Default to professional_layout=true, style_preset='consulting', smart_format=true. The renderer itself supplies clean typography, dark executive headers, subtle banding, sensible column sizing, hidden gridlines, frozen headers, filters, and semantic Status/Priority highlighting when applicable. " &
                "Do NOT waste tokens manually styling every cell unless a special exception is needed. Focus the tool call on workbook structure, data, formulas and useful interactive features. " &
                "For non-trivial workbooks, prefer multiple purpose-built worksheets (e.g. Executive Summary/Dashboard, Data, Assumptions, Calculations, Instructions). " &
                "Use native tables for datasets, charts for decision-relevant trends/comparisons, data_validations for controlled inputs/dropdowns, conditional_formats for thresholds/alerts, named_ranges for model clarity, and vba_modules only when automation is requested. " &
                "Keep dashboards lean: a few KPIs, 2-4 useful charts, clear labels, no decorative clutter. Use English formula syntax with comma separators. If the user requests a named design listed by the DESIGN REPOSITORY system context, pass its exact id as design_name; the tool applies configured design defaults and may use a clean repository workbook template as a theme carrier. Never claim a named organization's Excel design unless concrete design specifications or an authorized workbook/template source are available; otherwise use the neutral professional preset and say the specific corporate design was not available. The file is attached to the reply.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CreateExcel & """," &
                """description"":""Creates a professional multi-sheet Excel workbook (.xlsx/.xlsm) with formulas, native tables, charts, dropdowns/data validation, conditional formatting, named ranges, optional VBA automation, print setup and opinionated consulting-style formatting.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """cells"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """cell"":{""type"":""string"",""description"":""Cell/range address in A1 notation.""}," &
                """value"":{""description"":""Cell value (string or number).""}," &
                """formula"":{""type"":""string"",""description"":""Excel formula starting with =.""}," &
                """bold"":{""type"":""boolean""},""italic"":{""type"":""boolean""},""underline"":{""type"":""boolean""},""strikethrough"":{""type"":""boolean""}," &
                """font_name"":{""type"":""string""},""font_size"":{""type"":""number""},""font_color"":{""type"":""string""},""bg_color"":{""type"":""string""}," &
                """number_format"":{""type"":""string"",""description"":""Excel number format such as #,##0.00, 0.0%, yyyy-mm-dd.""}," &
                """h_align"":{""type"":""string"",""enum"":[""left"",""center"",""right""]},""v_align"":{""type"":""string"",""enum"":[""top"",""center"",""bottom""]}," &
                """wrap_text"":{""type"":""boolean""},""border"":{""type"":""string""},""border_color"":{""type"":""string""}," &
                """text_rotation"":{""type"":""integer""},""indent"":{""type"":""integer""},""comment"":{""type"":""string""}," &
                """hyperlink"":{""type"":""string""},""hyperlink_display"":{""type"":""string""}" &
                "}},""description"":""Cells for default/first sheet. Low-level formatting is optional because professional_layout provides a strong baseline.""}," &
                """sheets"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{""name"":{""type"":""string""},""cells"":{""type"":""array"",""items"":{""type"":""object""}}}},""description"":""Multiple sheets. Each sheet may also override column_widths, row_heights, auto_fit_columns, auto_fit_rows, merge_ranges, freeze_pane, auto_filter, data_validations, conditional_formats, print_setup, tab_color, show_gridlines, zoom, right_to_left, professional_layout, style_preset, accent_color, font_name, smart_format, header_row.""}," &
                """file_name"":{""type"":""string"",""description"":""Filename without extension.""}," &
                """design_name"":{""type"":""string"",""description"":""Exact design id/name from the configured AgentResources design repository. Explicit creator arguments override design defaults.""}," &
                """sheet_name"":{""type"":""string"",""description"":""Tab name for single-sheet mode.""}," &
                """professional_layout"":{""type"":""boolean"",""description"":""Default true. Enables opinionated professional workbook styling.""}," &
                """style_preset"":{""type"":""string"",""enum"":[""consulting"",""executive"",""minimal"",""plain""],""description"":""Default consulting.""}," &
                """accent_color"":{""type"":""string"",""description"":""Primary workbook accent as hex RGB. Default #17365D.""}," &
                """font_name"":{""type"":""string"",""description"":""Workbook font. Default Aptos.""}," &
                """smart_format"":{""type"":""boolean"",""description"":""Default true. Adds semantic formatting for common Status/Priority columns.""}," &
                """header_row"":{""type"":""integer"",""description"":""Header row for automatic styling/filter/freeze behavior. Default is the first used row. Set 0 for dashboard sheets without a tabular header.""}," &
                """column_widths"":{""type"":""object"",""description"":""Optional explicit {column_letter: width}. If omitted, professional mode auto-fits and caps widths.""}," &
                """row_heights"":{""type"":""object"",""description"":""Optional explicit row heights.""}," &
                """auto_fit_columns"":{""description"":""true/'all', a column letter, or an array of letters.""}," &
                """auto_fit_rows"":{""description"":""true/'all', a row number, or an array of row numbers.""}," &
                """tab_color"":{""type"":""string""},""show_gridlines"":{""type"":""boolean""},""zoom"":{""type"":""integer""},""right_to_left"":{""type"":""boolean""}," &
                """merge_ranges"":{""type"":""array"",""items"":{""type"":""string""}},""freeze_pane"":{""type"":""string""},""auto_filter"":{""type"":""string""}," &
                """data_validations"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """range"":{""type"":""string""},""type"":{""type"":""string""},""formula1"":{""type"":""string""},""formula2"":{""type"":""string""},""operator"":{""type"":""string""}," &
                """show_dropdown"":{""type"":""boolean""},""input_title"":{""type"":""string""},""input_message"":{""type"":""string""},""error_title"":{""type"":""string""},""error_message"":{""type"":""string""}" &
                "}},""description"":""Data validation rules, including list dropdowns.""}," &
                """conditional_formats"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """range"":{""type"":""string""},""type"":{""type"":""string"",""description"":""cell_value, text_contains, duplicate, unique, color_scale, data_bar, icon_set, top_10""},""operator"":{""type"":""string""}," &
                """formula1"":{""type"":""string""},""formula2"":{""type"":""string""},""format_font_color"":{""type"":""string""},""format_bg_color"":{""type"":""string""},""format_bold"":{""type"":""boolean""}" &
                "}},""description"":""Conditional formatting rules.""}," &
                """tables"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """sheet_name"":{""type"":""string""},""range"":{""type"":""string"",""description"":""Data range including headers, e.g. A1:H50.""},""name"":{""type"":""string""},""style"":{""type"":""string"",""description"":""Native Excel table style, default TableStyleMedium2.""}," &
                """show_totals"":{""type"":""boolean""},""show_row_stripes"":{""type"":""boolean""},""show_first_column"":{""type"":""boolean""},""show_last_column"":{""type"":""boolean""}" &
                "}},""description"":""Native Excel ListObjects for structured datasets. Prefer these for tables users will filter/sort/extend.""}," &
                """charts"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """type"":{""type"":""string"",""enum"":[""column"",""bar"",""line"",""pie"",""area"",""scatter"",""doughnut""]},""data_range"":{""type"":""string""},""title"":{""type"":""string""},""position"":{""type"":""string""}," &
                """width"":{""type"":""number""},""height"":{""type"":""number""},""sheet_name"":{""type"":""string""},""color"":{""type"":""string""},""series_colors"":{""type"":""array"",""items"":{""type"":""string""}}," &
                """show_legend"":{""type"":""boolean""},""legend_position"":{""type"":""string"",""enum"":[""top"",""bottom"",""left"",""right"",""corner""]},""show_data_labels"":{""type"":""boolean""},""x_axis_title"":{""type"":""string""},""y_axis_title"":{""type"":""string""}" &
                "}},""description"":""Charts. The renderer applies a clean executive chart style and a restrained default palette.""}," &
                """named_ranges"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{""name"":{""type"":""string""},""range"":{""type"":""string""}}}}," &
                """vba_modules"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{""name"":{""type"":""string""},""code"":{""type"":""string""},""type"":{""type"":""string""}}},""description"":""Optional VBA modules; presence saves as .xlsm. Use when the user requests workbook automation. Requires Excel Trust Center permission for programmatic VBProject access.""}," &
                """print_setup"":{""type"":""object"",""properties"":{""orientation"":{""type"":""string"",""enum"":[""portrait"",""landscape""]},""fit_to_pages_wide"":{""type"":""integer""},""fit_to_pages_tall"":{""type"":""integer""},""header_text"":{""type"":""string""},""footer_text"":{""type"":""string""}}}" &
                "},""required"":[]}}"
        })

        ' ── create_powerpoint ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CreatePowerPoint,
            .CapabilityTags = "explicit_operation",
            .ModelDescription = "Create Executive PowerPoint (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CreatePowerPoint & ": Creates a polished executive PowerPoint deck. Keep one clear message per slide, use concise action-oriented titles, and keep narrative content short enough for presentation-sized typography. " &
                "Supported layouts: title, section, bullets, two_column, kpi, table, chart, cards, process, structure, timeline, comparison, matrix, quote, closing. " &
                "PRIMARY AUTHORING CONTRACT: use slide.visual for every intended rich graphic and slide.bullet_items for structured or nested bullets. The visual object is typed and editable-first; the host derives the semantic layout from visual.type and renders native PowerPoint shapes/text whenever possible. data_json and flat bullet_points remain compatibility inputs only and should not be the primary format for new calls. " &
                "For bullet_items, never embed bullet glyphs or manual numbering such as '1.'/'2.' in text. Use level=0 for a main point and level=1 (or deeper) for supporting detail; the host creates the actual PowerPoint bullet and indentation. " &
                "For visual, keep label/title text short and put only concise supporting wording in detail/body. Use type='process' for operational workflows/actions, type='structure' for ownership/reporting/entity relationships, type='comparison' for alternatives, type='cards' for parallel themes, type='timeline' for milestone/phase roadmaps (with dates when available, but dates are not required), type='matrix' for two-dimensional positioning, type='kpi' for headline metrics, and type='chart' only for genuine quantitative data. A roadmap whose steps represent phases/milestones should normally be a timeline rather than a process. " &
                "Choose visuals by meaning: quantitative series/trends/comparisons -> chart; process/sequence -> process; corporate or ownership hierarchy -> structure; qualitative themes/benefits/risks -> cards; alternatives -> comparison; dated milestones -> timeline; two-dimensional strategic positioning -> matrix. " &
                "Do NOT invent numbers merely to create a chart. When real quantitative data exists, prefer a chart over repeating the numbers as prose. VISUAL BALANCE DEFAULT: an executive deck needs both native text/data slides and selected rich visuals. For decks with 6+ slides, normally keep rich graphics to roughly 30-60% of content slides and use native bullets, two_column, table or quote layouts for explanatory, legal, tax, caveat and detail-heavy material. Do not turn every slide into cards/process boxes merely to satisfy a visual quota, and normally use no more than two rich visual slides consecutively. Treat structure, process, comparison, cards, timeline, matrix and KPI as vector-infographic/illustration layouts only when that visual form genuinely improves comprehension. Use structure diagrams for ownership/entity relationships, process graphics for true sequences, comparison columns for alternatives, cards for a small set of parallel themes, and timeline/matrix where semantically appropriate. Do not include literal bullet glyphs such as •, -, *, or + at the start of body lines; provide plain newline-separated text because the renderer supplies bullet formatting. " &
                "Rich graphics have strict presentation-sized text budgets: keep visual labels short, keep detail/body to one concise sentence or short phrase, keep comparison items compact, and move dense legal/tax explanation to native text slides. The host refuses a native rich visual that cannot fit at readable diagram typography; it may fall back to native text rather than silently shrink content. " &
                "Preferred visual shapes: cards={type:'cards',items:[{label,detail,badge,tone}]}; process={type:'process',items:[{label,detail}]}; structure={type:'structure',nodes:[{id,label,detail,parent_id}]}; timeline={type:'timeline',items:[{label,title,detail}]}; comparison={type:'comparison',columns:[{title,items,verdict,tone}]}; matrix={type:'matrix',quadrants:[{title,body}],x_left,x_right,y_top,y_bottom}; kpi={type:'kpi',items:[{label,value,detail}]}; chart={type:'chart',chart_type,categories,series:[{name,values}]}. " &
                "Legacy data_json is accepted for backwards compatibility but new calls should use visual directly. " &
                "Keep tables concise: preferably <= 6 data rows per slide; split long tables across slides rather than forcing tiny text. If the user requests a named design listed by the DESIGN REPOSITORY system context, pass its exact id as design_name; the repository template is treated as a real PowerPoint master/layout carrier and its sample slides are removed by default. " &
                "If a requested .potx/.pptx template is listed under USER HOME FILES, load manage_user_files with action='use' before create_powerpoint and then pass the loaded filename as template_attachment_name. template_attachment_name may reference an existing .potx or .pptx. " &
                "When a template is available, the host selects among its actual masters/custom layouts per slide. Repository carriers may have a same-basename Markdown guide; when present, that natural-language guide is used together with the actual master/layout names, placeholder geometry/types, and sample slides. Template-specific rich-visual proportions may also be declared there with generic rich.* settings; do not invent brand-specific coordinates in the tool call. Rich timelines, charts, matrices, processes, structures, cards, comparisons and KPI visuals are freeform editable graphics: they may span the safe content canvas and intentionally leave native body placeholders empty, while the host protects title/footer/logo/slide-number zones. Without guidance, the host may use one bounded LLM interpretation call over that actual template metadata. Do not assume numeric layout positions or fixed organization-specific layout-name heuristics. template_master_name/template_layout_name are optional exact overrides only when their names are known; otherwise let the host interpret the template and fall back safely if no layout fits. " &
                "A named organization's design may be claimed only when design_name resolves, template_attachment_name resolves to an authorized template, or explicit design specifications are supplied; otherwise create a neutral professional deck and state that the specific corporate design was not available.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CreatePowerPoint & """," &
                """description"":""Creates a professionally formatted PowerPoint presentation with executive typography, visual business layouts, charts, and optional template support.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """slides"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """layout"":{""type"":""string"",""description"":""Optional layout: title, section, bullets, two_column, kpi, table, chart, cards, process, structure, timeline, comparison, matrix, quote, or closing.""}," &
                """title"":{""type"":""string""},""subtitle"":{""type"":""string""},""body"":{""type"":""string"",""description"":""Main text; use plain newline-separated points for bullet slides without literal bullet glyphs and keep it concise enough for large presentation text.""}," &
                """bullet_items"":{""type"":""array"",""description"":""Preferred structured bullets. Use level=0 for main points and level=1+ for supporting detail. Never put bullet glyphs or manual numeric prefixes such as 1. or 2. in text."",""items"":{""type"":""object"",""properties"":{""text"":{""type"":""string""},""level"":{""type"":""integer"",""description"":""Zero-based indentation level, normally 0 or 1.""}}, ""required"":[""text""]}}," &
                """bullet_points"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Legacy flat bullet list. Prefer bullet_items for new calls, especially for hierarchy. Do not prefix entries with bullet glyphs or manual numbering.""}," &
                """visual"":{""type"":""object"",""description"":""Preferred typed rich visual. Native editable PowerPoint shapes/text are the primary renderer. Keep labels concise; dense legal/tax explanation belongs on native text slides."",""properties"":{""type"":{""type"":""string"",""enum"":[""cards"",""process"",""structure"",""timeline"",""comparison"",""matrix"",""kpi"",""chart""]},""title"":{""type"":""string""},""caption"":{""type"":""string""},""items"":{""type"":""array"",""description"":""Flat array of item objects. Never wrap item objects in additional arrays."",""items"":{""type"":""object"",""properties"":{""label"":{""type"":""string""},""title"":{""type"":""string""},""detail"":{""type"":""string""},""body"":{""type"":""string""},""badge"":{""type"":""string""},""tone"":{""type"":""string""},""value"":{""type"":""string""}}}},""nodes"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{""id"":{""type"":""string""},""label"":{""type"":""string""},""detail"":{""type"":""string""},""parent_id"":{""type"":""string""}},""required"":[""id"",""label""]}},""columns"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{""title"":{""type"":""string""},""items"":{""type"":""array"",""items"":{""type"":""string""}},""verdict"":{""type"":""string""},""tone"":{""type"":""string""}}}},""quadrants"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{""title"":{""type"":""string""},""body"":{""type"":""string""}}}},""x_left"":{""type"":""string""},""x_right"":{""type"":""string""},""y_top"":{""type"":""string""},""y_bottom"":{""type"":""string""},""chart_type"":{""type"":""string""},""categories"":{""type"":""array"",""items"":{""type"":""string""}},""series"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{""name"":{""type"":""string""},""values"":{""type"":""array"",""items"":{""type"":""number""}}}}}},""required"":[""type""]}," &
                """notes"":{""type"":""string""},""source"":{""type"":""string""},""callout"":{""type"":""string""}," &
                """section_number"":{""type"":""string""},""left_title"":{""type"":""string""},""left_body"":{""type"":""string""},""right_title"":{""type"":""string""},""right_body"":{""type"":""string""}," &
                """template_master_name"":{""type"":""string"",""description"":""Optional exact PowerPoint master/design name override when known from the approved template. Do not guess.""}," &
                """template_layout_name"":{""type"":""string"",""description"":""Optional exact custom-layout name override when known from the approved template. Do not guess; otherwise omit and let the host choose.""}," &
                """quote"":{""type"":""string""},""attribution"":{""type"":""string""}," &
                """data_json"":{""type"":""string"",""description"":""Legacy compatibility input for structured visual data. Prefer the typed visual object for new calls.""}" &
                "}},""description"":""Slides defining the presentation.""}," &
                """file_name"":{""type"":""string""},""title"":{""type"":""string""},""design_name"":{""type"":""string"",""description"":""Exact design id/name from the configured AgentResources design repository. Repository templates are used as real master/theme/layout carriers and sample slides are removed by default.""},""template_attachment_name"":{""type"":""string"",""description"":""Exact loaded/attached .potx or .pptx template filename. For a USER HOME FILE, call manage_user_files action='use' first.""}," &
                """preserve_template_slides"":{""type"":""boolean"",""description"":""Keep existing slides from a .pptx template. Defaults to false for design-repository templates; .potx templates normally contain no sample slides.""}," &
                """allow_text_heavy"":{""type"":""boolean"",""description"":""Set true only when the user explicitly requests a mostly textual/minimal deck. Otherwise the host enforces a balanced text/visual mix for longer presentations.""}," &
                """allow_visual_heavy"":{""type"":""boolean"",""description"":""Set true only when the user explicitly requests an infographic-heavy/mostly visual deck. Otherwise the host preserves a meaningful share of native text/data slides.""}," &
                """style_preset"":{""type"":""string""},""accent_color"":{""type"":""string""},""secondary_color"":{""type"":""string""},""font_name"":{""type"":""string""}," &
                """text_color"":{""type"":""string""},""muted_color"":{""type"":""string""},""light_color"":{""type"":""string""},""line_color"":{""type"":""string""}," &
                """aspect_ratio"":{""type"":""string"",""description"":""16:9 or 4:3.""},""footer_text"":{""type"":""string""},""show_slide_numbers"":{""type"":""boolean""}" &
                "},""required"":[""slides""]}}"
        })

        ' ── create_code_file ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CreateCodeFile,
            .CapabilityTags = "explicit_operation",
            .ModelDescription = "Create Code/Script File (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CreateCodeFile & ": Creates a new code, script, or data file with the specified content. " &
                "Use this when the user asks you to create, generate, write, or produce any code file, script, " &
                "configuration file, or structured data file. Examples: HTML pages, Python scripts, JavaScript files, " &
                "JSON/YAML/XML data files, CSS stylesheets, SQL scripts, shell scripts (.sh/.bat/.ps1), " &
                "Markdown documents, CSV files, INI/TOML/ENV config files, Dockerfiles, etc. " &
                "You MUST determine the appropriate file extension based on the content and language. " &
                "You MUST provide the complete, functional, ready-to-execute file content — do NOT use placeholders " &
                "or incomplete code. The resulting file will be attached to the reply email so the user can save and run it.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CreateCodeFile & """," &
                """description"":""Creates a new code, script, or data file with the specified content and attaches it to the reply. " &
                "Supports any text-based file format: HTML, Python, JavaScript, TypeScript, JSON, YAML, XML, CSS, SQL, " &
                "shell scripts, batch files, PowerShell, Markdown, CSV, INI, TOML, Dockerfiles, and more. " &
                "Determine the correct filename and extension based on the content and user request.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """file_name"":{""type"":""string"",""description"":""Full filename including extension (e.g. 'index.html', 'analysis.py', 'config.json', 'setup.sh'). " &
                "Choose a descriptive name and the correct extension for the language/format.""}," &
                """content"":{""type"":""string"",""description"":""The complete file content. Must be functional and ready to use — no placeholders or TODOs.""}," &
                """description"":{""type"":""string"",""description"":""Optional brief description of what the file does, shown to the user in the response.""}" &
                "},""required"":[""file_name"",""content""]}}"
        })

        ' ── comment_pdf_document ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CommentPdf,
            .ModelDescription = "Comment PDF Document (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CommentPdf & ": Adds review comments as highlight annotations with popups to a PDF attachment. " &
                "Use this ONLY when the user explicitly asks to ADD, INSERT, or PLACE comments, annotations, " &
                "review notes, or feedback INSIDE a PDF file — i.e. the user wants the PDF itself modified with " &
                "embedded annotation bubbles. " &
                "Do NOT use this tool when the user asks to READ, EXTRACT, SUMMARIZE, or UNDERSTAND existing " &
                "comments or content from a PDF — use extract_pdf_text or read_attachment for that instead. " &
                "Do NOT use this tool when the user wants a textual summary or analysis of a PDF — only when " &
                "annotations should appear as highlight + popup comment pairs within the PDF itself. " &
                "Supports an optional author parameter: if the user asks for comments under a specific name " &
                "(e.g. the sender's name), pass it as author. If not specified, comments are authored as 'Inky'. " &
                "Comments that cannot be matched to specific text in the PDF are placed as sticky notes " &
                "at the top-right corner of the first page.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CommentPdf & """," &
                """description"":""Adds review comments as highlight annotations with popup bubbles directly inside a PDF file. " &
                "Use ONLY when the user wants to ADD or INSERT comments/annotations/review feedback INTO the PDF. " &
                "Do NOT use when the user wants to READ or EXTRACT existing content or comments from a PDF. " &
                "Matched text is highlighted in yellow with a popup comment; unmatched comments become sticky notes.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """instruction"":{""type"":""string"",""description"":""The review instruction (e.g., 'Review for legal risks', 'Check for inconsistencies', 'Suggest improvements')""}," &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of the PDF attachments to annotate. If empty or omitted, annotates all PDF attachments.""}," &
                """author"":{""type"":""string"",""description"":""Optional author name for the annotations. Use this when the user requests a specific name. If omitted, defaults to 'Inky'.""}" &
                "},""required"":[""instruction""]}}"
        })


        ' ── extract_data_from_attachments ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ExtractDataFromAttachments,
            .ModelDescription = "Extract structured data from attachments into a table (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ExtractDataFromAttachments & ": Extracts structured/tabular data from one or more attachments " &
                "(PDF, Word, Excel, text files, or files from a .zip archive) using AI-driven fact extraction. " &
                "You MUST provide an 'instruction' describing WHAT to extract (e.g. 'Extract invoice number, date, vendor name, and total amount'). " &
                "You SHOULD provide a 'schema' defining the output columns using the format 'ColumnName:type;ColumnName:type' " &
                "where type is one of: text, date, datetime, number, other. Example: 'Invoice Number:text;Date:date;Vendor:text;Amount:number'. " &
                "If no schema is provided, the AI will infer one automatically. " &
                "The tool processes each file individually with the AI, then merges and returns the combined result as a JSON table. " &
                "After receiving the result, YOU decide the best output format based on the user's request: " &
                "- Use create_excel_spreadsheet to produce a formatted .xlsx file " &
                "- Use create_word_document to produce a formatted .docx report " &
                "- Include the data directly in your reply as a formatted text table " &
                "- Or any other appropriate presentation. " &
                "This tool ONLY extracts and returns the structured data — it does NOT create any files.",
                .ToolDefinition =
                "{""name"":""" & AP_Tool_ExtractDataFromAttachments & """," &
                """description"":""Extracts structured/tabular data from one or more attachments using AI-driven fact extraction. " &
                "Returns a JSON object with 'schema' (column definitions) and 'rows' (extracted data). " &
                "Supports PDF, Word, Excel, text files, and files unpacked from .zip archives. " &
                "You MUST then decide how to present the result (create_excel_spreadsheet, create_word_document, or inline table).""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """instruction"":{""type"":""string"",""description"":""Natural-language instruction describing what data to extract. " &
                "Be specific about the fields/facts to capture. Example: 'Extract party names, contract date, governing law, and termination clauses from each document.'""}," &
                """schema"":{""type"":""string"",""description"":""Optional but recommended: column definitions in 'Name:type;Name:type' format. " &
                "Types: text, date, datetime, number, other. Append * to mark the sort column. " &
                "Example: 'Invoice No:text;Date:date*;Vendor:text;Amount:number;Notes:text'. " &
                "If omitted, the AI infers the schema automatically.""}," &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of attachments to extract data from. " &
                "If empty or omitted, processes all readable attachments (PDF, DOCX, XLSX, TXT, CSV, etc.).""}," &
                """output_language"":{""type"":""string"",""description"":""Language for extracted column names and textual values (e.g. 'English', 'German', 'French'). " &
                "Use the language the user expects the output in. If omitted, uses the language of the user's email.""}," &
                """sort_column"":{""type"":""integer"",""description"":""Optional: 1-based column index to sort by. Use 0 or omit for no sorting.""}," &
                """sort_direction"":{""type"":""string"",""enum"":[""ASC"",""DESC""],""description"":""Sort direction (default: ASC)""}," &
                """date_columns"":{""type"":""string"",""description"":""Optional: comma-separated 1-based column indices that contain dates, for normalization. Example: '2,5'""}" &
                "},""required"":[""instruction""]}}"
        })

        ' ── redact_pdf ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_RedactPdf,
            .CapabilityTags = "explicit_operation",
            .ModelDescription = "Redact PDF Document (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_RedactPdf & ": Redacts a PDF document by identifying text that matches the given instruction " &
                "and placing redaction boxes over it. Uses AI to analyze the PDF text and determine what should be redacted. " &
                "Operates in three modes controlled by the 'mode' parameter: " &
                "(1) 'prepare' (default): Creates removable red annotation boxes over identified text. " &
                "The user can review and adjust these in a PDF viewer before finalizing. " &
                "(2) 'finalize': Takes a previously prepared PDF (with redaction annotation boxes) and burns " &
                "them into permanent black rectangles by rasterizing each page. No AI analysis is performed. " &
                "(3) 'prepare_and_finalize': Performs both steps in one call — identifies text, places boxes, " &
                "and immediately burns them in as permanent black redactions. " &
                "IMPORTANT: When mode is 'prepare' or 'prepare_and_finalize', an 'instruction' is REQUIRED " &
                "describing what to redact (e.g. 'Redact all personal names and addresses', " &
                "'Redact financial information', 'Redact everything except party names and dates'). " &
                "When mode is 'finalize', no instruction is needed — it just burns in existing annotations. " &
                "The 'include_reason_codes' parameter adds brief labels (e.g. 'name', 'address') to each " &
                "redaction box, which are visible as white text inside the black box after finalization.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_RedactPdf & """," &
                """description"":""Redacts a PDF by identifying text with AI and placing redaction boxes, " &
                "or finalizes existing redaction boxes into permanent black rectangles. " &
                "Modes: 'prepare' (removable red boxes), 'finalize' (burn in existing boxes), " &
                "'prepare_and_finalize' (identify + burn in one step). " &
                "Requires 'instruction' for prepare modes (e.g. 'Redact all personal data').""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the PDF attachment to redact""}," &
                """instruction"":{""type"":""string"",""description"":""What to redact — required for 'prepare' and 'prepare_and_finalize' modes. " &
                "Examples: 'Redact all personal names, addresses, and phone numbers', " &
                "'Redact financial data including account numbers and amounts', " &
                "'Redact everything except the contract parties and effective dates'""}," &
                """mode"":{""type"":""string"",""enum"":[""prepare"",""finalize"",""prepare_and_finalize""]," &
                """description"":""Operation mode. 'prepare' = AI-driven removable boxes (default). " &
                "'finalize' = burn in existing annotation boxes. " &
                "'prepare_and_finalize' = AI-driven + immediate burn-in.""}," &
                """include_reason_codes"":{""type"":""boolean"",""description"":""Include brief reason labels (e.g. 'name', 'address') in each redaction box. Default: false.""}," &
                """output_filename"":{""type"":""string"",""description"":""Filename for the output PDF (default: derived from input with '_redacted' or '_final' suffix)""}" &
                "},""required"":[""attachment_name""]}}"
        })

        ' ── overlay_pdf ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_OverlayPdf,
            .CapabilityTags = "explicit_operation",
            .ModelDescription = "Overlay text and images on PDF pages (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_OverlayPdf & ": Places text labels and/or images at precise positions on PDF pages. " &
                "Use this when the user wants to add a logo, stamp, header, footer, badge, label, signature image, " &
                "or any positioned content onto a PDF. Supports per-element page targeting (single page, page range, or all pages), " &
                "font family/size/style/color, rotation, opacity, and image scaling. " &
                "Coordinates use PDF points (1 pt = 1/72 inch). A4 page = 595 × 842 pt. Letter = 612 × 792 pt. " &
                "Origin (0,0) is the TOP-LEFT corner of the page. " &
                "For images, reference an existing attachment by name via 'image_attachment_name'. " &
                "Text elements and image elements can be freely mixed in the same call. " &
                "The tool draws elements in array order (later elements overlay earlier ones).",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_OverlayPdf & """," &
                """description"":""Places text labels and/or images at precise positions on PDF pages. " &
                "Coordinates are in PDF points (1/72 inch). Origin (0,0) = top-left. A4 = 595×842 pt, Letter = 612×792 pt. " &
                "Elements are drawn in array order. Use for logos, stamps, headers, footers, labels, signatures, badges, etc.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the PDF attachment to overlay onto""}," &
                """elements"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """type"":{""type"":""string"",""enum"":[""text"",""image""],""description"":""Element type: 'text' for a text label, 'image' for an image file""}," &
                """pages"":{""type"":""string"",""description"":""Target pages: 'all' for every page, '1' for page 1 only, '1,3,5' for specific pages, '2-5' for a range. Default: 'all'""}," &
                """x"":{""type"":""number"",""description"":""X position in points from the left edge of the page""}," &
                """y"":{""type"":""number"",""description"":""Y position in points from the top edge of the page""}," &
                """text"":{""type"":""string"",""description"":""(text only) The text string to render. Supports \\n for line breaks.""}," &
                """font_family"":{""type"":""string"",""description"":""(text only) Font family name, e.g. 'Arial', 'Times New Roman', 'Calibri'. Default: 'Arial'""}," &
                """font_size"":{""type"":""number"",""description"":""(text only) Font size in points. Default: 12""}," &
                """bold"":{""type"":""boolean"",""description"":""(text only) Bold text. Default: false""}," &
                """italic"":{""type"":""boolean"",""description"":""(text only) Italic text. Default: false""}," &
                """font_color"":{""type"":""string"",""description"":""(text only) Hex RGB color, e.g. '#FF0000' for red, '#000000' for black. Default: '#000000'""}," &
                """h_align"":{""type"":""string"",""enum"":[""left"",""center"",""right""],""description"":""(text only) Horizontal alignment relative to x position. 'left' = x is left edge, 'center' = x is center point, 'right' = x is right edge. Default: 'left'""}," &
                """max_width"":{""type"":""number"",""description"":""(text only) Maximum width in points for text bounding box. Text is clipped or wrapped beyond this. Default: no limit""}," &
                """image_attachment_name"":{""type"":""string"",""description"":""(image only) Filename of the image attachment to place (PNG, JPG, BMP, GIF, TIFF, WEBP)""}," &
                """width"":{""type"":""number"",""description"":""(image only) Width in points to scale the image to""}," &
                """height"":{""type"":""number"",""description"":""(image only) Height in points to scale the image to""}," &
                """rotation"":{""type"":""number"",""description"":""Rotation angle in degrees (clockwise). Default: 0""}," &
                """opacity"":{""type"":""number"",""description"":""Opacity from 0.0 (fully transparent) to 1.0 (fully opaque). Default: 1.0""}" &
                "}}," &
                """description"":""Array of overlay elements (text and/or image) to place on the PDF""}," &
                """output_filename"":{""type"":""string"",""description"":""Filename for the output PDF (default: '<original>_overlay.pdf')""}" &
                "},""required"":[""attachment_name"",""elements""]}}"
        })

        ' ── create_audio_file ──
        AB_DetectTTSEngines()
        If AB_googleAvailable OrElse AB_openAIAvailable Then
            Dim engineHint As String = ""
            If AB_googleAvailable AndAlso AB_openAIAvailable Then
                engineHint = " Both Google and OpenAI TTS engines are available. " &
                    "For non-English text (German, French, etc.), prefer engine='google' for native-sounding pronunciation. " &
                    "For English text, prefer engine='openai' for natural-sounding voices. " &
                    "If the user explicitly requests an engine, honour their choice."
            ElseIf AB_googleAvailable Then
                engineHint = " Only Google TTS is available."
            Else
                engineHint = " Only OpenAI TTS is available."
            End If

            tools.Add(New ModelConfig() With {
                .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CreateAudioFile,
                .CapabilityTags = "explicit_operation",
                .ModelDescription = "Create Audio File — Podcast or Audiobook (built-in)",
                .ToolInstructionsPrompt =
                AP_Tool_CreateAudioFile & ": Generates an MP3 audio file from text content using text-to-speech. " &
                "Supports two modes: " &
                "(1) 'podcast' — converts the text into an engaging two-speaker dialogue (host and guest) via LLM, " &
                "then generates multi-voice audio. Use this when the user wants a podcast, dialogue, or discussion format. " &
                "(2) 'audiobook' — reads the text as narrated audio. " &
                "Use this when the user wants a straightforward narration or audio version of a document. " &
                "The text parameter accepts the full text to convert. For document attachments, first use " &
                "read_attachment to extract the text, then pass the extracted text to this tool. " &
                "For podcast mode, speaker names and TTS voice IDs are different. " &
                "Use host_name and guest_name for the names of the persons in the generated script. " &
                "Use host_gender and guest_gender to let the code select appropriate voices. " &
                "If the user says 'Marie and Marc' or gives other person names, treat them as speaker names, not as TTS voice IDs. " &
                "Do NOT put ordinary human names into voice_a or voice_b. " &
                "Only use voice_a or voice_b when the user explicitly provides a real provider voice ID such as an OpenAI voice name or a full Google voice ID. " &
                "For audiobook mode, use single_voice=true when the user wants a normal single narrator or simple audio file. " &
                "Use narrator_gender to choose the narrator voice for single-voice audiobook mode. " &
                "Only use single_voice=false when the user explicitly wants alternating narration voices. " &
                "If no speaker names are provided in podcast mode, choose suitable names. " &
                "If no genders are provided in podcast mode, choose sensible genders from context; if still unclear, use female for the host and male for the guest. " &
                "Available OpenAI voices: alloy (female), ash (male), ballad (male), coral (female), echo (male), " &
                "fable (male), nova (female), onyx (male), sage (male), shimmer (female), verse (male). " &
                "IMPORTANT: You MUST provide the 'language' parameter and it MUST match the language of the text content " &
                "using a full BCP-47 locale such as 'de-DE', 'de-CH', 'fr-FR', or 'en-US'. " &
                "Do NOT omit language and do NOT rely on defaults. " &
                "For German text, especially when it contains umlauts such as ä, ö, ü, or ß, use 'de-DE' or 'de-CH' as appropriate. " &
                "Preserve the original spelling in the text exactly as written; do not replace umlauts or ß with simplified forms." &
                engineHint &
                " For podcast mode, the duration parameter controls target script length (e.g., '5 minutes', '10 minutes'). " &
                "Output is an MP3 file attached to the reply.",
                .ToolDefinition =
                    "{""name"":""" & AP_Tool_CreateAudioFile & """," &
                    """description"":""Generates an MP3 audio file from text using text-to-speech. " &
                    "Supports 'podcast' mode (two-speaker dialogue generated by LLM) and 'audiobook' mode " &
                    "(single-voice or alternating narration). " &
                    "For document attachments, first extract text using read_attachment, then pass it here. " &
                    "CRITICAL: You MUST set 'language' to the actual language of the text using a full BCP-47 locale such as " &
                    "'de-DE', 'de-CH', 'fr-FR', or 'en-US'. Do NOT omit it. " &
                    "For German text with umlauts (ä, ö, ü, ß), use 'de-DE' or 'de-CH' and preserve the original spelling exactly. " &
                    "For non-English text, prefer engine='google' for native pronunciation. " &
                    "Use host_name and guest_name for podcast speaker names, host_gender and guest_gender for podcast voice selection, " &
                    "single_voice and narrator_gender for simple audiobook narration, and reserve voice_a and voice_b for explicit provider voice IDs only.""," &
                    """parameters"":{""type"":""object"",""properties"":{" &
                    """text"":{""type"":""string"",""description"":""The full text content to convert to audio. " &
                    "Preserve original spelling exactly, including diacritics and special characters such as ä, ö, ü, and ß. " &
                    "For large documents, pass the complete extracted text.""}," &
                    """mode"":{""type"":""string"",""enum"":[""podcast"",""audiobook""]," &
                    """description"":""Audio generation mode. 'podcast' = LLM-generated two-speaker dialogue. " &
                    "'audiobook' = narrated audio. Default: 'audiobook'""}," &
                    """language"":{""type"":""string"",""description"":""REQUIRED. Full BCP-47 language code matching the text content, " &
                    "for example 'en-US', 'de-DE', 'de-CH', or 'fr-FR'. This is critical for correct pronunciation. " &
                    "For German text with umlauts, use 'de-DE' or 'de-CH' as appropriate.""}," &
                    """engine"":{""type"":""string"",""enum"":[""google"",""openai"",""auto""]," &
                    """description"":""TTS engine to use. 'google' = Google Cloud TTS (best for non-English languages). " &
                    "'openai' = OpenAI TTS (best for English). 'auto' = automatically select based on language. Default: 'auto'""}," &
                    """host_name"":{""type"":""string"",""description"":""Podcast mode only. Name of the host / speaker A. This is a person name for the generated script, not a TTS voice ID.""}," &
                    """guest_name"":{""type"":""string"",""description"":""Podcast mode only. Name of the guest / speaker B. This is a person name for the generated script, not a TTS voice ID.""}," &
                    """host_gender"":{""type"":""string"",""enum"":[""female"",""male""],""description"":""Podcast mode only. Gender of the host / speaker A. Used to select the voice. If omitted, the caller should choose sensibly; default fallback is female.""}," &
                    """guest_gender"":{""type"":""string"",""enum"":[""female"",""male""],""description"":""Podcast mode only. Gender of the guest / speaker B. Used to select the voice. If omitted, the caller should choose sensibly; default fallback is male.""}," &
                    """single_voice"":{""type"":""boolean"",""description"":""Audiobook mode only. true = one narrator voice for the whole audio. false = alternating narration voices. For a normal simple audio file, use true.""}," &
                    """narrator_gender"":{""type"":""string"",""enum"":[""female"",""male""],""description"":""Audiobook mode only. Narrator gender for single_voice=true. Used to select the narrator voice. Default fallback is female.""}," &
                    """voice_a"":{""type"":""string"",""description"":""Optional explicit provider TTS voice ID for host / narrator A. Do NOT use ordinary human names here. Example OpenAI: 'nova'. Example Google: 'de-DE-Chirp3-HD-Achernar'.""}," &
                    """voice_b"":{""type"":""string"",""description"":""Optional explicit provider TTS voice ID for guest / narrator B. Do NOT use ordinary human names here. Example OpenAI: 'ash'. Example Google: 'de-DE-Chirp3-HD-Achird'.""}," &
                    """duration"":{""type"":""string"",""description"":""Target duration for podcast mode (e.g., '5 minutes', '10 minutes', '20 minutes'). Ignored in audiobook mode. Default: '5 minutes'""}," &
                    """instructions"":{""type"":""string"",""description"":""Specific instructions for the podcast script generation (e.g., 'Focus on the financial details', 'Be humorous', 'Explain for a child'). Optional.""}," &
                    """context"":{""type"":""string"",""description"":""Additional context or background info to help generating the podcast script. Optional.""}," &
                    """output_filename"":{""type"":""string"",""description"":""Filename for the output MP3 file (default: 'audiobook.mp3')""}" &
                    "},""required"":[""text"",""language""]}}"
            })
        End If

        ' ── generate_image ──
        ' Only register if an ImageGeneration special task model is configured
        Dim imgModelHasObjectCall As Boolean = False
        Dim imgModelAvailable As Boolean = IsImageGenerationAvailable(_context, imgModelHasObjectCall)

        If imgModelAvailable Then
            Dim imageEditHint As String = ""
            Dim imageEditParam As String = ""
            If imgModelHasObjectCall Then
                imageEditHint = " Supports image editing/modification: if the user provides a reference image " &
                    "(e.g. an attached photo to modify, a logo to restyle, a sketch to refine), pass its filename " &
                    "as 'image_attachment_name'. The reference image is sent alongside the description to the model. " &
                    "Only use this when the user explicitly wants to modify, edit, or base the generation on an existing image."
                imageEditParam = """image_attachment_name"":{""type"":""string"",""description"":""Optional: filename of an existing image attachment to use as reference for editing or modification. " &
                    "The model will use this image as a base and apply the changes described in 'description'. " &
                    "Only use when the user wants to modify, restyle, or build upon an existing image.""},"
            End If

            tools.Add(New ModelConfig() With {
                .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_GenerateImage,
                .CapabilityTags = "explicit_operation",
                .ModelDescription = "Generate Image (built-in)",
                .ToolInstructionsPrompt =
                    AP_Tool_GenerateImage & ": Generates an image from a text description using an image generation model. " &
                    "Use this when the user asks you to create, generate, draw, design, or produce an image, picture, illustration, " &
                    "diagram, logo, icon, or any visual content. " &
                    "The 'description' parameter should contain ONLY the image description — do NOT wrap it in any system prompt " &
                    "or additional instructions. Pass the user's visual intent directly as the description. " &
                    "The generated image is saved as a file and attached to the reply. " &
                    "You can optionally specify a filename for the output image." &
                    imageEditHint,
                .ToolDefinition =
                    "{""name"":""" & AP_Tool_GenerateImage & """," &
                    """description"":""Generates an image from a text description using an AI image generation model. " &
                    "Use when the user wants to create, generate, draw, or design any visual content (image, illustration, diagram, logo, etc.). " &
                    "Pass the image description directly — no wrapping prompt needed." &
                    If(imgModelHasObjectCall, " Also supports editing/modifying an existing image when image_attachment_name is provided.", "") & """," &
                    """parameters"":{""type"":""object"",""properties"":{" &
                    """description"":{""type"":""string"",""description"":""A clear, detailed description of the image to generate. " &
                    "Pass the user's visual intent directly. Do not add system instructions or prompt wrappers.""}," &
                    imageEditParam &
                    """output_filename"":{""type"":""string"",""description"":""Optional filename for the output image (default: auto-generated). " &
                    "Do not include an extension — the format is determined by the model.""}" &
                    "},""required"":[""description""]}}"
            })
        End If

        ' ── web_grounding ──
        If _apConfig IsNot Nothing AndAlso _apConfig.EnableWebGrounding Then
            Dim webGroundingTool =
                SharedLibrary.Agents.WebGroundingTool.Build(
                    _context,
                    enforcePrivacy:=_apConfig.EnablePrivacyProtection,
                    toolPriority:=997,
                    displaySuffix:="")

            If webGroundingTool IsNot Nothing Then
                tools.Add(webGroundingTool)
            End If
        End If


        ' ── manage_scheduled_tasks ──
        If (_apConfig IsNot Nothing AndAlso _apConfig.EnableScheduler) OrElse
           (Not _apActive AndAlso INI_AutoPilotSchedulerLocalChat AndAlso INI_WebServerBlock <> 4) Then
            tools.Add(BuildManageScheduledTasksTool())
        End If

        ' ── manage_user_memory ──
        If _apConfig IsNot Nothing AndAlso _apConfig.EnableUserMemory Then
            tools.Add(New ModelConfig() With {
                .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ManageUserMemory,
                .ModelDescription = "Manage User Memory (built-in)",
                .ToolInstructionsPrompt =
                    AP_Tool_ManageUserMemory & ": Manages the per-user persistent memory for the e-mail sender. " &
                    "Memory stores user preferences, corrections, and working style that persist across sessions. " &
                    "Actions: " &
                    "'enable' — activates memory for this user (opt-in). " &
                    "'disable' — deactivates memory and deletes all stored preferences (opt-out). " &
                    "'list' — returns all current memory items. " &
                    "'add' — adds a new memory item (value parameter required). " &
                    "'remove' — removes a memory item by fuzzy match (value parameter required). " &
                    "'amend' — changes an existing item (value = old text, new_value = replacement). " &
                    "'clear' — removes all memory items but keeps memory enabled. " &
                    "'toggle_auto_learn' — turns automatic learning on or off (auto_learn parameter: true/false). " &
                    "When auto_learn is on (default when memory is enabled), the model automatically detects and stores " &
                    "user preferences from conversations using <INKY_MEMORY> blocks. " &
                    "When the user asks you to remember, forget, or change a preference, use this tool. " &
                    "The sender_email is automatically determined from the incoming e-mail — never ask the user for it.",
                .ToolDefinition =
                    "{""name"":""" & AP_Tool_ManageUserMemory & """," &
                    """description"":""Manages per-user persistent memory (preferences, corrections, working style). " &
                    "Memory persists across AutoPilot sessions. Users can enable/disable, add/remove/amend items, " &
                    "or toggle automatic learning.""," &
                    """parameters"":{""type"":""object"",""properties"":{" &
                    """action"":{""type"":""string"",""enum"":[""enable"",""disable"",""list"",""add"",""remove"",""amend"",""clear"",""toggle_auto_learn""]," &
                    """description"":""The operation to perform""}," &
                    """value"":{""type"":""string"",""description"":""Memory item text (for add/remove) or old text (for amend)""}," &
                    """new_value"":{""type"":""string"",""description"":""Replacement text (for amend action only)""}," &
                    """auto_learn"":{""type"":""boolean"",""description"":""Enable or disable automatic learning (for toggle_auto_learn)""}" &
                    "},""required"":[""action""]}}"
            })
        End If

        ' ── manage_user_files ──
        If _apConfig IsNot Nothing AndAlso _apConfig.EnableUserFiles Then
            tools.Add(New ModelConfig() With {
                .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ManageUserFiles,
                .ModelDescription = "Manage User Files (built-in)",
                .ToolInstructionsPrompt =
                    AP_Tool_ManageUserFiles & ": Manages the per-user persistent file storage (home directory). " &
                    "Users can store files (templates, letterheads, reference documents) that persist across sessions " &
                    "and are available for use in future requests. " &
                    "Actions: " &
                    "'list' — returns all files in the user's home directory with names and sizes. " &
                    "'add' — copies a file from the current e-mail's attachments into the user's home directory. " &
                    "The file_name parameter specifies which attachment to store. An optional target_name " &
                    "renames the file. IMPORTANT: Inform the user that stored files will be visible to " &
                    "all future AutoPilot sessions and can be used by document processing tools. " &
                    "'remove' — deletes a file from the user's home directory. " &
                    "'replace' — replaces an existing file with a new version from the current attachments. " &
                    "'checkout' — retrieves a file from the home directory and includes it as an attachment " &
                    "in the reply to the user. " &
                    "'use' — loads a file from the home directory into the current processing session so that " &
                    "other tools (e.g. process_word_document) can reference it by name. The file is NOT " &
                    "returned to the user unless explicitly requested. " &
                    "Per-user storage is capped at " & CStr(AP_UserHomeMaxBytes \ 1024 \ 1024) & " MB. " &
                    "The sender_email is automatically determined from the incoming e-mail.",
                .ToolDefinition =
                    "{""name"":""" & AP_Tool_ManageUserFiles & """," &
                    """description"":""Manages per-user persistent file storage. Users can store, list, remove, replace, " &
                    "checkout (retrieve), or use (load into session) files that persist across AutoPilot sessions. " &
                    "Storage is capped at " & CStr(AP_UserHomeMaxBytes \ 1024 \ 1024) & " MB per user.""," &
                    """parameters"":{""type"":""object"",""properties"":{" &
                    """action"":{""type"":""string"",""enum"":[""list"",""add"",""remove"",""replace"",""checkout"",""use""]," &
                    """description"":""The operation to perform""}," &
                    """file_name"":{""type"":""string"",""description"":""Filename of the attachment (for add/replace) or home file (for remove/checkout/use)""}," &
                    """target_name"":{""type"":""string"",""description"":""Optional: rename the file when storing (for add/replace). If omitted, uses the original name.""}" &
                    "},""required"":[""action""]}}"
            })
        End If


        ' ── data collector ──
        If IsDataCollectorToolAvailable() Then
            tools.Add(BuildListCollectionUseCasesTool())
            tools.Add(BuildCollectDataTool())
            tools.Add(BuildPreviewCollectionTool())
        End If

        ' ── report_inability ──

        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ReportInability,
            .ToolPriority = 9999,
            .ModelDescription = "Report Inability to Fulfill Request (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ReportInability & ": Call this tool when you determine that you CANNOT fulfill the user's request " &
                "with the available tools and capabilities. Provide a brief reason. " &
                "The tool returns helpful suggestions for the user. You MUST naturally incorporate " &
                "the returned content into your reply — do NOT add labels, headers, or prefixes around it. " &
                "You MUST call this tool instead of simply telling the user you cannot help. " &
                "Also call this tool when attachments exceed the size limit and cannot be processed.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ReportInability & """," &
                """description"":""Call this when you cannot fulfill the user's request. Provide the reason. " &
                "The tool returns helpful suggestions for the user. Naturally incorporate the returned text " &
                "into your reply without adding labels or headers around it. " &
                "Always call this instead of simply saying you cannot help.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """reason"":{""type"":""string"",""description"":""Brief reason why the request cannot be fulfilled (e.g. 'attachment exceeds size limit', 'no tool available for image generation', 'task requires manual interaction')""}" &
                "},""required"":[""reason""]}}"
        })

        For Each tool As ModelConfig In tools
            If tool Is Nothing Then Continue For
            tool.ModelDescription = StripSelectorOwnedToolSuffixes(tool.ModelDescription)
            If IsOptionalSingleFileAutoPilotArtifactTool(tool.ToolName) Then
                SharedLibrary.Agents.ArtifactDelivery.EnableOptionalSingleFileArtifactProtocol(tool)
            End If
        Next

        Return tools

    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL DISPATCH
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Resolves and executes one AutoPilot internal tool call.
    ''' </summary>
    ''' <param name="toolCall">The parsed tool invocation payload.</param>
    ''' <param name="context">Execution context used for logging and correlation.</param>
    ''' <param name="cancellationToken">Optional cancellation token for async operations.</param>
    ''' <returns>
    ''' A <see cref="ToolResponse"/> when the tool is recognized; otherwise <c>Nothing</c>
    ''' so the caller can continue with external tool handling.
    ''' </returns>
    Private Shared Function IsOptionalSingleFileAutoPilotArtifactTool(toolName As String) As Boolean
        If String.IsNullOrWhiteSpace(toolName) Then Return False
        Select Case toolName
            Case AP_Tool_MergePdfs,
                 AP_Tool_CompareWordDocs,
                 AP_Tool_CreatePdfFromText,
                 AP_Tool_ExcelCompleteLiveWorkbook,
                 AP_Tool_AddPdfWatermark,
                 AP_Tool_WordToPdf,
                 AP_Tool_PdfToWord,
                 AP_Tool_CreateWordDoc,
                 AP_Tool_CreateExcel,
                 AP_Tool_CreatePowerPoint,
                 AP_Tool_CreateCodeFile,
                 AP_Tool_RedactPdf,
                 AP_Tool_OverlayPdf,
                 AP_Tool_CreateAudioFile,
                 AP_Tool_GenerateImage,
                 AP_Tool_ManageUserFiles
                Return True
        End Select
        Return False
    End Function

    Private Shared Function HasExplicitArtifactIdentityArguments(arguments As IDictionary(Of String, Object)) As Boolean
        If arguments Is Nothing Then Return False
        For Each key As String In New String() {
            "artifact_id", "logical_deliverable_id", "output_slot_id", "supersedes_artifact_id",
            "artifact_state", "artifact_delivery_intent", "storage_kind"
        }
            Dim value As Object = Nothing
            If arguments.TryGetValue(key, value) AndAlso value IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(value.ToString()) Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Shared Function ComputeAutoPilotOutputFingerprint(path As String) As String
        Try
            If String.IsNullOrWhiteSpace(path) OrElse Not File.Exists(path) Then Return "missing"
            Using sha As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Using stream As FileStream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite Or FileShare.Delete)
                    Dim hash As Byte() = sha.ComputeHash(stream)
                    Return "sha256:" & BitConverter.ToString(hash).Replace("-", "")
                End Using
            End Using
        Catch
            Try
                Dim info As New FileInfo(path)
                Return "meta:" & info.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) & ":" & info.LastWriteTimeUtc.Ticks.ToString(System.Globalization.CultureInfo.InvariantCulture)
            Catch
                Return "unavailable"
            End Try
        End Try
    End Function

    Friend Async Function TryExecuteAutoPilotTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            Optional cancellationToken As CancellationToken = Nothing) As System.Threading.Tasks.Task(Of ToolResponse)

        Dim artifactMetadata As SharedLibrary.Agents.OptionalToolArtifactMetadata = Nothing
        If HasExplicitArtifactIdentityArguments(toolCall.Arguments) Then
            If toolCall.ToolName.Equals(AP_Tool_ManageUserFiles, StringComparison.OrdinalIgnoreCase) Then
                Dim manageAction As String = If(GetArgString(toolCall.Arguments, "action"), "").Trim()
                If Not manageAction.Equals("checkout", StringComparison.OrdinalIgnoreCase) Then
                    Return New ToolResponse() With {
                        .CallId = toolCall.CallId,
                        .ToolName = toolCall.ToolName,
                        .Success = False,
                        .ResultKind = "error",
                        .ErrorCode = "explicit_artifact_requires_checkout_action",
                        .ErrorMessage = "manage_user_files may use explicit artifact metadata only with action='checkout', because other actions do not produce one user-facing output file."
                    }
                End If
            End If

            If Not IsOptionalSingleFileAutoPilotArtifactTool(toolCall.ToolName) Then
                Return New ToolResponse() With {
                    .CallId = toolCall.CallId,
                    .ToolName = toolCall.ToolName,
                    .Success = False,
                    .ResultKind = "error",
                    .ErrorCode = "explicit_artifact_not_supported_for_tool_cardinality",
                    .ErrorMessage = "This tool can produce zero, multiple, or dynamic files and cannot safely bind one explicit artifact identity to the whole call. Use its legacy/multi-output flow or a tool with an explicit per-output slot protocol."
                }
            End If

            Dim artifactFailureCode As String = ""
            Dim artifactFailureMessage As String = ""
            If Not SharedLibrary.Agents.ArtifactDelivery.TryPrepareOptionalToolArtifactMetadata(
                toolCall.Arguments,
                SharedLibrary.Agents.ArtifactStorageKind.Unknown,
                artifactMetadata,
                artifactFailureCode,
                artifactFailureMessage) Then

                Return New ToolResponse() With {
                    .CallId = toolCall.CallId,
                    .ToolName = toolCall.ToolName,
                    .Success = False,
                    .ResultKind = "error",
                    .ErrorCode = artifactFailureCode,
                    .ErrorMessage = artifactFailureMessage
                }
            End If
        End If

        Dim outputSnapshot As Dictionary(Of String, String) = SnapshotAutoPilotOutputFiles()
        Dim response As ToolResponse = Nothing
        Dim enableLocalToolingMirror As Boolean = _chatAgentActive AndAlso Not _apActive

        If enableLocalToolingMirror Then
            System.Threading.Interlocked.Increment(_apMirrorDashboardLogToLocalToolingDepth)
        End If

        Try
            Select Case toolCall.ToolName
                Case AP_Tool_ProcessWordDoc
                    response = Await ExecuteProcessWordDocTool(toolCall, context, cancellationToken)
                Case AP_Tool_ExtractPdfText
                    response = Await ExecuteExtractPdfTextTool(toolCall, context, cancellationToken)
                Case AP_Tool_MergePdfs
                    response = Await ExecuteMergePdfsTool(toolCall, context, cancellationToken)
                Case AP_Tool_ReadAttachment
                    response = Await ExecuteReadAttachmentTool(toolCall, context, cancellationToken)
                Case AP_Tool_ListAttachments
                    response = ExecuteListAttachmentsTool(toolCall, context)
                Case AP_Tool_DescribeBinary
                    response = Await ExecuteDescribeBinaryTool(toolCall, context, cancellationToken)
                Case AP_Tool_CommentWordDoc
                    response = Await ExecuteCommentWordDocTool(toolCall, context, cancellationToken)
                Case AP_Tool_CommentPdf
                    response = Await ExecuteCommentPdfTool(toolCall, context, cancellationToken)
                Case AP_Tool_CompareWordDocs
                    response = Await ExecuteCompareWordDocsTool(toolCall, context, cancellationToken)
                Case AP_Tool_ReadWordDocDetails
                    response = Await ExecuteReadWordDocDetailsTool(toolCall, context, cancellationToken)
                Case AP_Tool_CreatePdfFromText
                    response = ExecuteCreatePdfFromTextTool(toolCall, context)
                Case AP_Tool_ExtractExcelData
                    response = ExecuteExtractExcelDataTool(toolCall, context)
                Case AP_Tool_ExcelListLiveWorksheets
                    response = Await ExecuteExcelListLiveWorksheetsTool(toolCall, context, cancellationToken)
                Case AP_Tool_ExcelReadLiveRange
                    response = Await ExecuteExcelReadLiveRangeTool(toolCall, context, cancellationToken)
                Case AP_Tool_ExcelCompleteLiveWorkbook
                    response = Await ExecuteExcelCompleteLiveWorkbookTool(toolCall, context, cancellationToken)
                Case AP_Tool_SplitPdf
                    response = ExecuteSplitPdfTool(toolCall, context)
                Case AP_Tool_AddPdfWatermark
                    response = ExecuteAddPdfWatermarkTool(toolCall, context)
                Case AP_Tool_WordToPdf
                    response = Await ExecuteWordToPdfTool(toolCall, context, cancellationToken)
                Case AP_Tool_SearchInAttachments
                    response = Await ExecuteSearchInAttachmentsTool(toolCall, context, cancellationToken)
                Case AP_Tool_SummarizeThread
                    response = ExecuteSummarizeThreadTool(toolCall, context)
                Case AP_Tool_PdfToWord
                    response = Await ExecutePdfToWordTool(toolCall, context, cancellationToken)
                Case AP_Tool_CreateWordDoc
                    response = Await ExecuteCreateWordDocTool(toolCall, context, cancellationToken)
                Case AP_Tool_CreateExcel
                    response = Await ExecuteCreateExcelTool(toolCall, context, cancellationToken)
                Case AP_Tool_CreatePowerPoint
                    response = Await ExecuteCreatePowerPointTool(toolCall, context, cancellationToken)
                Case AP_Tool_CreateCodeFile
                    response = Await ExecuteCreateCodeFileTool(toolCall, context, cancellationToken)
                Case AP_Tool_ExtractDataFromAttachments
                    response = Await ExecuteExtractDataFromAttachmentsTool(toolCall, context, cancellationToken)
                Case AP_Tool_RedactPdf
                    response = Await ExecuteRedactPdfTool(toolCall, context, cancellationToken)
                Case AP_Tool_OverlayPdf
                    response = Await ExecuteOverlayPdfTool(toolCall, context, cancellationToken)
                Case AP_Tool_CreateAudioFile
                    response = Await ExecuteCreateAudioFileTool(toolCall, context, cancellationToken)
                Case AP_Tool_GenerateImage
                    response = Await ExecuteGenerateImageTool(toolCall, context, cancellationToken)
                Case AP_Tool_WebGrounding
                    response = Await ExecuteWebGroundingTool(toolCall, context, cancellationToken)
                Case AP_Tool_ManageScheduledTasks
                    response = Await ExecuteManageScheduledTasksTool(toolCall, context, cancellationToken)
                Case AP_Tool_ManageUserMemory
                    response = Await ExecuteManageUserMemoryTool(toolCall, context, cancellationToken)
                Case AP_Tool_ManageUserFiles
                    response = Await ExecuteManageUserFilesTool(toolCall, context, cancellationToken)
                Case AP_Tool_ListCollectionUseCases
                    response = Await ExecuteListCollectionUseCasesTool(toolCall, context, cancellationToken)
                Case AP_Tool_CollectData
                    response = Await ExecuteCollectDataTool(toolCall, context, cancellationToken)
                Case AP_Tool_PreviewCollection
                    response = Await ExecutePreviewCollectionTool(toolCall, context, cancellationToken)
                Case AP_Tool_CompleteWordTables
                    response = Await ExecuteCompleteWordTablesTool(toolCall, context, cancellationToken)
                Case AP_Tool_ReportInability
                    response = Await ExecuteReportInabilityTool(toolCall, context, cancellationToken)
                Case Agents.PythonExecuteTool.ToolName
                    response = Await ExecutePythonExecuteTool(toolCall, context, cancellationToken)
                Case Else
                    Return Nothing
            End Select

            Return NormalizeAutoPilotToolResponse(toolCall, response, outputSnapshot, artifactMetadata)
        Finally
            If enableLocalToolingMirror Then
                System.Threading.Interlocked.Decrement(_apMirrorDashboardLogToLocalToolingDepth)
            End If
        End Try
    End Function

    ''' <summary>
    ''' Registers one physical file created by an AutoPilot/Local-Agent tool as an explicit
    ''' host-side output. This also works when the current mail/session has no inbound
    ''' attachments; in that case a tool-output carrier is added so per-call fingerprinting,
    ''' ArtifactDelivery normalization, and bounded legacy compatibility can all observe the file.
    ''' No artifact identity is inferred from the path.
    ''' </summary>
    Private Sub RegisterAutoPilotGeneratedOutputFile(outputPath As String)
        If String.IsNullOrWhiteSpace(outputPath) Then Return

        Dim normalizedPath As String

        Try
            normalizedPath = System.IO.Path.GetFullPath(outputPath)
        Catch ex As System.Exception
            Return
        End Try

        If Not System.IO.File.Exists(normalizedPath) Then Return

        If _apCurrentAttachments Is Nothing Then
            _apCurrentAttachments = New List(Of AutoPilotAttachmentInfo)()
        End If

        Dim registrationTarget As AutoPilotAttachmentInfo = Nothing

        For Each existingAttachment As AutoPilotAttachmentInfo In _apCurrentAttachments
            If existingAttachment Is Nothing Then Continue For

            If existingAttachment.OutputFiles IsNot Nothing Then
                For Each existingOutput As String In existingAttachment.OutputFiles
                    If String.Equals(
                        If(existingOutput, String.Empty).Trim(),
                        normalizedPath,
                        StringComparison.OrdinalIgnoreCase) Then

                        Return
                    End If
                Next
            End If

            If registrationTarget Is Nothing Then
                registrationTarget = existingAttachment
            End If
        Next

        If registrationTarget Is Nothing Then
            Dim fileInfo As New System.IO.FileInfo(normalizedPath)

            registrationTarget = New AutoPilotAttachmentInfo With {
                .OriginalFileName = System.IO.Path.GetFileName(normalizedPath),
                .TempFilePath = normalizedPath,
                .SourcePath = normalizedPath,
                .Extension = System.IO.Path.GetExtension(normalizedPath),
                .SizeBytes = fileInfo.Length,
                .IsOverSizeLimit = False,
                .StatusMessage = "Generated tool output",
                .CreatedTime = fileInfo.CreationTime,
                .LastModifiedTime = fileInfo.LastWriteTime,
                .IsToolOutput = True,
                .OutputFiles = New List(Of String)()
            }

            _apCurrentAttachments.Add(registrationTarget)
        ElseIf registrationTarget.OutputFiles Is Nothing Then
            registrationTarget.OutputFiles = New List(Of String)()
        End If

        registrationTarget.OutputFiles.Add(normalizedPath)
    End Sub

    Private Function SnapshotAutoPilotOutputFiles() As Dictionary(Of String, String)
        Dim snapshot As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        If _apCurrentAttachments Is Nothing Then
            Return snapshot
        End If

        For Each att In _apCurrentAttachments
            If att Is Nothing OrElse att.OutputFiles Is Nothing Then Continue For

            For Each outputPath In att.OutputFiles
                Dim normalized As String = If(outputPath, "").Trim()
                If normalized = "" Then Continue For

                snapshot(normalized) = ComputeAutoPilotOutputFingerprint(normalized)
            Next
        Next

        Return snapshot
    End Function

    Private Function GetProducedAutoPilotOutputFiles(previousSnapshot As IDictionary(Of String, String)) As List(Of String)
        Dim produced As New List(Of String)()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If _apCurrentAttachments Is Nothing Then
            Return produced
        End If

        For Each att In _apCurrentAttachments
            If att Is Nothing OrElse att.OutputFiles Is Nothing Then Continue For

            For Each outputPath In att.OutputFiles
                Dim normalized As String = If(outputPath, "").Trim()
                If normalized = "" Then Continue For
                If Not File.Exists(normalized) Then Continue For

                Dim currentStamp As String = ComputeAutoPilotOutputFingerprint(normalized)
                Dim previousStamp As String = Nothing
                Dim wasKnown As Boolean =
                    previousSnapshot IsNot Nothing AndAlso
                    previousSnapshot.TryGetValue(normalized, previousStamp)

                If Not wasKnown OrElse Not String.Equals(previousStamp, currentStamp, StringComparison.Ordinal) Then
                    If seen.Add(normalized) Then
                        produced.Add(normalized)
                    End If
                End If
            Next
        Next

        Return produced
    End Function

    Private Shared Function TryParseToolResultToken(responseText As String) As JToken
        Dim raw As String = If(responseText, "").Trim()
        If raw = "" Then
            Return Nothing
        End If

        Try
            Return JToken.Parse(raw)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function HasNormalizedToolResultMetadata(responseText As String) As Boolean
        Dim obj As JObject = TryCast(TryParseToolResultToken(responseText), JObject)
        If obj Is Nothing Then
            Return False
        End If

        Return obj("producesIntermediateData") IsNot Nothing OrElse
               obj("producesUserDeliverable") IsNot Nothing OrElse
               obj("outputArtifactRef") IsNot Nothing OrElse
               obj("outputFilePath") IsNot Nothing OrElse
               obj("outputFileName") IsNot Nothing OrElse
               obj("created") IsNot Nothing OrElse
               obj("saved") IsNot Nothing OrElse
               obj("exported") IsNot Nothing
    End Function

    Private Function BuildNormalizedAutoPilotToolResult(toolName As String,
                                                        summary As String,
                                                        producesIntermediateData As Boolean,
                                                        producesUserDeliverable As Boolean,
                                                        outputFiles As IList(Of String),
                                                        sourcePayload As JToken) As String
        Dim normalizedSummary As String = If(summary, "").Trim()

        If normalizedSummary = "" Then
            normalizedSummary =
                If(producesUserDeliverable,
                   "User-facing deliverable created.",
                   "Structured tool result available.")
        End If

        Dim obj As New JObject(
            New JProperty("toolName", If(toolName, "")),
            New JProperty("summary", normalizedSummary),
            New JProperty("producesIntermediateData", producesIntermediateData),
            New JProperty("producesUserDeliverable", producesUserDeliverable))

        If producesUserDeliverable Then
            obj("created") = True
        End If

        If outputFiles IsNot Nothing AndAlso outputFiles.Count > 0 Then
            Dim outputNames As New JArray()

            For Each outputPath In outputFiles
                Dim outputName As String = Path.GetFileName(If(outputPath, ""))
                If outputName = "" Then Continue For
                outputNames.Add(outputName)
            Next

            If outputNames.Count > 0 Then
                Dim firstOutputName As String = outputNames(0).ToString()
                obj("outputFileName") = firstOutputName
                obj("outputFilePath") = firstOutputName
                obj("outputArtifactRef") = firstOutputName
                obj("outputFiles") = outputNames
            End If
        End If

        If sourcePayload IsNot Nothing Then
            obj("result") = sourcePayload
        End If

        Return obj.ToString(Newtonsoft.Json.Formatting.None)
    End Function

    Private Function NormalizeAutoPilotToolResponse(toolCall As ToolCall,
                                                    toolResponse As ToolResponse,
                                                    outputSnapshot As IDictionary(Of String, String),
                                                    Optional artifactMetadata As SharedLibrary.Agents.OptionalToolArtifactMetadata = Nothing) As ToolResponse
        If toolResponse Is Nothing OrElse Not toolResponse.Success Then
            Return toolResponse
        End If

        Dim parsedToken As JToken = TryParseToolResultToken(toolResponse.Response)
        Dim producedOutputs As List(Of String) = GetProducedAutoPilotOutputFiles(outputSnapshot)

        If artifactMetadata IsNot Nothing Then
            If producedOutputs.Count = 1 Then
                toolResponse.Response = SharedLibrary.Agents.ArtifactDelivery.AttachOptionalSingleFileArtifactToResult(
                    toolResponse.Response,
                    artifactMetadata,
                    producedOutputs(0))
                toolResponse.ResultKind = "json_object"
                Return toolResponse
            End If

            ' A safe single-file producer may legitimately produce no file (for example
            ' a no-op/finalize check). In that case no artifact is fabricated; a locked
            ' expected-artifact contract remains unsatisfied and completion is rejected.
            If producedOutputs.Count = 0 Then Return toolResponse

            toolResponse.Success = False
            toolResponse.ResultKind = "error"
            toolResponse.ErrorCode = "single_file_artifact_cardinality_violation"
            toolResponse.ErrorMessage = "The tool produced more than one changed output while a single explicit artifact identity was supplied."
            Return toolResponse
        End If

        If HasNormalizedToolResultMetadata(toolResponse.Response) Then
            If String.IsNullOrWhiteSpace(toolResponse.ResultKind) Then
                toolResponse.ResultKind =
                    If(TypeOf parsedToken Is JArray, "json_array", "json_object")
            End If

            Return toolResponse
        End If

        If producedOutputs.Count > 0 Then
            toolResponse.Response =
                BuildNormalizedAutoPilotToolResult(
                    If(toolCall?.ToolName, toolResponse.ToolName),
                    toolResponse.Response,
                    producesIntermediateData:=False,
                    producesUserDeliverable:=True,
                    outputFiles:=producedOutputs,
                    sourcePayload:=If(TypeOf parsedToken Is JObject OrElse TypeOf parsedToken Is JArray, parsedToken, Nothing))
            toolResponse.ResultKind = "json_object"
            Return toolResponse
        End If

        If TypeOf parsedToken Is JObject OrElse TypeOf parsedToken Is JArray Then
            toolResponse.Response =
                BuildNormalizedAutoPilotToolResult(
                    If(toolCall?.ToolName, toolResponse.ToolName),
                    "Structured tool result available.",
                    producesIntermediateData:=True,
                    producesUserDeliverable:=False,
                    outputFiles:=Nothing,
                    sourcePayload:=parsedToken)
            toolResponse.ResultKind = "json_object"
        End If

        Return toolResponse
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  HELPER: Get argument values
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Reads a string argument from a tool-argument dictionary.
    ''' </summary>
    ''' <param name="args">Argument dictionary from the tool call.</param>
    ''' <param name="key">Argument key to read.</param>
    ''' <returns>The string value if present; otherwise <c>Nothing</c>.</returns>
    Private Shared Function GetArgString(args As Dictionary(Of String, Object), key As String) As String
        If args Is Nothing OrElse Not args.ContainsKey(key) Then Return Nothing
        Return args(key)?.ToString()
    End Function

    ''' <summary>
    ''' Reads a Boolean argument with fallback default.
    ''' </summary>
    Private Shared Function GetArgBool(args As Dictionary(Of String, Object), key As String, defaultVal As Boolean) As Boolean
        Dim s = GetArgString(args, key)
        If String.IsNullOrWhiteSpace(s) Then Return defaultVal
        Dim result As Boolean
        If Boolean.TryParse(s, result) Then Return result
        Return defaultVal
    End Function

    ''' <summary>
    ''' Reads an Integer argument with fallback default.
    ''' </summary>
    Private Shared Function GetArgInt(args As Dictionary(Of String, Object), key As String, defaultVal As Integer) As Integer
        Dim s = GetArgString(args, key)
        If String.IsNullOrWhiteSpace(s) Then Return defaultVal
        Dim result As Integer
        If Integer.TryParse(s, result) Then Return result
        Return defaultVal
    End Function

    ''' <summary>
    ''' Reads a JSON array argument as a list of strings.
    ''' </summary>
    Private Shared Function GetArgStringArray(args As Dictionary(Of String, Object), key As String) As List(Of String)
        Dim result As New List(Of String)()
        If args Is Nothing OrElse Not args.ContainsKey(key) Then Return result
        Dim namesObj = args(key)
        If TypeOf namesObj Is JArray Then
            For Each item In DirectCast(namesObj, JArray)
                result.Add(item.ToString())
            Next
        End If
        Return result
    End Function

    ''' <summary>
    ''' Finds an attachment by filename (case-insensitive) from either:
    ''' (1) original mail attachments, or
    ''' (2) output files produced by prior tool calls in the same run.
    ''' </summary>
    ''' <remarks>
    ''' If an output file is matched, this method returns a transient
    ''' <see cref="AutoPilotAttachmentInfo"/> marked with <c>IsToolOutput=True</c>.
    ''' </remarks>
    Private Function FindAttachment(fileName As String) As AutoPilotAttachmentInfo
        If String.IsNullOrWhiteSpace(fileName) OrElse _apCurrentAttachments Is Nothing Then Return Nothing

        ' Rescan the active session staging directory on every lookup so that files produced
        ' by any tool (file_copy, js_run outputs, completed workbooks, etc.) become resolvable
        ' even when the model never learned their name. Registration is deduped and cheap.
        RefreshSessionStagingAttachments()

        Dim trimmedName = fileName.Trim()

        Dim found = _apCurrentAttachments.FirstOrDefault(
            Function(a) a.OriginalFileName.Equals(trimmedName, StringComparison.OrdinalIgnoreCase))
        If found IsNot Nothing Then
            Return EnsureSessionAttachmentAvailable(found)
        End If

        For Each att In _apCurrentAttachments
            If att.OutputFiles Is Nothing Then Continue For
            For Each outputPath In att.OutputFiles
                If String.IsNullOrEmpty(outputPath) Then Continue For
                Dim outputName = Path.GetFileName(outputPath)
                If outputName.Equals(trimmedName, StringComparison.OrdinalIgnoreCase) AndAlso
                   File.Exists(outputPath) Then
                    Return New AutoPilotAttachmentInfo() With {
                        .OriginalFileName = outputName,
                        .Extension = Path.GetExtension(outputPath).ToLowerInvariant(),
                        .TempFilePath = outputPath,
                        .SizeBytes = New FileInfo(outputPath).Length,
                        .IsOverSizeLimit = False,
                        .StatusMessage = "Tool output",
                        .IsToolOutput = True,
                        .OutputFiles = New List(Of String)()
                    }
                End If
            Next
        Next

        ' Fallback: resolve the name against the granted workspace so that AutoPilot
        ' Office tools (process_word_document, comment_word_document, compare_word_documents,
        ' etc.) work on local workspace documents, not only on mail attachments.
        Dim workspaceMatch As AutoPilotAttachmentInfo = TryResolveWorkspaceFile(trimmedName)
        If workspaceMatch IsNot Nothing Then Return workspaceMatch

        Return Nothing
    End Function

    ''' <summary>
    ''' Attempts to resolve <paramref name="fileName"/> against the granted workspace root.
    ''' Returns a transient <see cref="AutoPilotAttachmentInfo"/> pointing at the workspace
    ''' file (marked <c>IsToolOutput=True</c> so it is treated as an in-session file), or
    ''' <c>Nothing</c> when no workspace is configured or the file is not found.
    ''' </summary>
    Private Function TryResolveWorkspaceFile(fileName As String) As AutoPilotAttachmentInfo
        If String.IsNullOrWhiteSpace(fileName) Then Return Nothing

        Dim state As SharedLibrary.Agents.WorkspaceState = SharedLibrary.Agents.WorkspaceStore.Load("outlook")
        If state Is Nothing OrElse String.IsNullOrWhiteSpace(state.RootPath) OrElse
           Not Directory.Exists(state.RootPath) Then
            Return Nothing
        End If

        Dim candidate As String
        Try
            ' Accept both a bare leaf name and a workspace-relative path.
            candidate = Path.GetFullPath(Path.Combine(state.RootPath, fileName))
        Catch
            Return Nothing
        End Try

        ' Confine resolution to the workspace root.
        Dim rootFull As String = Path.GetFullPath(state.RootPath)
        If Not candidate.StartsWith(rootFull, StringComparison.OrdinalIgnoreCase) Then Return Nothing
        If Not File.Exists(candidate) Then Return Nothing

        Return New AutoPilotAttachmentInfo() With {
            .OriginalFileName = Path.GetFileName(candidate),
            .Extension = Path.GetExtension(candidate).ToLowerInvariant(),
            .TempFilePath = candidate,
            .SourcePath = candidate,
            .SizeBytes = New FileInfo(candidate).Length,
            .IsOverSizeLimit = False,
            .StatusMessage = "Workspace file",
            .IsToolOutput = True,
            .OutputFiles = New List(Of String)()
        }
    End Function

    ''' <summary>
    ''' Returns all filenames currently available for tool resolution:
    ''' original attachments plus existing tool output files.
    ''' </summary>
    Private Function GetAllAvailableFileNames() As List(Of String)
        Dim names As New List(Of String)()
        If _apCurrentAttachments Is Nothing Then Return names
        For Each att In _apCurrentAttachments
            names.Add(att.OriginalFileName)
            If att.OutputFiles IsNot Nothing Then
                For Each outputPath In att.OutputFiles
                    If Not String.IsNullOrEmpty(outputPath) AndAlso File.Exists(outputPath) Then
                        names.Add(Path.GetFileName(outputPath))
                    End If
                Next
            End If
        Next
        Return names
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  HELPER: Read single attachment text (with caching)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Reads text from a single attachment, using cache when available.
    ''' Prefers sandboxed (COM-free) readers for OpenXML and mail formats.
    ''' Respects <see cref="INI_AllowLegacyDocFiles"/> for .doc files.
    ''' </summary>
    Private Async Function ReadSingleAttachmentText(att As AutoPilotAttachmentInfo,
                                                    context As ToolExecutionContext,
                                                    Optional returnMarkdown As Boolean = False) As System.Threading.Tasks.Task(Of String)
        If att Is Nothing Then Return Nothing
        If att.TempFilePath Is Nothing OrElse Not File.Exists(att.TempFilePath) Then Return Nothing

        Dim ext As String = Path.GetExtension(att.TempFilePath).ToLowerInvariant()
        Dim markdownCapable As Boolean = (ext = ".docx" OrElse ext = ".pdf")
        Dim useMarkdown As Boolean = returnMarkdown AndAlso markdownCapable

        ' Return cache if available
        If useMarkdown Then
            If att.CachedMarkdownText IsNot Nothing Then Return att.CachedMarkdownText
        ElseIf att.CachedText IsNot Nothing Then
            Return att.CachedText
        End If

        Dim text As String = Nothing
        Dim extracted As Boolean = False

        ' ── Sandboxed readers first (no COM interop) ──
        Try
            Select Case ext
                Case ".docx"
                    text = SharedMethods.ReadDocxSandboxed(att.TempFilePath, useMarkdown)
                    extracted = Not String.IsNullOrWhiteSpace(text) AndAlso Not text.StartsWith("Error")
                Case ".xlsx"
                    text = SharedMethods.ReadXlsxSandboxed(att.TempFilePath)
                    extracted = Not String.IsNullOrWhiteSpace(text) AndAlso Not text.StartsWith("Error")
                Case ".pptx"
                    text = SharedMethods.ReadPptxSandboxed(att.TempFilePath)
                    extracted = Not String.IsNullOrWhiteSpace(text) AndAlso Not text.StartsWith("Error")
                Case ".eml"
                    text = SharedMethods.ReadEmlSandboxed(att.TempFilePath)
                    extracted = Not String.IsNullOrWhiteSpace(text) AndAlso Not text.StartsWith("Error")
                Case ".msg"
                    text = ReadMsgAttachmentText(att.TempFilePath)
                    extracted = Not String.IsNullOrWhiteSpace(text) AndAlso Not text.StartsWith("Error")
                Case ".doc"
                    If Not INI_AllowLegacyDocFiles Then
                        text = "Error: .doc format disabled for security."
                        extracted = False
                    End If
                    ' Fall through to COM-based TryExtractOfficeText below
            End Select
        Catch
        End Try

        ' ── Fallback: Office COM interop for .doc, .xls, .ppt, .rtf (legacy formats) ──
        If Not extracted AndAlso (Not useMarkdown) AndAlso (ext <> ".doc" OrElse (ext = ".doc" AndAlso INI_AllowLegacyDocFiles)) Then
            Try
                Dim label As String = Nothing
                extracted = TryExtractOfficeText(att.TempFilePath, text, label)
            Catch
            End Try
        End If

        ' ── Fallback: text-like files ──
        If Not extracted AndAlso (Not useMarkdown) Then
            Try
                Dim label As String = Nothing
                extracted = TryExtractTextLike(att.TempFilePath, text, label)
            Catch
            End Try
        End If

        ' ── PDF extraction ──
        If Not extracted AndAlso ext = ".pdf" Then
            Try
                text = Await SharedMethods.ReadPdfAsText(
                    att.TempFilePath,
                    ReturnErrorInsteadOfEmpty:=True,
                    DoOCR:=False,
                    AskUser:=False,
                    ReturnMarkdown:=useMarkdown)
                extracted = Not String.IsNullOrWhiteSpace(text) AndAlso Not text.StartsWith("Error")
            Catch
            End Try
        End If

        If extracted AndAlso Not String.IsNullOrWhiteSpace(text) Then
            If useMarkdown Then
                att.CachedMarkdownText = text
            Else
                att.CachedText = text
            End If
            Return text
        End If

        Return Nothing
    End Function


    Private Async Function oldReadSingleAttachmentText(att As AutoPilotAttachmentInfo, context As ToolExecutionContext) As System.Threading.Tasks.Task(Of String)
        ' Return cache if available
        If att.CachedText IsNot Nothing Then Return att.CachedText

        If att.TempFilePath Is Nothing OrElse Not File.Exists(att.TempFilePath) Then Return Nothing

        Dim text As String = Nothing
        Dim label As String = Nothing
        Dim extracted As Boolean = False

        Try
            extracted = TryExtractOfficeText(att.TempFilePath, text, label)
        Catch
        End Try

        If Not extracted Then
            Dim ext = Path.GetExtension(att.TempFilePath).ToLowerInvariant()
            Try
                Select Case ext
                    Case ".xlsx", ".xls"
                        text = ExtractExcelText(att.TempFilePath)
                        extracted = Not String.IsNullOrWhiteSpace(text) AndAlso Not text.StartsWith("Error")
                    Case ".pptx"
                        text = ExtractPowerPointText(att.TempFilePath)
                        extracted = Not String.IsNullOrWhiteSpace(text) AndAlso Not text.StartsWith("Error")
                End Select
            Catch
            End Try
        End If

        If Not extracted Then
            Try
                extracted = TryExtractTextLike(att.TempFilePath, text, label)
            Catch
            End Try
        End If

        If Not extracted AndAlso att.Extension = ".pdf" Then
            Try
                text = Await SharedMethods.ReadPdfAsText(att.TempFilePath, ReturnErrorInsteadOfEmpty:=True, DoOCR:=False, AskUser:=False)
                extracted = Not String.IsNullOrWhiteSpace(text)
            Catch
            End Try
        End If

        If extracted AndAlso Not String.IsNullOrWhiteSpace(text) Then
            att.CachedText = text
            Return text
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Detects comment and tracked change counts in a .docx for hinting in read_attachment.
    ''' Result is cached in att.CachedDocxHint.
    ''' </summary>
    Private Function GetDocxMetadataHint(att As AutoPilotAttachmentInfo) As String
        If att.CachedDocxHint IsNot Nothing Then Return att.CachedDocxHint

        If att.Extension <> ".docx" OrElse att.TempFilePath Is Nothing OrElse Not File.Exists(att.TempFilePath) Then
            att.CachedDocxHint = ""
            Return ""
        End If

        Try
            Dim commentCount As Integer = 0
            Dim revisionCount As Integer = 0
            Dim tempDir = Path.Combine(Path.GetTempPath(), "ap_hint_" & Guid.NewGuid().ToString("N"))
            ZipFile.ExtractToDirectory(att.TempFilePath, tempDir)

            Dim commentsPath = Path.Combine(tempDir, "word", "comments.xml")
            If File.Exists(commentsPath) Then
                Dim commDoc As New XmlDocument()
                commDoc.Load(commentsPath)
                Dim nsMgr As New XmlNamespaceManager(commDoc.NameTable)
                nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                commentCount = commDoc.SelectNodes("//w:comment", nsMgr).Count
            End If

            Dim docPath = Path.Combine(tempDir, "word", "document.xml")
            If File.Exists(docPath) Then
                Dim docXml As New XmlDocument()
                docXml.Load(docPath)
                Dim nsMgr As New XmlNamespaceManager(docXml.NameTable)
                nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                revisionCount = docXml.SelectNodes("//w:ins", nsMgr).Count +
                                docXml.SelectNodes("//w:del", nsMgr).Count +
                                docXml.SelectNodes("//w:rPrChange", nsMgr).Count
            End If

            Try : Directory.Delete(tempDir, True) : Catch : End Try

            Dim hint As String = ""
            If commentCount > 0 OrElse revisionCount > 0 Then
                Dim parts As New List(Of String)()
                If commentCount > 0 Then parts.Add($"{commentCount} comment(s)")
                If revisionCount > 0 Then parts.Add($"{revisionCount} tracked change(s)")
                hint = $"(This document contains {String.Join(" and ", parts)}. Use {AP_Tool_ReadWordDocDetails} to inspect them.)"
            End If

            att.CachedDocxHint = hint
            Return hint
        Catch
            att.CachedDocxHint = ""
            Return ""
        End Try
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL IDENTIFICATION
    ' ═══════════════════════════════════════════════════════════════════════════

    Friend Function IsAutoPilotInternalTool(toolName As String) As Boolean
        Select Case toolName
            Case AP_Tool_ProcessWordDoc,
     AP_Tool_CommentWordDoc,
     AP_Tool_ExtractPdfText,
     AP_Tool_MergePdfs,
     AP_Tool_ReadAttachment,
     AP_Tool_ListAttachments,
     AP_Tool_DescribeBinary,
     AP_Tool_CompareWordDocs,
     AP_Tool_ReadWordDocDetails,
     AP_Tool_CreatePdfFromText,
     AP_Tool_ExtractExcelData,
     AP_Tool_ExcelListLiveWorksheets,
     AP_Tool_ExcelReadLiveRange,
     AP_Tool_ExcelCompleteLiveWorkbook,
     AP_Tool_SplitPdf,
     AP_Tool_AddPdfWatermark,
     AP_Tool_WordToPdf,
     AP_Tool_SearchInAttachments,
     AP_Tool_SummarizeThread,
     AP_Tool_PdfToWord,
     AP_Tool_CreateWordDoc,
     AP_Tool_CreateExcel,
     AP_Tool_CreatePowerPoint,
     AP_Tool_CreateCodeFile,
     AP_Tool_CommentPdf,
     AP_Tool_ExtractDataFromAttachments,
     AP_Tool_RedactPdf,
     AP_Tool_OverlayPdf,
     AP_Tool_CreateAudioFile,
     AP_Tool_GenerateImage,
     AP_Tool_WebGrounding,
     AP_Tool_ManageScheduledTasks,
     AP_Tool_ManageUserMemory,
     AP_Tool_ManageUserFiles,
     AP_Tool_ListCollectionUseCases,
     AP_Tool_CollectData,
     AP_Tool_PreviewCollection,
     AP_Tool_CompleteWordTables,
     AP_Tool_ReportInability
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Shared Function StripSelectorOwnedToolSuffixes(value As String) As String
        Dim result As String = If(value, "").Trim()

        If result.EndsWith(" (Outlook only)", StringComparison.OrdinalIgnoreCase) Then
            result = result.Substring(0, result.Length - " (Outlook only)".Length).TrimEnd()
        End If

        If result.EndsWith(" (Word only)", StringComparison.OrdinalIgnoreCase) Then
            result = result.Substring(0, result.Length - " (Word only)".Length).TrimEnd()
        End If

        If result.EndsWith(" (built-in)", StringComparison.OrdinalIgnoreCase) Then
            result = result.Substring(0, result.Length - " (built-in)".Length).TrimEnd()
        End If

        If result.EndsWith(" (internal)", StringComparison.OrdinalIgnoreCase) Then
            result = result.Substring(0, result.Length - " (internal)".Length).TrimEnd()
        End If

        Return result
    End Function

End Class
