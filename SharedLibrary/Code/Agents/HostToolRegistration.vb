' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: HostToolRegistration.vb
' Purpose: Central authority for host-internal tool names, host registration,
'          deliverable-capable tool classification, and selector display suffixes.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Linq

Namespace Agents

    Public Module HostToolRegistration

        Private ReadOnly CommonInternalToolNames As String() = New String() {
            "retrieve_web_content",
            "web_content_retriever",
            "download_web_files",
            "internet_search",
            "web_grounding",
            "knowledge_search",
            ToolLoaderTool.LoaderToolName,
            MemoryTools.ToolPut,
            MemoryTools.ToolGet,
            MemoryTools.ToolList,
            MemoryTools.ToolDelete,
            TextTools.ToolRead,
            TextTools.ToolWrite,
            TextTools.ToolSearch,
            FileTools.ToolCopy,
            FileTools.ToolMove,
            FileTools.ToolRename,
            FileTools.ToolDelete,
            FileTools.ToolMakeDir,
            FileTools.ToolRemoveDir,
            JsRunTool.ToolName,
            PythonExecuteTool.ToolName,
            SkillInvokeTool.ToolName,
            SharedLibrary.M365ToolService.SearchToolName,
            SharedLibrary.M365ToolService.GetMailToolName,
            SharedLibrary.M365ToolService.GetMailThreadToolName,
            SharedLibrary.M365ToolService.GetFileToolName,
            SharedLibrary.M365ToolService.GetEventToolName,
            SharedLibrary.M365ToolService.GetChatThreadToolName,
            SharedLibrary.M365ToolService.GetOneNotePageToolName,
            WordTools.ToolExtract,
            WordTools.ToolSearch,
            WordTools.ToolWrite,
            WordTools.ToolMarkup,
            WordTools.ToolCommentAdd,
            WordTools.ToolCommentList,
            WordTools.ToolCommentRemove,
            WordTools.ToolFormat,
            WordTools.ToolApplyTemplate,
            WordTools.ToolSaveAs
        }

        Private ReadOnly OutlookOnlyInternalToolNames As String() = New String() {
            "process_word_document",
            "comment_word_document",
            "extract_pdf_text",
            "merge_pdfs",
            "read_attachment",
            "list_attachments",
            "describe_binary_attachment",
            "compare_word_documents",
            "read_word_document_details",
            "create_pdf_from_text",
            "extract_excel_data",
            "excel_list_live_worksheets",
            "excel_read_live_range",
            "excel_complete_live_workbook",
            "split_pdf",
            "add_pdf_watermark",
            "word_to_pdf",
            "search_in_attachments",
            "summarize_thread",
            "pdf_to_word",
            "create_word_document",
            "complete_word_tables",
            "create_excel_spreadsheet",
            "create_powerpoint",
            "create_code_file",
            "comment_pdf_document",
            "extract_data_from_attachments",
            "redact_pdf",
            "overlay_pdf",
            "create_audio_file",
            "generate_image",
            "manage_scheduled_tasks",
            "manage_user_memory",
            "manage_user_files",
            "report_inability",
            "agent_workspace_list",
            "agent_workspace_read",
            "agent_workspace_write",
            "agent_workspace_file_op",
            "agent_workspace_save_session_file",
            "agent_workspace_search",
            "agent_workspace_find_files",
            "agent_workspace_move_to",
            "agent_workspace_copy_to",
            "agent_workspace_rename",
            "agent_workspace_bulk_rename",
            "agent_workspace_file_details",
            "agent_workspace_recent_files",
            "agent_workspace_create_folder_structure",
            "agent_workspace_trash",
            "agent_workspace_inventory_report"
        }

        Private ReadOnly WordOnlyInternalToolNames As String() = New String() {
            WorkspaceTools.ToolGet,
            WorkspaceTools.ToolInventory,
            WorkspaceTools.ToolRead,
            WorkspaceTools.ToolReadMany,
            WorkspaceTools.ToolWrite,
            WorkspaceTools.ToolSearch,
            WorkspaceTools.ToolCopy,
            WorkspaceTools.ToolMove,
            WorkspaceTools.ToolRename,
            WorkspaceTools.ToolDelete,
            WorkspaceTools.ToolMakeDir,
            WorkspaceTools.ToolExtractText,
            WorkspaceTools.ToolExtractTextMany,
            WordDocTools.ToolListOpen,
            WordDocTools.ToolGetActive,
            WordDocTools.ToolExtract,
            WordDocTools.ToolSearch,
            WordDocTools.ToolListComments,
            WordDocTools.ToolInsert,
            WordDocTools.ToolReplace,
            WordDocTools.ToolDelete,
            WordDocTools.ToolCommentAdd,
            WordDocTools.ToolFormat,
            "word_doc_read",
            "word_doc_edit",
            "word_doc_create",
            "word_doc_export_pdf"
        }

        Private ReadOnly OutlookDeliverableToolNames As String() = New String() {
            "download_web_files",
            WorkspaceTools.ToolWrite,
            TextTools.ToolWrite,
            FileTools.ToolCopy,
            PythonExecuteTool.ToolName,
            "merge_pdfs",
            "create_pdf_from_text",
            "split_pdf",
            "add_pdf_watermark",
            "word_to_pdf",
            "pdf_to_word",
            "create_word_document",
            "complete_word_tables",
            "create_excel_spreadsheet",
            "create_powerpoint",
            "create_code_file",
            "redact_pdf",
            "overlay_pdf",
            "create_audio_file",
            "generate_image"
        }

        Private ReadOnly WordDeliverableToolNames As String() = New String() {
            "download_web_files",
            WorkspaceTools.ToolWrite,
            TextTools.ToolWrite,
            FileTools.ToolCopy,
            PythonExecuteTool.ToolName,
            WordTools.ToolWrite,
            WordTools.ToolMarkup,
            WordTools.ToolApplyTemplate,
            WordTools.ToolSaveAs,
            "word_doc_create",
            "word_doc_edit",
            "word_doc_export_pdf"
        }

        Private ReadOnly CommonInternalToolNameSet As HashSet(Of String) =
            BuildToolNameSet(CommonInternalToolNames)

        Private ReadOnly OutlookOnlyInternalToolNameSet As HashSet(Of String) =
            BuildToolNameSet(OutlookOnlyInternalToolNames)

        Private ReadOnly WordOnlyInternalToolNameSet As HashSet(Of String) =
            BuildToolNameSet(WordOnlyInternalToolNames)

        Private ReadOnly AllInternalToolNameSet As HashSet(Of String) =
            BuildAllInternalToolNameSet()

        Private Function BuildToolNameSet(names As IEnumerable(Of String)) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If names Is Nothing Then
                Return result
            End If

            For Each rawName As String In names
                Dim name As String = If(rawName, "").Trim()
                If name <> "" Then
                    result.Add(name)
                End If
            Next

            Return result
        End Function

        Private Function BuildAllInternalToolNameSet() As HashSet(Of String)
            Dim result As New HashSet(Of String)(CommonInternalToolNameSet, StringComparer.OrdinalIgnoreCase)
            result.UnionWith(OutlookOnlyInternalToolNameSet)
            result.UnionWith(WordOnlyInternalToolNameSet)
            Return result
        End Function

        Private Sub RegisterSet(host As ToolingHostKind, names As IEnumerable(Of String))
            If names Is Nothing Then Return

            For Each rawName As String In names
                Dim name As String = If(rawName, "").Trim()
                If name = "" Then Continue For
                ToolExecutorRegistry.RegisterInternal(host, name)
            Next
        End Sub

        Public Sub RegisterAll(host As ToolingHostKind)
            Select Case host
                Case ToolingHostKind.Outlook
                    RegisterOutlookInternals()
                Case ToolingHostKind.Word
                    RegisterWordInternals()
            End Select
        End Sub

        Public Sub RegisterOutlookInternals()
            Dim host As ToolingHostKind = ToolingHostKind.Outlook
            ToolExecutorRegistry.Reset(host)
            RegisterSet(host, CommonInternalToolNameSet)
            RegisterSet(host, OutlookOnlyInternalToolNameSet)
        End Sub

        Public Sub RegisterWordInternals()
            Dim host As ToolingHostKind = ToolingHostKind.Word
            ToolExecutorRegistry.Reset(host)
            RegisterSet(host, CommonInternalToolNameSet)
            RegisterSet(host, WordOnlyInternalToolNameSet)
        End Sub

        Public Sub RegisterResolvedInternalTools(host As ToolingHostKind, tools As IEnumerable(Of SharedLibrary.ModelConfig))
            If tools Is Nothing Then Return

            For Each tool As SharedLibrary.ModelConfig In tools
                If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For
                ToolExecutorRegistry.RegisterInternal(host, tool.ToolName.Trim())
            Next
        End Sub

        Public Function IsInternalToolName(toolName As String) As Boolean
            Dim name As String = If(toolName, "").Trim()
            Return name <> "" AndAlso AllInternalToolNameSet.Contains(name)
        End Function

        Public Function IsSharedInternalToolName(toolName As String) As Boolean
            Dim name As String = If(toolName, "").Trim()
            Return name <> "" AndAlso CommonInternalToolNameSet.Contains(name)
        End Function

        Public Function IsOutlookOnlyInternalToolName(toolName As String) As Boolean
            Dim name As String = If(toolName, "").Trim()
            Return name <> "" AndAlso OutlookOnlyInternalToolNameSet.Contains(name)
        End Function

        Public Function IsWordOnlyInternalToolName(toolName As String) As Boolean
            Dim name As String = If(toolName, "").Trim()
            Return name <> "" AndAlso WordOnlyInternalToolNameSet.Contains(name)
        End Function

        Public Function GetSelectorDisplaySuffix(toolName As String) As String
            Dim name As String = If(toolName, "").Trim()

            If name = "" OrElse Not AllInternalToolNameSet.Contains(name) Then
                Return ""
            End If

            If OutlookOnlyInternalToolNameSet.Contains(name) Then
                Return " (built-in) (Outlook only)"
            End If

            If WordOnlyInternalToolNameSet.Contains(name) Then
                Return " (built-in) (Word only)"
            End If

            Return " (built-in)"
        End Function

        Public Function GetDeliverableCapableToolNames(host As ToolingHostKind) As IReadOnlyCollection(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Select Case host
                Case ToolingHostKind.Word
                    For Each name As String In WordDeliverableToolNames
                        result.Add(name)
                    Next
                Case ToolingHostKind.Outlook
                    For Each name As String In OutlookDeliverableToolNames
                        result.Add(name)
                    Next
            End Select

            Return result.ToList().AsReadOnly()
        End Function

    End Module

End Namespace
