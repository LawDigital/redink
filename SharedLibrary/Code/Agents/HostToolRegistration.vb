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
            TextTools.ToolExportToText,
            TextTools.ToolSemanticIndexCreateFromFile,
            TextTools.ToolSemanticIndexCreateFromText,
            TextTools.ToolSemanticIndexValidate,
            TextTools.ToolSemanticIndexSearch,
            TextTools.ToolSemanticIndexSearchContinuation,
            TextTools.ToolSemanticIndexLoadEntries,
            TextTools.ToolSemanticIndexVerifyAnswer,
            TextTools.ToolSemanticIndexRetrieveAfterVerification,
            TextTools.ToolSemanticIndexResetConversation,
            TextTools.ToolSemanticIndexInvalidateCache,
            FileTools.ToolCopy,
            FileTools.ToolList,
            FileTools.ToolMove,
            FileTools.ToolRename,
            FileTools.ToolDelete,
            FileTools.ToolMakeDir,
            FileTools.ToolRemoveDir,
            JsRunTool.ToolName,
            PythonExecuteTool.ToolName,
            SkillInvokeTool.ToolName,
            ContextExpandTool.ToolName,
            ContextCompactTool.ToolName,
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
            WordTools.ToolSaveAs,
            BrowserTools.BrowserOpenToolName,
            BrowserTools.BrowserSnapshotToolName,
            BrowserTools.BrowserInteractToolName
        }

        Private ReadOnly OutlookOnlyInternalToolNames As String() = New String() {
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

        ' LLM timeout load policy. This classifies MODEL-CALL planning/interpretation cost,
        ' not the local execution duration of the tool itself. Keep this explicit and host-agnostic.
        Private ReadOnly HeavyLlmToolNames As String() = New String() {
            "create_word_document",
            "complete_word_tables",
            "comment_word_document",
            "compare_word_documents",
            "read_word_document_details",
            "create_excel_spreadsheet",
            "extract_excel_data",
            "excel_list_live_worksheets",
            "excel_read_live_range",
            "excel_complete_live_workbook",
            "create_powerpoint",
            "create_pdf_from_text",
            "create_code_file",
            "create_audio_file",
            "comment_pdf_document",
            "redact_pdf",
            "extract_data_from_attachments",
            PythonExecuteTool.ToolName,
            JsRunTool.ToolName,
            "word_doc_read",
            "word_doc_edit",
            "word_doc_create"
        }

        Private ReadOnly HeavyLlmToolNameSet As HashSet(Of String) =
            BuildToolNameSet(HeavyLlmToolNames)

        ' A very large active tool/payload context is itself a model-call cost signal.
        ' This threshold is deliberately mechanical; it does not infer semantics from descriptions.
        Public Const HeavyLlmPayloadThresholdChars As Integer = 60000
        Public Const HeavyLlmTimeoutMultiplier As Integer = 2

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
            WorkspaceTools.ToolCopy,
            WorkspaceTools.ToolMove,
            WorkspaceTools.ToolRename,
            TextTools.ToolWrite,
            TextTools.ToolExportToText,
            FileTools.ToolCopy,
            FileTools.ToolMove,
            FileTools.ToolRename,
            PythonExecuteTool.ToolName,
            WordTools.ToolWrite,
            WordTools.ToolMarkup,
            WordTools.ToolCommentAdd,
            WordTools.ToolCommentRemove,
            WordTools.ToolFormat,
            WordTools.ToolApplyTemplate,
            WordTools.ToolSaveAs,
            "comment_word_document",
            "merge_pdfs",
            "compare_word_documents",
            "create_pdf_from_text",
            "excel_complete_live_workbook",
            "split_pdf",
            "add_pdf_watermark",
            "word_to_pdf",
            "pdf_to_word",
            "create_word_document",
            "complete_word_tables",
            "create_excel_spreadsheet",
            "create_powerpoint",
            "create_code_file",
            "comment_pdf_document",
            "redact_pdf",
            "overlay_pdf",
            "create_audio_file",
            "generate_image",
            "manage_user_files",
            "agent_workspace_write",
            "agent_workspace_file_op",
            "agent_workspace_save_session_file",
            "agent_workspace_move_to",
            "agent_workspace_copy_to",
            "agent_workspace_rename",
            "agent_workspace_bulk_rename",
            "agent_workspace_inventory_report"
        }

        Private ReadOnly WordDeliverableToolNames As String() = New String() {
            "download_web_files",
            WorkspaceTools.ToolWrite,
            WorkspaceTools.ToolCopy,
            WorkspaceTools.ToolMove,
            WorkspaceTools.ToolRename,
            TextTools.ToolWrite,
            TextTools.ToolExportToText,
            FileTools.ToolCopy,
            FileTools.ToolMove,
            FileTools.ToolRename,
            PythonExecuteTool.ToolName,
            WordTools.ToolWrite,
            WordTools.ToolMarkup,
            WordTools.ToolCommentAdd,
            WordTools.ToolCommentRemove,
            WordTools.ToolFormat,
            WordTools.ToolApplyTemplate,
            WordTools.ToolSaveAs,
            "create_word_document",
            "create_excel_spreadsheet",
            "create_powerpoint",
            "create_code_file",
            "create_pdf_from_text",
            "merge_pdfs",
            "add_pdf_watermark",
            "word_to_pdf",
            "pdf_to_word",
            "redact_pdf",
            "overlay_pdf",
            "create_audio_file",
            "generate_image",
            "word_doc_create",
            "word_doc_edit",
            "word_doc_export_pdf"
        }

        ' Additional host-specific/non-shared legacy tools may be registered explicitly
        ' by their owning source module. This keeps the compatibility boundary extensible
        ' without inferring deliverable capability from filenames, paths, descriptions,
        ' or tool-name patterns.
        Private ReadOnly DynamicOutlookDeliverableToolNames As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
        Private ReadOnly DynamicWordDeliverableToolNames As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
        Private ReadOnly DeliverableCapabilitySyncRoot As New System.Object()



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

                Dim normalizedToolName As String = tool.ToolName.Trim()
                ToolExecutorRegistry.RegisterInternal(host, normalizedToolName)

                ' artifact_generation is explicit model metadata, not a heuristic. It means
                ' the resolved tool can physically produce user-facing artifacts even if that
                ' tool still uses the bounded legacy delivery path rather than artifacts[].
                If HasCapabilityTag(tool.CapabilityTags, "artifact_generation") Then
                    RegisterDeliverableCapableToolName(host, normalizedToolName)
                End If
            Next
        End Sub

        Public Sub RegisterDeliverableCapableToolName(host As ToolingHostKind, toolName As String)
            Dim normalizedToolName As String = If(toolName, "").Trim()
            If normalizedToolName = "" Then Return

            SyncLock DeliverableCapabilitySyncRoot
                Select Case host
                    Case ToolingHostKind.Outlook
                        DynamicOutlookDeliverableToolNames.Add(normalizedToolName)
                    Case ToolingHostKind.Word
                        DynamicWordDeliverableToolNames.Add(normalizedToolName)
                End Select
            End SyncLock
        End Sub

        Public Sub RegisterDeliverableCapableToolNames(host As ToolingHostKind, toolNames As IEnumerable(Of String))
            If toolNames Is Nothing Then Return

            For Each toolName As String In toolNames
                RegisterDeliverableCapableToolName(host, toolName)
            Next
        End Sub

        Private Function HasCapabilityTag(capabilityTags As String, requiredTag As String) As Boolean
            Dim wanted As String = If(requiredTag, "").Trim()
            If wanted = "" OrElse String.IsNullOrWhiteSpace(capabilityTags) Then Return False

            For Each rawTag As String In capabilityTags.Split(New Char() {","c, ";"c}, System.StringSplitOptions.RemoveEmptyEntries)
                If System.String.Equals(rawTag.Trim(), wanted, System.StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Public Function IsInternalToolName(toolName As String) As Boolean
            Dim name As String = If(toolName, "").Trim()
            Return name <> "" AndAlso AllInternalToolNameSet.Contains(name)
        End Function

        ''' <summary>
        ''' Classifies a tool for the special <c>selected_online_sources</c> agent/skill alias.
        ''' The caller supplies the authoritative registry for the current run, which is already
        ''' narrowed to tools selected/authorized for that run (including explicit dependencies of
        ''' a selected skill). Keep this provider/jurisdiction agnostic:
        ''' selected external resource tools qualify automatically; only generic built-in retrieval
        ''' families qualify among host-internal tools. Mutation/deliverable tools never qualify.
        ''' </summary>
        Public Function IsSelectedOnlineSourceToolName(toolName As String) As Boolean
            Dim name As String = If(toolName, "").Trim()
            If name = "" Then Return False

            If name.StartsWith("skill_", System.StringComparison.OrdinalIgnoreCase) OrElse
               name.StartsWith("agent_", System.StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            Select Case name.ToLowerInvariant()
                Case "internet_search",
                     "web_grounding",
                     "retrieve_web_content",
                     "web_content_retriever",
                     "knowledge_search"
                    Return True
            End Select

            If SharedLibrary.M365ToolService.IsM365ToolName(name) Then Return True
            If BrowserTools.IsBrowserTool(name) Then Return True

            ' Every other built-in tool is an application/action capability, not a source.
            ' External/plugin/MCP tools remain eligible because this method is only called on
            ' the authoritative selected/authorized registry for the current run.
            If IsInternalToolName(name) Then Return False

            Return True
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

        Public Function IsHeavyLlmToolName(toolName As String) As Boolean
            Dim name As String = If(toolName, "").Trim()
            Return name <> "" AndAlso HeavyLlmToolNameSet.Contains(name)
        End Function

        Public Function GetLlmTimeoutMultiplier(toolNames As IEnumerable(Of String),
                                                toolInstructionsChars As Integer,
                                                toolResponsesChars As Integer) As Integer
            If toolNames IsNot Nothing Then
                For Each rawName As String In toolNames
                    If IsHeavyLlmToolName(rawName) Then Return HeavyLlmTimeoutMultiplier
                Next
            End If

            Dim combinedChars As Long = CLng(System.Math.Max(0, toolInstructionsChars)) +
                                        CLng(System.Math.Max(0, toolResponsesChars))
            If combinedChars >= HeavyLlmPayloadThresholdChars Then Return HeavyLlmTimeoutMultiplier
            Return 1
        End Function

        Public Function GetPerCallLlmTimeoutMs(configuredTimeoutMs As Long,
                                               toolNames As IEnumerable(Of String),
                                               toolInstructionsChars As Integer,
                                               toolResponsesChars As Integer) As Integer
            Dim baseMs As Long = configuredTimeoutMs
            If baseMs <= 0 Then baseMs = 30000
            Dim multiplier As Integer = GetLlmTimeoutMultiplier(toolNames, toolInstructionsChars, toolResponsesChars)
            Dim effectiveMs As Long = baseMs * CLng(multiplier)
            If effectiveMs > System.Int32.MaxValue Then effectiveMs = System.Int32.MaxValue
            Return CInt(System.Math.Max(1L, effectiveMs))
        End Function

        Public Function GetDeliverableCapableToolNames(host As ToolingHostKind) As IReadOnlyCollection(Of String)
            Return GetDeliverableCapableToolNames(host, Nothing)
        End Function

        Public Function GetDeliverableCapableToolNames(host As ToolingHostKind,
                                                       resolvedTools As IEnumerable(Of SharedLibrary.ModelConfig)) As IReadOnlyCollection(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            SyncLock DeliverableCapabilitySyncRoot
                Select Case host
                    Case ToolingHostKind.Word
                        For Each name As String In WordDeliverableToolNames
                            result.Add(name)
                        Next
                        result.UnionWith(DynamicWordDeliverableToolNames)

                    Case ToolingHostKind.Outlook
                        For Each name As String In OutlookDeliverableToolNames
                            result.Add(name)
                        Next
                        result.UnionWith(DynamicOutlookDeliverableToolNames)
                End Select
            End SyncLock

            ' Late-resolved/non-shared tools can declare artifact_generation directly in
            ' ModelConfig. This is deterministic explicit metadata and therefore safe to
            ' merge into the per-run compatibility allow-list without name heuristics.
            If resolvedTools IsNot Nothing Then
                For Each tool As SharedLibrary.ModelConfig In resolvedTools
                    If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For
                    If HasCapabilityTag(tool.CapabilityTags, "artifact_generation") Then
                        result.Add(tool.ToolName.Trim())
                    End If
                Next
            End If

            Return result.ToList().AsReadOnly()
        End Function

    End Module

End Namespace
