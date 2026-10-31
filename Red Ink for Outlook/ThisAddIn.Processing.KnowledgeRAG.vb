' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.KnowledgeRAG.vb
' Purpose: Queries the configured knowledge stores and injects context into
'          the system prompt for non-tooling Freestyle (kb) trigger processing.
' Note:    Mirrors the Word add-in's KnowledgeRAG logic.
'
' KEY ARCHITECTURE
'   • Surface: ThisAddIn (host partial/UI).
'   • Entry points/helpers: ConsultKnowledgeStore.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Queries the knowledge store for documents relevant to the user's prompt
    ''' and returns combined document text as tagged blocks for system prompt injection.
    ''' </summary>
    Public Function ConsultKnowledgeStore(
        userQuery As String
    ) As (Success As Boolean, KnowledgeContext As String)

        Try
            If Not KnowledgeStoreManager.IsConfigured(_context) Then
                Return (False, "No knowledge store paths are configured. " &
                        "Set 'KnowledgeStorePath' and/or 'KnowledgeStorePathLocal' in your configuration file.")
            End If

            InfoBox.ShowInfoBox("Searching the knowledge store...")

            Dim request As New KnowledgeTriggerHelper.KnowledgeRequest() With {
                .LoadAll = False,
                .SearchQuery = userQuery,
                .Tags = Nothing,
                .RawTrigger = ""
            }

            If String.IsNullOrWhiteSpace(userQuery) OrElse userQuery.Trim().Length < 3 Then
                request.LoadAll = True
            End If

            Dim result = KnowledgeTriggerHelper.ResolveKnowledge(request, _context)

            InfoBox.ShowInfoBox("")

            If String.IsNullOrWhiteSpace(result.Content) Then
                Return (False, If(String.IsNullOrWhiteSpace(result.StatusMessage),
                                  "No relevant documents found.", result.StatusMessage))
            End If

            If Not String.IsNullOrWhiteSpace(result.StatusMessage) Then
                InfoBox.ShowInfoBox(result.StatusMessage, 3)
            End If

            Return (True, result.Content)

        Catch ex As Exception
            InfoBox.ShowInfoBox("")
            Return (False, $"Knowledge store query failed: {ex.Message}")
        End Try

    End Function

End Class