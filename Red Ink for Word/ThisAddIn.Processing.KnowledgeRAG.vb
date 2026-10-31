' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.KnowledgeRAG.vb
' Purpose: Queries the configured knowledge stores and injects context into
'          the system prompt for non-tooling Freestyle (kb) trigger processing.
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
    ''' Uses the KnowledgeTriggerHelper to resolve the query against the merged index.
    ''' </summary>
    ''' <param name="userQuery">The user's prompt text (after removing the (kb) trigger).</param>
    ''' <returns>
    ''' A tuple: (Success, KnowledgeContext). On success, KnowledgeContext contains
    ''' the assembled knowledge text. On failure, it contains an error message.
    ''' </returns>
    Public Function ConsultKnowledgeStore(
        userQuery As String
    ) As (Success As Boolean, KnowledgeContext As String)

        Try
            ' Check configuration
            If Not KnowledgeStoreManager.IsConfigured(_context) Then
                Return (False, "No knowledge store paths are configured. " &
                        "Set 'KnowledgeStorePath' and/or 'KnowledgeStorePathLocal' in your configuration file.")
            End If

            InfoBox.ShowInfoBox("Searching the knowledge store...")

            ' Build a KnowledgeRequest from the user query
            Dim request As New KnowledgeTriggerHelper.KnowledgeRequest() With {
                .LoadAll = False,
                .SearchQuery = userQuery,
                .Tags = Nothing,
                .RawTrigger = ""
            }

            ' If the query is empty or very short, load all documents
            If String.IsNullOrWhiteSpace(userQuery) OrElse userQuery.Trim().Length < 3 Then
                request.LoadAll = True
            End If

            ' Resolve the knowledge request
            Dim result = KnowledgeTriggerHelper.ResolveKnowledge(request, _context)

            InfoBox.ShowInfoBox("")

            If String.IsNullOrWhiteSpace(result.Content) Then
                Dim msg = If(String.IsNullOrWhiteSpace(result.StatusMessage),
                             "No relevant documents found in the knowledge store.",
                             result.StatusMessage)
                Return (False, msg)
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