' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolResultStore.vb
' Purpose: Workflow-scoped store for full tool result bodies. Large results are
'          replaced in model replay by a short reference; the full text remains
'          retrievable on demand via context_expand. Shared by Outlook and Word
'          so both hosts behave identically.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Collections.Concurrent

Namespace Agents

    Public NotInheritable Class ToolResultStore

        Private Sub New()
        End Sub

        Private Shared ReadOnly _entries As New ConcurrentDictionary(Of String, StoredResult)(StringComparer.OrdinalIgnoreCase)
        Private Shared ReadOnly _compactionRequests As New ConcurrentDictionary(Of String, Integer)(StringComparer.Ordinal)
        Private Shared _counter As Integer = 0

        Public NotInheritable Class StoredResult
            Public Property Ref As String
            Public Property WorkflowId As String
            Public Property ToolName As String
            Public Property FullContent As String
            Public Property TotalChars As Integer
            Public Property CreatedUtc As DateTime
        End Class

        ''' <summary>Stores a full body and returns a short reference token.</summary>
        Public Shared Function Put(workflowId As String, toolName As String, fullContent As String) As StoredResult
            Dim body As String = If(fullContent, "")
            Dim seq As Integer = System.Threading.Interlocked.Increment(_counter)
            Dim ref As String = "tref_" & seq.ToString("D6")

            Dim stored As New StoredResult() With {
                .Ref = ref,
                .WorkflowId = If(workflowId, ""),
                .ToolName = If(toolName, ""),
                .FullContent = body,
                .TotalChars = body.Length,
                .CreatedUtc = DateTime.UtcNow
            }

            _entries(ref) = stored
            Return stored
        End Function

        Public Shared Function TryGet(ref As String, ByRef stored As StoredResult) As Boolean
            stored = Nothing
            If String.IsNullOrWhiteSpace(ref) Then Return False
            Return _entries.TryGetValue(ref.Trim(), stored)
        End Function

        ''' <summary>
        ''' Resolves a stored result while enforcing workflow ownership whenever the
        ''' caller is workflow-scoped. Empty caller workflow ids retain the legacy
        ''' unscoped behavior for non-tooling callers.
        ''' </summary>
        Public Shared Function TryGetForWorkflow(ref As System.String,
                                                workflowId As System.String,
                                                ByRef stored As StoredResult) As System.Boolean
            stored = Nothing

            Dim candidate As StoredResult = Nothing
            If Not TryGet(ref, candidate) OrElse candidate Is Nothing Then Return False

            Dim requestedWorkflowId As System.String = If(workflowId, System.String.Empty).Trim()
            Dim storedWorkflowId As System.String = If(candidate.WorkflowId, System.String.Empty).Trim()

            If requestedWorkflowId <> System.String.Empty Then
                If storedWorkflowId = System.String.Empty OrElse
                   Not System.String.Equals(requestedWorkflowId, storedWorkflowId, System.StringComparison.Ordinal) Then
                    Return False
                End If
            End If

            stored = candidate
            Return True
        End Function

        ''' <summary>
        ''' Reuses an existing response-owned reference only when it still points to the
        ''' exact immutable result body for the same workflow and producing tool. This is
        ''' deliberately not global/content deduplication.
        ''' </summary>
        Public Shared Function TryReuseReference(ref As System.String,
                                                 workflowId As System.String,
                                                 toolName As System.String,
                                                 fullContent As System.String,
                                                 ByRef stored As StoredResult) As System.Boolean
            stored = Nothing

            Dim candidate As StoredResult = Nothing
            If Not TryGetForWorkflow(ref, workflowId, candidate) OrElse candidate Is Nothing Then Return False

            If Not System.String.Equals(
                If(candidate.ToolName, System.String.Empty),
                If(toolName, System.String.Empty),
                System.StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            Dim body As System.String = If(fullContent, System.String.Empty)
            If candidate.TotalChars <> body.Length Then Return False

            Dim storedBody As System.String = If(candidate.FullContent, System.String.Empty)
            If Not System.Object.ReferenceEquals(storedBody, body) AndAlso
               Not System.String.Equals(storedBody, body, System.StringComparison.Ordinal) Then
                Return False
            End If

            stored = candidate
            Return True
        End Function

        ''' <summary>Returns a character window from a stored body.</summary>
        Public Shared Function GetWindow(ref As String, startChar As Integer, maxChars As Integer) As String
            Dim stored As StoredResult = Nothing
            If Not TryGet(ref, stored) Then Return ""

            Dim body As System.String = If(stored.FullContent, System.String.Empty)
            Dim start As System.Int32 = System.Math.Max(0, System.Math.Min(startChar, body.Length))
            Dim remaining As System.Int32 = body.Length - start
            If remaining = 0 Then Return System.String.Empty

            ' Preserve the established minimum of one character only while content remains.
            Dim take As System.Int32 = System.Math.Min(System.Math.Max(1, maxChars), remaining)
            Return body.Substring(start, take)
        End Function

        ''' <summary>
        ''' Records a model-requested compaction for a workflow. Only tightens (keeps the
        ''' smallest requested recent-full count). Honoured by the host on the next rebuild.
        ''' </summary>
        Public Shared Sub RequestCompaction(workflowId As String, keepRecentFullCount As Integer)
            If String.IsNullOrWhiteSpace(workflowId) Then Return
            Dim keep As Integer = Math.Max(0, keepRecentFullCount)
            _compactionRequests.AddOrUpdate(workflowId, keep, Function(k, existing) Math.Min(existing, keep))
        End Sub

        ''' <summary>Returns True and the requested recent-full count when the model asked to compact this workflow.</summary>
        Public Shared Function TryGetRequestedKeepRecent(workflowId As String, ByRef keepRecentFullCount As Integer) As Boolean
            keepRecentFullCount = 0
            If String.IsNullOrWhiteSpace(workflowId) Then Return False
            Return _compactionRequests.TryGetValue(workflowId, keepRecentFullCount)
        End Function

        ''' <summary>Clears all bodies for a workflow (call at end of a run).</summary>
        Public Shared Sub ClearWorkflow(workflowId As String)
            If String.IsNullOrWhiteSpace(workflowId) Then Return
            For Each kvp In _entries.ToArray()
                If String.Equals(kvp.Value.WorkflowId, workflowId, StringComparison.Ordinal) Then
                    Dim removed As StoredResult = Nothing
                    _entries.TryRemove(kvp.Key, removed)
                End If
            Next

            Dim removedKeep As Integer
            _compactionRequests.TryRemove(workflowId, removedKeep)
        End Sub

    End Class

End Namespace
