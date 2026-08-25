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

        ''' <summary>Returns a character window from a stored body.</summary>
        Public Shared Function GetWindow(ref As String, startChar As Integer, maxChars As Integer) As String
            Dim stored As StoredResult = Nothing
            If Not TryGet(ref, stored) Then Return ""

            Dim body As String = If(stored.FullContent, "")
            Dim start As Integer = Math.Max(0, Math.Min(startChar, body.Length))
            Dim take As Integer = Math.Max(1, Math.Min(maxChars, body.Length - start))
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
