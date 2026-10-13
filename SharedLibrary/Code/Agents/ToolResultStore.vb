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
        Private Shared ReadOnly _contextRecoveryRequests As New System.Collections.Concurrent.ConcurrentDictionary(Of System.String, System.Boolean)(System.StringComparer.Ordinal)

        Public NotInheritable Class StoredResult
            Public Property Ref As String
            Public Property WorkflowId As String
            Public Property ToolName As String
            Public Property FullContent As String
            Public Property TotalChars As Integer
            Public Property CreatedUtc As DateTime
            Public Property ExpiresUtc As System.DateTime
            Public Property ContentHash As System.String
            Public Property IsDurable As System.Boolean
        End Class

        ''' <summary>Stores a full body and returns a short reference token.</summary>
        Public Shared Function Put(workflowId As String, toolName As String, fullContent As String) As StoredResult
            Dim body As String = If(fullContent, "")
            Dim ref As System.String = "tref_" & System.Guid.NewGuid().ToString("N")

            Dim stored As New StoredResult() With {
                .Ref = ref,
                .WorkflowId = If(workflowId, ""),
                .ToolName = If(toolName, ""),
                .FullContent = body,
                .TotalChars = body.Length,
                .CreatedUtc = DateTime.UtcNow
            }

            stored.ExpiresUtc = System.DateTime.UtcNow.AddDays(WorkflowContinuity.DefaultContinuationRetentionDays)
            stored.ContentHash = ComputeContentHash(body)
            stored.IsDurable = TryPersist(stored)
            _entries(ref) = stored
            Return stored
        End Function

        Public Shared Function TryGet(ref As String, ByRef stored As StoredResult) As Boolean
            stored = Nothing
            If String.IsNullOrWhiteSpace(ref) Then Return False
            If _entries.TryGetValue(ref.Trim(), stored) Then Return True
            Return TryRestore(ref.Trim(), WorkflowContinuity.CurrentWorkflowId, stored)
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
            If System.String.IsNullOrWhiteSpace(ref) Then Return False
            If Not _entries.TryGetValue(ref.Trim(), candidate) AndAlso
               Not TryRestore(ref.Trim(), workflowId, candidate) Then Return False
            If candidate Is Nothing Then Return False

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

        ' Snapshots are private to the Windows user, immutable, workflow-scoped and
        ' integrity checked. Failure retains the full active body; callers must not shelve it.
        Private Shared Function ComputeContentHash(body As System.String) As System.String
            Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Return System.Convert.ToBase64String(algorithm.ComputeHash(System.Text.Encoding.UTF8.GetBytes(body)))
            End Using
        End Function

        Private Shared Function SnapshotPath(workflowId As System.String, resultRef As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(workflowId) OrElse
               Not System.Text.RegularExpressions.Regex.IsMatch(If(resultRef, ""), "^tref_[a-f0-9]{32}$") Then Return ""
            Dim checkpoint As System.String = WorkflowContinuity.GetCheckpointPath(workflowId)
            Return System.IO.Path.Combine(System.IO.Path.GetDirectoryName(checkpoint),
                                          System.IO.Path.GetFileNameWithoutExtension(checkpoint) & "_artifacts",
                                          "tool_results", resultRef & ".bin")
        End Function

        Private Shared Function TryPersist(stored As StoredResult) As System.Boolean
            Dim temporaryPath As System.String = ""
            Try
                Dim target As System.String = SnapshotPath(stored.WorkflowId, stored.Ref)
                If target = "" Then Return False
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(target))
                Dim json As System.String = Newtonsoft.Json.JsonConvert.SerializeObject(stored)
                Dim encrypted As System.Byte() = System.Security.Cryptography.ProtectedData.Protect(
                    System.Text.Encoding.UTF8.GetBytes(json), Nothing,
                    System.Security.Cryptography.DataProtectionScope.CurrentUser)
                temporaryPath = target & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
                System.IO.File.WriteAllBytes(temporaryPath, encrypted)
                System.IO.File.Move(temporaryPath, target)
                System.IO.Directory.SetLastWriteTimeUtc(
                    System.IO.Directory.GetParent(System.IO.Path.GetDirectoryName(target)).FullName, System.DateTime.UtcNow)
                Dim verified As StoredResult = Nothing
                Return TryRestore(stored.Ref, stored.WorkflowId, verified) AndAlso
                       System.String.Equals(verified.FullContent, stored.FullContent, System.StringComparison.Ordinal)
            Catch ex As System.Exception
                System.Diagnostics.Trace.TraceWarning("Tool-result snapshot failed; full content retained: " & ex.Message)
                Return False
            Finally
                If temporaryPath <> "" Then
                    Try
                        If System.IO.File.Exists(temporaryPath) Then System.IO.File.Delete(temporaryPath)
                    Catch ex As System.Exception
                        System.Diagnostics.Trace.TraceWarning("Tool-result temporary cleanup failed: " & ex.Message)
                    End Try
                End If
            End Try
        End Function

        Private Shared Function TryRestore(resultRef As System.String, workflowId As System.String,
                                           ByRef stored As StoredResult) As System.Boolean
            stored = Nothing
            Try
                Dim target As System.String = SnapshotPath(workflowId, resultRef)
                If target = "" OrElse Not System.IO.File.Exists(target) Then Return False
                Dim bytes As System.Byte() = System.Security.Cryptography.ProtectedData.Unprotect(
                    System.IO.File.ReadAllBytes(target), Nothing,
                    System.Security.Cryptography.DataProtectionScope.CurrentUser)
                Dim candidate As StoredResult = Newtonsoft.Json.JsonConvert.DeserializeObject(Of StoredResult)(
                    System.Text.Encoding.UTF8.GetString(bytes))
                If candidate Is Nothing OrElse WorkflowContinuity.GetToolResultExpiryUtc(workflowId, candidate.ExpiresUtc) <= System.DateTime.UtcNow OrElse
                   Not System.String.Equals(candidate.Ref, resultRef, System.StringComparison.Ordinal) OrElse
                   Not System.String.Equals(candidate.WorkflowId, workflowId, System.StringComparison.Ordinal) OrElse
                   candidate.FullContent Is Nothing OrElse candidate.TotalChars <> candidate.FullContent.Length OrElse
                   Not System.String.Equals(candidate.ContentHash, ComputeContentHash(candidate.FullContent), System.StringComparison.Ordinal) Then Return False
                candidate.ExpiresUtc = WorkflowContinuity.GetToolResultExpiryUtc(workflowId, candidate.ExpiresUtc)
                candidate.IsDurable = True
                stored = candidate
                Return True
            Catch ex As System.Exception
                System.Diagnostics.Trace.TraceWarning("Tool-result snapshot unavailable; no evidence reconstructed: " & ex.Message)
                Return False
            End Try
        End Function

        Public Shared Function DescribeReference(stored As StoredResult) As Newtonsoft.Json.Linq.JObject
            Return New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("result_ref", stored.Ref),
                New Newtonsoft.Json.Linq.JProperty("content_complete", False),
                New Newtonsoft.Json.Linq.JProperty("representation", "reference_stub"),
                New Newtonsoft.Json.Linq.JProperty("is_semantic_summary", False),
                New Newtonsoft.Json.Linq.JProperty("full_content_available", True),
                New Newtonsoft.Json.Linq.JProperty("durable_snapshot", stored.IsDurable),
                New Newtonsoft.Json.Linq.JProperty("sha256", stored.ContentHash),
                New Newtonsoft.Json.Linq.JProperty("hash_scope", "stored_tool_result"),
                New Newtonsoft.Json.Linq.JProperty("character_unit", "utf16_code_units"),
                New Newtonsoft.Json.Linq.JProperty("expires_utc", stored.ExpiresUtc),
                New Newtonsoft.Json.Linq.JProperty("total_chars", stored.TotalChars),
                New Newtonsoft.Json.Linq.JProperty("retrieval_required", "Incomplete reference. You MUST retrieve every task-required range before using omitted evidence; whole-source tasks require complete coverage. Use context_expand, or an exposed isolated analyzer accepting input_result_ref to read the complete stored text without filling parent context. If the stored text contains nested previews/references, follow them: reading a wrapper is not reading its sources. Never infer unavailable content."))
        End Function

        Public Shared Function AnnotateReferenceEnvelope(envelope As Newtonsoft.Json.Linq.JObject) As Newtonsoft.Json.Linq.JObject
            Dim stored As StoredResult = Nothing
            If TryGetForWorkflow(envelope.Value(Of System.String)("result_ref"), WorkflowContinuity.CurrentWorkflowId, stored) Then
                Dim metadata As Newtonsoft.Json.Linq.JObject = DescribeReference(stored)
                For Each prop As Newtonsoft.Json.Linq.JProperty In metadata.Properties()
                    envelope(prop.Name) = prop.Value.DeepClone()
                Next
            End If
            Return envelope
        End Function

        Public Shared Function GetUnicodeSafePrefixLength(body As System.String, maxChars As System.Int32) As System.Int32
            Dim source As System.String = If(body, System.String.Empty)
            Dim length As System.Int32 = System.Math.Min(System.Math.Max(0, maxChars), source.Length)
            If length > 0 AndAlso length < source.Length AndAlso
               System.Char.IsHighSurrogate(source(length - 1)) AndAlso System.Char.IsLowSurrogate(source(length)) Then
                length -= 1
            End If
            Return length
        End Function

        Public Shared Function BuildSourcePreview(workflowId As System.String, toolName As System.String,
                                                  fullContent As System.String, maxChars As System.Int32,
                                                  sourcePath As System.String) As Newtonsoft.Json.Linq.JObject
            Dim body As System.String = If(fullContent, "")
            Dim stored As StoredResult = Put(workflowId, toolName, body)
            Dim returnedChars As System.Int32 = If(stored.IsDurable, GetUnicodeSafePrefixLength(body, maxChars), body.Length)
            Dim result As Newtonsoft.Json.Linq.JObject = DescribeReference(stored)
            result("path") = sourcePath
            result("tool") = toolName
            result("text") = body.Substring(0, returnedChars)
            result("returned_chars") = returnedChars
            result("next_offset") = returnedChars
            result("truncated") = returnedChars < body.Length
            result("content_complete") = returnedChars = body.Length
            result("representation") = If(returnedChars < body.Length, "source_preview", "full_extraction")
            Return result
        End Function

        Public Shared Function ExportReferences(workflowId As System.String) As System.Collections.Generic.List(Of System.String)
            Dim refs As New System.Collections.Generic.List(Of System.String)()
            For Each entry In _entries.Values
                If entry.IsDurable AndAlso System.String.Equals(entry.WorkflowId, workflowId, System.StringComparison.Ordinal) Then refs.Add(entry.Ref)
            Next
            Return refs
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
        Public Shared Sub RequestContextRecovery(workflowId As System.String)
            If Not System.String.IsNullOrWhiteSpace(workflowId) Then _contextRecoveryRequests(workflowId) = True
        End Sub

        Public Shared Function IsContextRecoveryRequested(workflowId As System.String) As System.Boolean
            Return Not System.String.IsNullOrWhiteSpace(workflowId) AndAlso _contextRecoveryRequests.ContainsKey(workflowId)
        End Function

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

        ''' <summary>Releases durable in-memory bodies; failed snapshots remain available in memory for recovery.</summary>
        Public Shared Sub ClearWorkflow(workflowId As String)
            If String.IsNullOrWhiteSpace(workflowId) Then Return
            For Each kvp In _entries.ToArray()
                If String.Equals(kvp.Value.WorkflowId, workflowId, StringComparison.Ordinal) AndAlso kvp.Value.IsDurable Then
                    Dim verified As StoredResult = Nothing
                    If TryRestore(kvp.Key, workflowId, verified) Then
                        Dim removed As StoredResult = Nothing
                        _entries.TryRemove(kvp.Key, removed)
                    Else
                        kvp.Value.IsDurable = False
                        System.Diagnostics.Trace.TraceWarning("Snapshot verification failed during cleanup; full result retained in memory: " & kvp.Key)
                    End If
                End If
            Next

            Dim removedKeep As Integer
            _compactionRequests.TryRemove(workflowId, removedKeep)
            Dim removedRecovery As System.Boolean
            _contextRecoveryRequests.TryRemove(workflowId, removedRecovery)
        End Sub

    End Class

End Namespace
