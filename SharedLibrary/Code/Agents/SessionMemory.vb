' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SessionMemory.vb
' Purpose: Process-wide session-scoped key/value memory that the agent can write
'          to and read from. Persists to disk under the user-local agent resources
'          path so it survives Outlook/Word restarts until the user clears it.
'
' Storage Layout:
'  <INI_AgentResourcesPathLocal>\.session\memory.json
'  (Falls back to %LOCALAPPDATA%\RedInk\.session\memory.json if unset.)
'
' Entry Structure:
'  - key: opaque, stable identifier (assistant-chosen or auto-generated).
'  - summary: short textual stub (model sees this in subsequent turns).
'  - value: actual payload (any JSON value or string).
'  - createdAt, updatedAt, tags, metadata: operational fields.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public Class SessionMemoryEntry
        Public Property Key As String
        Public Property Summary As String
        Public Property Value As JToken
        Public Property CreatedAt As DateTime
        Public Property UpdatedAt As DateTime
        Public Property Tags As List(Of String)
        Public Property Metadata As SessionMemoryMetadata
    End Class

    Public NotInheritable Class SessionMemory

        Private Sub New()
        End Sub

        Private Shared ReadOnly _sync As New Object()
        Private Shared _loaded As Boolean
        Private Shared _entries As New Dictionary(Of String, SessionMemoryEntry)(StringComparer.Ordinal)

        ' ------------------------------------------------------------------ persistence

        Private Shared Function GetStorePath() As String
            Dim root As String = TryGetSharedPath("INI_AgentResourcesPathLocal")
            If Not System.String.IsNullOrWhiteSpace(root) Then
                root = SharedLibrary.SharedMethods.ExpandEnvironmentVariables(root)
            End If
            If System.String.IsNullOrWhiteSpace(root) Then
                root = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "RedInk")
            End If
            Dim dir = System.IO.Path.Combine(root, ".session")
            Try
                If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
            Catch
            End Try
            Return Path.Combine(dir, "memory.json")
        End Function

        Private Shared Sub EnsureLoaded()
            If _loaded Then Return
            SyncLock _sync
                If _loaded Then Return
                _loaded = True
                Try
                    Dim p = GetStorePath()
                    If File.Exists(p) Then
                        Dim raw = File.ReadAllText(p, Encoding.UTF8)
                        Dim list = JsonConvert.DeserializeObject(Of List(Of SessionMemoryEntry))(raw)
                        If list IsNot Nothing Then
                            For Each e In list
                                If e IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(e.Key) Then
                                    _entries(e.Key) = e
                                End If
                            Next
                        End If
                    End If
                Catch
                    ' ignore corrupt memory; start clean
                    _entries = New Dictionary(Of String, SessionMemoryEntry)(StringComparer.Ordinal)
                End Try
            End SyncLock
        End Sub

        Private Shared Sub SaveUnlocked()
            Try
                Dim p = GetStorePath()
                Dim list = _entries.Values.OrderBy(Function(e) e.CreatedAt).ToList()
                Dim json = JsonConvert.SerializeObject(list, Formatting.Indented)
                File.WriteAllText(p, json, Encoding.UTF8)
            Catch
                ' best-effort
            End Try
        End Sub

        ' ------------------------------------------------------------------ API

        Public Shared Function Put(key As String,
                           summary As String,
                           value As JToken,
                           Optional tags As IEnumerable(Of String) = Nothing,
                           Optional metadata As SessionMemoryMetadata = Nothing) As SessionMemoryEntry
            EnsureLoaded()

            SyncLock _sync
                If String.IsNullOrWhiteSpace(key) Then
                    key = "mem_" & Guid.NewGuid().ToString("N").Substring(0, 12)
                End If

                Dim normalizedMetadata As SessionMemoryMetadata =
            WorkflowContinuity.EnsureMetadataDefaults(metadata)

                Dim nowUtc As DateTime = DateTime.UtcNow
                Dim tagList As New List(Of String)()

                If tags IsNot Nothing Then
                    tagList.AddRange(tags.Where(Function(t) Not String.IsNullOrWhiteSpace(t)).Select(Function(t) t.Trim()))
                End If

                If Not String.IsNullOrWhiteSpace(normalizedMetadata.WorkflowId) Then
                    Dim workflowTag As String = "workflow:" & normalizedMetadata.WorkflowId.Trim()
                    If Not tagList.Any(Function(t) t.Equals(workflowTag, StringComparison.OrdinalIgnoreCase)) Then
                        tagList.Add(workflowTag)
                    End If
                End If

                Dim hostPipeline As String = If(WorkflowContinuity.CurrentHostPipeline, "").Trim().ToLowerInvariant()

                If hostPipeline <> "" Then
                    Dim hostTag As String = "host:" & hostPipeline
                    If Not tagList.Any(Function(t) t.Equals(hostTag, StringComparison.OrdinalIgnoreCase)) Then
                        tagList.Add(hostTag)
                    End If
                End If

                Dim existing As SessionMemoryEntry = Nothing
                _entries.TryGetValue(key, existing)

                Dim createdAt As DateTime =
            If(existing IsNot Nothing AndAlso existing.CreatedAt <> DateTime.MinValue,
               existing.CreatedAt,
               nowUtc)

                Dim entry As New SessionMemoryEntry With {
            .Key = key,
            .Summary = If(summary, ""),
            .Value = If(value, JValue.CreateNull()),
            .CreatedAt = createdAt,
            .UpdatedAt = nowUtc,
            .Tags = tagList.Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
            .Metadata = normalizedMetadata
        }

                _entries(key) = entry
                SaveUnlocked()

                Dim sourceRecord As WorkflowSourceRecord = Nothing
                If normalizedMetadata IsNot Nothing AndAlso
           String.Equals(normalizedMetadata.ContentKind, "source_record", StringComparison.OrdinalIgnoreCase) Then
                    WorkflowContinuity.TryParseSourceRecord(entry, sourceRecord)
                End If

                If Not String.IsNullOrWhiteSpace(normalizedMetadata.WorkflowId) Then
                    Dim checkpointWritten As Boolean =
                WorkflowContinuity.NoteMemoryReferenceCreated(
                    normalizedMetadata.WorkflowId,
                    WorkflowContinuity.CurrentHostPipeline,
                    entry.Key,
                    normalizedMetadata,
                    sourceRecord)

                    Debug.WriteLine(
                WorkflowContinuity.BuildWorkflowLogLabel(
                    normalizedMetadata.WorkflowId,
                    If(
                        String.Equals(normalizedMetadata.ContentKind, "source_record", StringComparison.OrdinalIgnoreCase),
                        "source_reference_created",
                        "memory_reference_created")) &
                " memory_key=" & entry.Key &
                " source=" & normalizedMetadata.Source &
                " kind=" & normalizedMetadata.ContentKind &
                " checkpoint_written=" & If(checkpointWritten, "true", "false"))
                End If

                Return entry
            End SyncLock
        End Function


        Public Shared Function [Get](key As String) As SessionMemoryEntry
            EnsureLoaded()
            SyncLock _sync
                Dim e As SessionMemoryEntry = Nothing
                _entries.TryGetValue(If(key, ""), e)
                Return e
            End SyncLock
        End Function

        Public Shared Function List() As List(Of SessionMemoryEntry)
            EnsureLoaded()
            SyncLock _sync
                Return _entries.Values.OrderBy(Function(e) e.CreatedAt).ToList()
            End SyncLock
        End Function

        Public Shared Function Delete(key As String) As Boolean
            EnsureLoaded()
            SyncLock _sync
                If String.IsNullOrWhiteSpace(key) Then Return False
                Dim removed = _entries.Remove(key)
                If removed Then SaveUnlocked()
                Return removed
            End SyncLock
        End Function

        Public Shared Sub Clear()
            EnsureLoaded()
            SyncLock _sync
                _entries.Clear()
                SaveUnlocked()
            End SyncLock
        End Sub

        ''' <summary>
        ''' Produces a one-line stub the model can see in lieu of a full value.
        ''' Used when tool results are large and we want context to stay lean.
        ''' </summary>
        Public Shared Function BuildStub(entry As SessionMemoryEntry) As String
            If entry Is Nothing Then Return ""
            Dim summary = If(entry.Summary, "").Trim()
            If summary.Length = 0 Then summary = "(stored in session memory)"
            Return "[memory:" & entry.Key & "] " & summary
        End Function

        ' ------------------------------------------------------------------ shared-properties bridge

        Private Shared Function TryGetSharedPath(propertyName As String) As String
            Try
                Dim asm = GetType(SharedLibrary.SharedContext).Assembly
                For Each typeFullName In {"SharedLibrary.SharedProperties", "SharedLibrary.SharedContext"}
                    Dim t = asm.GetType(typeFullName, throwOnError:=False, ignoreCase:=False)
                    If t Is Nothing Then Continue For
                    Dim pi = t.GetProperty(propertyName,
                        Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance)
                    If pi Is Nothing Then Continue For
                    Dim val As Object = Nothing
                    Dim getter = pi.GetGetMethod()
                    If getter IsNot Nothing AndAlso getter.IsStatic Then
                        val = pi.GetValue(Nothing, Nothing)
                    End If
                    If TypeOf val Is String AndAlso Not String.IsNullOrWhiteSpace(CStr(val)) Then
                        Return CStr(val)
                    End If
                Next
            Catch
            End Try
            Return Nothing
        End Function

        Public Shared Function ListByWorkflowId(workflowId As String,
                                        Optional maxItems As Integer = 20) As List(Of SessionMemoryEntry)
            EnsureLoaded()

            If String.IsNullOrWhiteSpace(workflowId) Then
                Return New List(Of SessionMemoryEntry)()
            End If

            Dim workflowTag As String = "workflow:" & workflowId.Trim()

            SyncLock _sync
                Return _entries.Values.
                    Where(
                        Function(entry)
                            If entry Is Nothing Then Return False

                            Dim metadataWorkflowId As String = If(entry.Metadata?.WorkflowId, "").Trim()

                            If metadataWorkflowId.Equals(workflowId.Trim(), StringComparison.OrdinalIgnoreCase) Then
                                Return True
                            End If

                            If entry.Tags Is Nothing Then Return False

                            Return entry.Tags.Any(
                                Function(tag)
                                    Return Not String.IsNullOrWhiteSpace(tag) AndAlso
                                           tag.Trim().Equals(workflowTag, StringComparison.OrdinalIgnoreCase)
                                End Function)
                        End Function).
                    OrderByDescending(
                        Function(entry)
                            If entry Is Nothing Then Return DateTime.MinValue
                            If entry.UpdatedAt <> DateTime.MinValue Then Return entry.UpdatedAt
                            Return entry.CreatedAt
                        End Function).
                    Take(Math.Max(1, maxItems)).
                    ToList()
            End SyncLock
        End Function

        Public Shared Function ListMostRecentWorkflowEntries(Optional excludedWorkflowId As String = "",
                                                             Optional maxItems As Integer = 20) As List(Of SessionMemoryEntry)
            EnsureLoaded()

            Dim normalizedExcludedWorkflowId As String = If(excludedWorkflowId, "").Trim()

            SyncLock _sync
                Dim latestWorkflowId As String =
                    _entries.Values.
                        Where(
                            Function(entry)
                                If entry Is Nothing Then Return False

                                Dim metadataWorkflowId As String = If(entry.Metadata?.WorkflowId, "").Trim()
                                If metadataWorkflowId = "" Then Return False

                                Return Not metadataWorkflowId.Equals(normalizedExcludedWorkflowId, StringComparison.OrdinalIgnoreCase)
                            End Function).
                        OrderByDescending(
                            Function(entry)
                                If entry Is Nothing Then Return DateTime.MinValue
                                If entry.UpdatedAt <> DateTime.MinValue Then Return entry.UpdatedAt
                                Return entry.CreatedAt
                            End Function).
                        Select(Function(entry) If(entry.Metadata?.WorkflowId, "").Trim()).
                        FirstOrDefault()

                If String.IsNullOrWhiteSpace(latestWorkflowId) Then
                    Return New List(Of SessionMemoryEntry)()
                End If

                Return _entries.Values.
                    Where(
                        Function(entry)
                            If entry Is Nothing Then Return False
                            Return String.Equals(
                                If(entry.Metadata?.WorkflowId, "").Trim(),
                                latestWorkflowId,
                                StringComparison.OrdinalIgnoreCase)
                        End Function).
                    OrderByDescending(
                        Function(entry)
                            If entry Is Nothing Then Return DateTime.MinValue
                            If entry.UpdatedAt <> DateTime.MinValue Then Return entry.UpdatedAt
                            Return entry.CreatedAt
                        End Function).
                    Take(Math.Max(1, maxItems)).
                    ToList()
            End SyncLock
        End Function

        Private Shared Function IsTransientContentKind(contentKind As String) As Boolean
            Select Case If(contentKind, "").Trim().ToLowerInvariant()
                Case "runtime_state", "tool_result", "draft", "source_record"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Public Shared Function ClearTransientEntriesForHost(hostPipeline As String,
                                                            Optional keepWorkflowId As String = "") As Integer
            EnsureLoaded()

            Dim normalizedHost As String = If(hostPipeline, "").Trim().ToLowerInvariant()
            Dim normalizedKeepWorkflowId As String = If(keepWorkflowId, "").Trim()

            If normalizedHost = "" Then
                Return 0
            End If

            Dim hostTag As String = "host:" & normalizedHost

            SyncLock _sync
                Dim keysToRemove As List(Of String) =
                    _entries.Values.
                        Where(
                            Function(entry)
                                If entry Is Nothing OrElse String.IsNullOrWhiteSpace(entry.Key) Then
                                    Return False
                                End If

                                If Not IsTransientContentKind(If(entry.Metadata?.ContentKind, "")) Then
                                    Return False
                                End If

                                If entry.Tags Is Nothing OrElse
                                   Not entry.Tags.Any(
                                       Function(tag)
                                           Return Not String.IsNullOrWhiteSpace(tag) AndAlso
                                                  tag.Trim().Equals(hostTag, StringComparison.OrdinalIgnoreCase)
                                       End Function) Then
                                    Return False
                                End If

                                Dim entryWorkflowId As String = If(entry.Metadata?.WorkflowId, "").Trim()

                                If normalizedKeepWorkflowId <> "" AndAlso
                                   entryWorkflowId.Equals(normalizedKeepWorkflowId, StringComparison.OrdinalIgnoreCase) Then
                                    Return False
                                End If

                                Return True
                            End Function).
                        Select(Function(entry) entry.Key).
                        ToList()

                For Each key As String In keysToRemove
                    _entries.Remove(key)
                Next

                If keysToRemove.Count > 0 Then
                    SaveUnlocked()
                End If

                Return keysToRemove.Count
            End SyncLock
        End Function

    End Class

End Namespace