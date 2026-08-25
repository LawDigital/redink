' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolRegistry.vb
' Purpose: Lazy registry of tools (ModelConfig-shaped entries) keyed by name.
'          Lets the host add tools by manifest (cheap) and defer materialization
'          (full ModelConfig with ToolDefinition/ToolInstructionsPrompt) until
'          the tool is actually selected for an LLM call.
'
' Registration:
'  - RegisterEager(...): accepts fully-built ModelConfig (used to wrap existing
'    tool lists without behavior change).
'  - RegisterLazy(...): accepts manifest + factory; factory invoked once, cached.
'
' Thread-Safety:
'  - Per-instance registry (one per tooling loop or add-in lifetime).
'  - Coarse sync lock protects all mutations and lookups.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading

Namespace Agents

    ''' <summary>Lightweight manifest exposed for selection/UI without materializing the tool.</summary>
    Public Class ToolManifest
        Public Property Name As String
        Public Property Description As String
        Public Property Category As String       ' e.g. "builtin", "m365", "mcp", "autopilot", "workspace", "skill", "agent"
        Public Property [Source] As String       ' optional provenance (file path, skill name, ...)
    End Class

    Public NotInheritable Class ToolRegistry

        Private ReadOnly _sync As New Object()
        Private ReadOnly _entries As New Dictionary(Of String, ToolEntry)(StringComparer.OrdinalIgnoreCase)

        Private Class ToolEntry
            Public Manifest As ToolManifest
            Public Factory As Func(Of SharedLibrary.ModelConfig)
            Public Materialized As SharedLibrary.ModelConfig
        End Class

        ' --------------------------------------------------------------- registration

        ''' <summary>Registers a tool whose full <see cref="ModelConfig"/> is already known.</summary>
        Public Sub RegisterEager(config As SharedLibrary.ModelConfig, Optional category As String = "builtin", Optional source As String = Nothing)
            If config Is Nothing OrElse String.IsNullOrWhiteSpace(config.ToolName) Then Return
            ExplicitOperationToolContract.ApplyToModelConfig(config)
            SyncLock _sync
                _entries(config.ToolName) = New ToolEntry With {
                    .Manifest = New ToolManifest With {
                        .Name = config.ToolName,
                        .Description = DescribeFromConfig(config),
                        .Category = category,
                        .Source = source
                    },
                    .Factory = Nothing,
                    .Materialized = config
                }
            End SyncLock
        End Sub


        Public Function Snapshot() As ToolRegistry
            Dim copy As New ToolRegistry()

            SyncLock _sync
                For Each kv In _entries
                    copy._entries(kv.Key) = kv.Value
                Next
            End SyncLock

            Return copy
        End Function

        ''' <summary>Registers a tool by manifest; the factory is invoked once on first <see cref="Get"/>.</summary>
        Public Sub RegisterLazy(manifest As ToolManifest, factory As Func(Of SharedLibrary.ModelConfig))
            If manifest Is Nothing OrElse String.IsNullOrWhiteSpace(manifest.Name) Then Return
            If factory Is Nothing Then Return
            SyncLock _sync
                _entries(manifest.Name) = New ToolEntry With {
                    .Manifest = manifest,
                    .Factory = factory,
                    .Materialized = Nothing
                }
            End SyncLock
        End Sub

        ''' <summary>Removes a tool by name (no-op if not present).</summary>
        Public Sub Remove(name As String)
            If String.IsNullOrWhiteSpace(name) Then Return
            SyncLock _sync
                _entries.Remove(name)
            End SyncLock
        End Sub

        Public Sub Clear()
            SyncLock _sync
                _entries.Clear()
            End SyncLock
        End Sub

        ' --------------------------------------------------------------- query

        Public Function Contains(name As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False
            SyncLock _sync
                Return _entries.ContainsKey(name)
            End SyncLock
        End Function

        Public Function ListManifests() As IReadOnlyList(Of ToolManifest)
            SyncLock _sync
                Return _entries.Values.
                    Select(Function(e) e.Manifest).
                    OrderBy(Function(m) m.Name, StringComparer.OrdinalIgnoreCase).
                    ToList()
            End SyncLock
        End Function

        Public Function ListNames() As IReadOnlyList(Of String)
            SyncLock _sync
                Return _entries.Keys.OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase).ToList()
            End SyncLock
        End Function

        ''' <summary>Materializes a single tool by name. Returns Nothing if not registered or factory failed.</summary>
        Public Function [Get](name As String) As SharedLibrary.ModelConfig
            If String.IsNullOrWhiteSpace(name) Then Return Nothing
            Dim entry As ToolEntry = Nothing
            SyncLock _sync
                If Not _entries.TryGetValue(name, entry) Then Return Nothing
                If entry.Materialized IsNot Nothing Then Return entry.Materialized
            End SyncLock

            ' Materialize outside the lock to avoid holding it across user code.
            Dim built As SharedLibrary.ModelConfig = Nothing
            Try
                built = entry.Factory.Invoke()
            Catch
                built = Nothing
            End Try
            If built Is Nothing Then Return Nothing
            If String.IsNullOrWhiteSpace(built.ToolName) Then built.ToolName = name
            ExplicitOperationToolContract.ApplyToModelConfig(built)

            SyncLock _sync
                ' Re-check in case another caller materialized concurrently.
                Dim again As ToolEntry = Nothing
                If _entries.TryGetValue(name, again) Then
                    If again.Materialized Is Nothing Then
                        again.Materialized = built
                    End If
                    Return again.Materialized
                End If
            End SyncLock
            Return built
        End Function

        ''' <summary>
        ''' Materializes the given set of names in registration order. Unknown names are skipped.
        ''' </summary>
        Public Function Materialize(names As IEnumerable(Of String)) As List(Of SharedLibrary.ModelConfig)
            Dim result As New List(Of SharedLibrary.ModelConfig)()
            If names Is Nothing Then Return result
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each n In names
                If String.IsNullOrWhiteSpace(n) Then Continue For
                If Not seen.Add(n) Then Continue For
                Dim mc = [Get](n)
                If mc IsNot Nothing Then result.Add(mc)
            Next
            Return result
        End Function

        ''' <summary>Materializes all registered tools (use sparingly — defeats laziness).</summary>
        Public Function MaterializeAll() As List(Of SharedLibrary.ModelConfig)
            Return Materialize(ListNames())
        End Function

        ' --------------------------------------------------------------- filtering

        ''' <summary>
        ''' Returns a copy of this registry containing only entries whose name appears in
        ''' <paramref name="allowed"/>. Used to enforce skill/agent <c>allowed-tools</c>
        ''' as a NARROWING filter (never widens). Pass Nothing or an empty list to keep all.
        ''' </summary>
        Public Function Narrow(allowed As IEnumerable(Of String)) As ToolRegistry
            Dim child As New ToolRegistry()
            Dim hasAllowList As Boolean = (allowed IsNot Nothing)
            Dim allowSet As HashSet(Of String) = Nothing

            If hasAllowList Then
                allowSet = New HashSet(Of String)(
            allowed.
                Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                Select(Function(s) s.Trim()),
            StringComparer.OrdinalIgnoreCase)
            End If

            SyncLock _sync
                For Each kv In _entries
                    If hasAllowList AndAlso Not allowSet.Contains(kv.Key) Then Continue For
                    child._entries(kv.Key) = kv.Value
                Next
            End SyncLock

            Return child
        End Function

        ' --------------------------------------------------------------- helpers

        Private Shared Function DescribeFromConfig(config As SharedLibrary.ModelConfig) As String
            If config Is Nothing Then Return ""
            If Not String.IsNullOrWhiteSpace(config.ToolInstructionsPrompt) Then
                Dim p = config.ToolInstructionsPrompt.Trim()
                If p.Length > 200 Then p = p.Substring(0, 200) & "…"
                Return p
            End If
            Return If(config.ModelDescription, "")
        End Function

    End Class

End Namespace