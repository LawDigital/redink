' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolExecutorRegistry.vb
' Purpose: Per-host advertise/execute coupling. Solves problem #2 from user
'          report ("Tool has no APICall template defined") by ensuring a tool
'          can only be ADVERTISED if a host-side EXECUTOR is registered for it
'          (or it's a remote MCP/HTTP tool with populated APICall template).
'
' Registry Pattern:
'  1. Host registers internal/external tools at startup.
'  2. Before advertising, host calls IsAdvertisable(host, name).
'  3. Before dispatching, host calls GetExecutorKind(host, name).
'  4. Thread-safe for typical startup-register + runtime-read pattern.
'
' KEY ARCHITECTURE
'   • Surface: ToolExecutorRegistry (VB implementation).
'   • Entry points/helpers: Reset, RegisterInternal, RegisterExternal.
' =============================================================================

Option Explicit On
Option Strict On

Namespace Agents

    Public Enum ToolingHostKind
        Outlook
        Word
        Excel
    End Enum

    Public Enum ExecutorKind
        None
        Internal
        External
    End Enum

    ''' <summary>
    ''' Per-host registry of which tool names have a working executor on this host.
    ''' Thread-safe for the typical "register at startup, read at runtime" pattern.
    ''' </summary>
    Public Module ToolExecutorRegistry

        Private ReadOnly _lock As New Object()
        Private ReadOnly _byHost As New Dictionary(Of ToolingHostKind, Dictionary(Of String, ExecutorKind))()

        Public Sub Reset(host As ToolingHostKind)
            SyncLock _lock
                If _byHost.ContainsKey(host) Then _byHost.Remove(host)
            End SyncLock
        End Sub

        Public Sub RegisterInternal(host As ToolingHostKind, toolName As String)
            Register(host, toolName, ExecutorKind.Internal)
        End Sub

        Public Sub RegisterExternal(host As ToolingHostKind, toolName As String)
            Register(host, toolName, ExecutorKind.External)
        End Sub

        Public Sub Unregister(host As ToolingHostKind, toolName As String)
            If String.IsNullOrWhiteSpace(toolName) Then Return
            SyncLock _lock
                Dim map As Dictionary(Of String, ExecutorKind) = Nothing
                If _byHost.TryGetValue(host, map) Then
                    If map.ContainsKey(toolName) Then map.Remove(toolName)
                End If
            End SyncLock
        End Sub

        Private Sub Register(host As ToolingHostKind, toolName As String, kind As ExecutorKind)
            If String.IsNullOrWhiteSpace(toolName) Then Return
            SyncLock _lock
                Dim map As Dictionary(Of String, ExecutorKind) = Nothing
                If Not _byHost.TryGetValue(host, map) Then
                    map = New Dictionary(Of String, ExecutorKind)(StringComparer.OrdinalIgnoreCase)
                    _byHost(host) = map
                End If
                map(toolName.Trim()) = kind
            End SyncLock
        End Sub

        Public Function GetExecutorKind(host As ToolingHostKind, toolName As String) As ExecutorKind
            If String.IsNullOrWhiteSpace(toolName) Then Return ExecutorKind.None
            SyncLock _lock
                Dim map As Dictionary(Of String, ExecutorKind) = Nothing
                If Not _byHost.TryGetValue(host, map) Then Return ExecutorKind.None
                Dim kind As ExecutorKind = ExecutorKind.None
                If map.TryGetValue(toolName.Trim(), kind) Then Return kind
                Return ExecutorKind.None
            End SyncLock
        End Function

        ''' <summary>True if this tool may be ADVERTISED to the model on this host.</summary>
        Public Function IsAdvertisable(host As ToolingHostKind, toolName As String) As Boolean
            Return GetExecutorKind(host, toolName) <> ExecutorKind.None
        End Function

        ''' <summary>Snapshot of all currently registered tool names for the host (for diagnostics).</summary>
        Public Function ListRegistered(host As ToolingHostKind) As IReadOnlyList(Of String)
            SyncLock _lock
                Dim map As Dictionary(Of String, ExecutorKind) = Nothing
                If Not _byHost.TryGetValue(host, map) Then Return New List(Of String)().AsReadOnly()
                Dim copy As New List(Of String)(map.Keys)
                copy.Sort(StringComparer.OrdinalIgnoreCase)
                Return copy.AsReadOnly()
            End SyncLock
        End Function

        ''' <summary>
        ''' Builds the canonical "tool not executable on this host" structured response.
        ''' Hosts return this instead of the legacy "Tool has no APICall template defined" string.
        ''' </summary>
        Public Function BuildNotExecutablePayload(toolName As String, hostName As String) As String
            Dim obj As New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("summary", "Tool is not executable on this host."),
                New Newtonsoft.Json.Linq.JProperty("result", Nothing),
                New Newtonsoft.Json.Linq.JProperty("resultKind", "error"),
                New Newtonsoft.Json.Linq.JProperty("error",
                    New Newtonsoft.Json.Linq.JObject(
                        New Newtonsoft.Json.Linq.JProperty("code", "tool_not_executable"),
                        New Newtonsoft.Json.Linq.JProperty("phase", "tool_dispatch"),
                        New Newtonsoft.Json.Linq.JProperty("message",
                            "The tool '" & If(toolName, "") &
                            "' is not executable on host '" & If(hostName, "") &
                            "'. The model must choose a different tool, or declare blocked only after attempting available fallback tools.")
                    )
                )
            )
            Return obj.ToString(Newtonsoft.Json.Formatting.None)
        End Function

    End Module

End Namespace