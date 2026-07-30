' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: WorkspaceState.vb
' Purpose: Per-host persistent workspace state. Mirrors the shape of
'          ChatAgentWorkspaceState (Outlook) but lives in SharedLibrary so any
'          host (Word, Outlook, future Excel) can use it.
'
' Storage:
'  <localagentresources>\.session\workspace_<hostKey>.json
'  (Fallback to %LOCALAPPDATA%\RedInk\.session\ if no local agent resources path.)
'
' Properties:
'  - RootPath: workspace directory for agent file operations.
'  - Permissions: AllowRead, AllowWrite, AllowMoveCopyRename, AllowDelete.
'  - PersistUntilRevoked: survives host restarts until user clears it.
'  - IncludeHiddenSystem: whether hidden/system files are visible.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports SharedLibrary.SharedLibrary

Namespace Agents

    Public Class WorkspaceState
        Public Property RootPath As String = ""
        Public Property PersistUntilRevoked As Boolean = True
        Public Property AllowRead As Boolean = True
        Public Property AllowWrite As Boolean = True
        Public Property AllowMoveCopyRename As Boolean = True
        Public Property AllowDelete As Boolean = SharedMethods.WorkspaceDeleteByDefaultOn
        Public Property IncludeHiddenSystem As Boolean = False
    End Class

    Public NotInheritable Class WorkspaceStore

        Private Sub New()
        End Sub

        Private Shared ReadOnly _sync As New Object()

        Private Shared Function ResolveDir() As String
            Dim root As String = TryGetSharedPath("INI_AgentResourcesPathLocal")
            If String.IsNullOrWhiteSpace(root) Then
                root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RedInk")
            End If
            Dim dir = Path.Combine(root, ".session")
            Try
                If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
            Catch
            End Try
            Return dir
        End Function

        Private Shared Function FileFor(hostKey As String) As String
            Dim safe = If(hostKey, "default").Trim().ToLowerInvariant()
            If String.IsNullOrWhiteSpace(safe) Then safe = "default"
            Return Path.Combine(ResolveDir(), "workspace_" & safe & ".json")
        End Function

        Public Shared Function Load(hostKey As String) As WorkspaceState
            SyncLock _sync
                Try
                    Dim p = FileFor(hostKey)
                    If File.Exists(p) Then
                        Dim raw = File.ReadAllText(p, Encoding.UTF8)
                        Dim st = JsonConvert.DeserializeObject(Of WorkspaceState)(raw)
                        If st IsNot Nothing AndAlso st.PersistUntilRevoked AndAlso
                           Not String.IsNullOrWhiteSpace(st.RootPath) AndAlso
                           Directory.Exists(st.RootPath) Then
                            Return st
                        End If
                    End If
                Catch
                End Try
                Return New WorkspaceState()
            End SyncLock
        End Function

        Public Shared Sub Save(hostKey As String, state As WorkspaceState)
            If state Is Nothing Then Return
            SyncLock _sync
                Try
                    Dim p = FileFor(hostKey)
                    If state.PersistUntilRevoked AndAlso Not String.IsNullOrWhiteSpace(state.RootPath) Then
                        File.WriteAllText(p, JsonConvert.SerializeObject(state, Formatting.Indented), Encoding.UTF8)
                    ElseIf File.Exists(p) Then
                        File.Delete(p)
                    End If
                Catch
                End Try
            End SyncLock
        End Sub

        Public Shared Sub Clear(hostKey As String)
            SyncLock _sync
                Try
                    Dim p = FileFor(hostKey)
                    If File.Exists(p) Then File.Delete(p)
                Catch
                End Try
            End SyncLock
        End Sub

        Private Shared Function TryGetSharedPath(propertyName As String) As String
            Try
                Dim asm = GetType(SharedContext).Assembly
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

    End Class

End Namespace
