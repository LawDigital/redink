' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SkillAuthorMode.vb
' Purpose: Process-wide flag + helpers for "skill author" mode. When active,
'          the agent layer is allowed to read AND write inside any discovered
'          skill's own folder (so the model can create/edit SKILL.md and scripts/).
'
' Activation:
'  - SkillAuthorMode.Enable() / .Disable() for persistent activation.
'  - BeginScope() for scoped, per-call activation (preferred for tool calls).
'  - IsActive property reflects both persistent and scoped states.
'  - Consumed by PathPolicy and SkillAuthorPathPolicy for read/write enforcement.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading

Namespace Agents

    Public NotInheritable Class SkillAuthorMode

        Private Sub New()
        End Sub

        Private Shared _persistent As Integer = 0
        Private Shared ReadOnly _scope As New AsyncLocal(Of Integer)

        ''' <summary>True when the mode is active (persistent and/or in a scope).</summary>
        Public Shared ReadOnly Property IsActive As Boolean
            Get
                Return Volatile.Read(_persistent) > 0 OrElse _scope.Value > 0
            End Get
        End Property

        Public Shared Sub Enable()
            Interlocked.Increment(_persistent)
            ' Make sure the local .inky tree (root + skills/ + agents/) exists so new
            ' resources and Inky.md can be created even on a fresh setup. If no local root
            ' is configured this is a no-op and author mode still works against central
            ' (when central writes are permitted).
            Try
                AgentResources.EnsureLocalResourceDirectories()
            Catch
            End Try
            PersistState()
        End Sub

        ''' <summary>Persists the current author-mode flags to My.Settings (best-effort, silent).</summary>
        Private Shared Sub PersistState()
            Try
                My.Settings.Item("SkillAuthorModeEnabled") = (Volatile.Read(_persistent) > 0)
                My.Settings.Item("SkillAuthorCentralWrites") = (Volatile.Read(_allowCentralWrites) > 0)
                My.Settings.Save()
            Catch
                ' My.Settings entry may not exist yet; ignore.
            End Try
        End Sub

        ''' <summary>
        ''' Restores persisted author-mode flags at startup. Enabling here also ensures the
        ''' local resource tree exists when a local root is configured; when there is no local
        ''' root, author mode is still restored and operates against central if permitted.
        ''' </summary>
        Public Shared Sub RestorePersistedState()
            Try
                Dim enabled As Boolean = False
                Dim central As Boolean = False
                Try : enabled = CBool(My.Settings.Item("SkillAuthorModeEnabled")) : Catch : End Try
                Try : central = CBool(My.Settings.Item("SkillAuthorCentralWrites")) : Catch : End Try

                If enabled AndAlso Volatile.Read(_persistent) = 0 Then
                    Enable()
                End If
                If central Then
                    Volatile.Write(_allowCentralWrites, 1)
                End If
            Catch
            End Try
        End Sub

        ' When False (default), author-mode writes are confined to the LOCAL resource
        ' root (the user's .inky directory). Set True to also permit writing into the
        ' shared/central resource root. Kept separate so the safe default is local-only.
        Private Shared _allowCentralWrites As Integer = 0

        Public Shared Property AllowCentralWrites As Boolean
            Get
                Return Volatile.Read(_allowCentralWrites) > 0
            End Get
            Set(value As Boolean)
                Volatile.Write(_allowCentralWrites, If(value, 1, 0))
                ' Make sure the central resource tree (root + skills/ + agents/) exists so
                ' new resources can be created there once central writing is permitted.
                If value Then
                    Try
                        AgentResources.EnsureCentralResourceDirectories()
                    Catch
                    End Try
                End If
                PersistState()
            End Set
        End Property

        Public Shared Sub Disable()
            If Volatile.Read(_persistent) > 0 Then Interlocked.Decrement(_persistent)
            PersistState()
        End Sub

        ''' <summary>Push/pop a scope. Best when scoping per-call (chat surface).</summary>
        Public Shared Function BeginScope() As IDisposable
            _scope.Value = _scope.Value + 1
            Return New Releaser()
        End Function

        Private Class Releaser
            Implements IDisposable
            Private _done As Boolean
            Public Sub Dispose() Implements IDisposable.Dispose
                If _done Then Return
                _done = True
                _scope.Value = Math.Max(0, _scope.Value - 1)
            End Sub
        End Class

    End Class

End Namespace
