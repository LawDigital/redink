' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Explicit async-flow boundary for unattended hosts; interactive Office calls are unchanged.
Option Strict On
Option Explicit On

Namespace SharedLibrary
    Partial Public Class SharedMethods
        Private Shared ReadOnly HeadlessStateSlot As New System.Threading.AsyncLocal(Of HeadlessExecutionState)()

        Private NotInheritable Class HeadlessExecutionState
            Friend Failure As HeadlessInteractionRequiredException
        End Class

        Public NotInheritable Class HeadlessInteractionRequiredException
            Inherits System.InvalidOperationException
            Public ReadOnly Property Code As System.String

            Public Sub New(operation As System.String, code As System.String)
                MyBase.New(code & ": Interactive setup is required for " & operation & ". Complete it in the interactive application before retrying this operation.")
                Me.Code = code
            End Sub
        End Class

        Public NotInheritable Class HeadlessExecutionScope
            Implements System.IDisposable
            Private ReadOnly _previous As HeadlessExecutionState
            Private ReadOnly _state As HeadlessExecutionState
            Private _disposed As System.Boolean

            Friend Sub New()
                _previous = HeadlessStateSlot.Value
                _state = If(_previous, New HeadlessExecutionState())
                HeadlessStateSlot.Value = _state
            End Sub

            Public Sub ThrowIfInteractionRequested()
                Dim failure As HeadlessInteractionRequiredException = System.Threading.Volatile.Read(_state.Failure)
                If failure IsNot Nothing Then Throw failure
            End Sub

            Public Sub Dispose() Implements System.IDisposable.Dispose
                If _disposed Then Return
                HeadlessStateSlot.Value = _previous
                _disposed = True
            End Sub
        End Class

        Public Shared Function BeginHeadlessExecution() As HeadlessExecutionScope
            Return New HeadlessExecutionScope()
        End Function

        Public Shared ReadOnly Property IsHeadlessExecution As System.Boolean
            Get
                Return HeadlessStateSlot.Value IsNot Nothing
            End Get
        End Property

        ''' <summary>Call before creating a dialog, broker, browser or interactive listener.</summary>
        Public Shared Sub RequireInteractiveExecution(operation As System.String,
                                                       Optional code As System.String = "headless_interaction_required")
            Dim state As HeadlessExecutionState = HeadlessStateSlot.Value
            If state Is Nothing Then Return
            Dim failure As New HeadlessInteractionRequiredException(operation, code)
            System.Threading.Interlocked.CompareExchange(state.Failure, failure, Nothing)
            Throw failure
        End Sub
    End Class
End Namespace
