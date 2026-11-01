' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: BackgroundMaintenanceCoordinator.vb
' Purpose:
'   Generic background-provider scheduling, idle gating, interactive-work cancellation
'   and processing windows.
'
' Architecture / Function:
'   Coordinates registered providers independently of host/business tools and yields
'   when interactive work takes priority.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    ''' <summary>Provider-neutral Office maintenance scheduling. Each registration owns its controls,
    ''' cancellation, retry deadline and queue. Office state is supplied only as a UI-thread snapshot.</summary>
    Public NotInheritable Class BackgroundMaintenanceCoordinator
        Implements System.IDisposable

        Public NotInheritable Class Provider
            Public Property Id As String = ""
            Public Property RefreshControls As System.Action
            Public Property IsEnabled As System.Func(Of Boolean)
            Public Property CanRunNow As System.Func(Of Boolean)
            Public Property RunBatchAsync As System.Func(Of System.Threading.CancellationToken, System.Threading.Tasks.Task(Of Integer))
            Public Property Shutdown As System.Action
            Public Property MinimumIntervalSeconds As Integer
            Public Property IdleIntervalSeconds As Integer = 60
            Friend Cancellation As System.Threading.CancellationTokenSource
            Friend NextRunUtc As System.DateTime = System.DateTime.MinValue
            Friend FailureCount As Integer
        End Class

        Private Shared ReadOnly AmbientCancellation As New System.Threading.AsyncLocal(Of System.Threading.CancellationToken)()
        Private Shared ReadOnly InstancesLock As New Object()
        Private Shared ReadOnly Instances As New System.Collections.Generic.List(Of BackgroundMaintenanceCoordinator)()
        Private Shared _interactiveCount As Integer
        Private ReadOnly _lock As New Object()
        Private ReadOnly _providers As New System.Collections.Generic.List(Of Provider)()
        Private _running As Provider
        Private _lastProvider As Integer = -1
        Private _ticking As Integer
        Private _disposed As Integer

        Public Sub New()
            SyncLock InstancesLock
                Instances.Add(Me)
            End SyncLock
        End Sub

        Public Shared ReadOnly Property CurrentCancellationToken As System.Threading.CancellationToken
            Get
                Return AmbientCancellation.Value
            End Get
        End Property

        ''' <summary>Interactive retrieval never enters the idle gate. It asks automatic providers to
        ''' yield at their next checkpoint, without cancelling manual jobs or waiting for a chat to idle.</summary>
        Public Shared Function EnterInteractiveWork() As System.IDisposable
            System.Threading.Interlocked.Increment(_interactiveCount)
            SyncLock InstancesLock
                For Each coordinator In Instances
                    coordinator.CancelRunningBatch()
                Next
            End SyncLock
            Return New InteractiveLease()
        End Function

        Public Sub Register(provider As Provider)
            If provider Is Nothing OrElse System.String.IsNullOrWhiteSpace(provider.Id) OrElse
               provider.IsEnabled Is Nothing OrElse provider.CanRunNow Is Nothing OrElse provider.RunBatchAsync Is Nothing Then
                Throw New System.ArgumentException("Maintenance provider registration is incomplete.", NameOf(provider))
            End If
            SyncLock _lock
                If System.Threading.Volatile.Read(_disposed) <> 0 Then Throw New System.ObjectDisposedException(NameOf(BackgroundMaintenanceCoordinator))
                For Each existing In _providers
                    If System.String.Equals(existing.Id, provider.Id, System.StringComparison.OrdinalIgnoreCase) Then
                        Throw New System.ArgumentException("A maintenance provider with this identity is already registered.")
                    End If
                Next
                _providers.Add(provider)
            End SyncLock
        End Sub

        ''' <summary>Refreshes each provider independently even while another batch is running. All
        ''' control reads, scans and batches execute on workers; the caller passes no Office objects.</summary>
        Public Function OnTickAsync(hostIsIdle As Boolean) As System.Threading.Tasks.Task
            If System.Threading.Volatile.Read(_disposed) <> 0 Then Return System.Threading.Tasks.Task.CompletedTask
            Return System.Threading.Tasks.Task.Run(Function() TickCoreAsync(hostIsIdle))
        End Function

        Private Async Function TickCoreAsync(hostIsIdle As Boolean) As System.Threading.Tasks.Task
            If System.Threading.Interlocked.CompareExchange(_ticking, 1, 0) <> 0 Then Return
            Dim selected As Provider = Nothing
            Try
                Dim snapshot As Provider()
                SyncLock _lock
                    snapshot = _providers.ToArray()
                End SyncLock
                Dim eligibility As New System.Collections.Generic.Dictionary(Of Provider, Boolean)()
                For Each provider In snapshot
                    Try
                        If provider.RefreshControls IsNot Nothing Then provider.RefreshControls.Invoke()
                        eligibility(provider) = provider.IsEnabled.Invoke() AndAlso provider.CanRunNow.Invoke()
                    Catch ex As System.Exception
                        eligibility(provider) = False
                        System.Diagnostics.Debug.WriteLine("Maintenance controls [" & provider.Id & "]: " & ex.Message)
                    End Try
                Next
                SyncLock _lock
                    If System.Threading.Volatile.Read(_disposed) <> 0 Then Return
                    If _running IsNot Nothing AndAlso
                       (Not hostIsIdle OrElse System.Threading.Volatile.Read(_interactiveCount) > 0 OrElse
                        Not eligibility.ContainsKey(_running) OrElse Not eligibility(_running)) Then
                        CancelRunningBatchCore()
                    End If
                    If _running IsNot Nothing OrElse Not hostIsIdle OrElse System.Threading.Volatile.Read(_interactiveCount) > 0 Then Return
                    For offset As Integer = 1 To _providers.Count
                        Dim index = (_lastProvider + offset) Mod _providers.Count
                        Dim provider = _providers(index)
                        If eligibility.ContainsKey(provider) AndAlso eligibility(provider) AndAlso System.DateTime.UtcNow >= provider.NextRunUtc Then
                            selected = provider
                            _lastProvider = index
                            provider.Cancellation = New System.Threading.CancellationTokenSource()
                            _running = provider
                            Exit For
                        End If
                    Next
                End SyncLock
            Finally
                System.Threading.Interlocked.Exchange(_ticking, 0)
            End Try
            If selected Is Nothing Then Return

            Dim previous = AmbientCancellation.Value
            AmbientCancellation.Value = selected.Cancellation.Token
            Try
                Dim completed = Await selected.RunBatchAsync.Invoke(selected.Cancellation.Token).ConfigureAwait(False)
                selected.FailureCount = 0
                selected.NextRunUtc = System.DateTime.UtcNow.AddSeconds(System.Math.Max(0, If(completed > 0, selected.MinimumIntervalSeconds, selected.IdleIntervalSeconds)))
            Catch ex As System.OperationCanceledException
                selected.NextRunUtc = System.DateTime.UtcNow
            Catch ex As System.Exception
                selected.FailureCount = System.Math.Min(selected.FailureCount + 1, 8)
                selected.NextRunUtc = System.DateTime.UtcNow.AddSeconds(System.Math.Min(300, 5 * System.Math.Pow(2, selected.FailureCount)))
                System.Diagnostics.Debug.WriteLine("Maintenance batch [" & selected.Id & "] failed; retry after " & selected.NextRunUtc.ToString("O") & ": " & ex.Message)
            Finally
                AmbientCancellation.Value = previous
                SyncLock _lock
                    If selected.Cancellation IsNot Nothing Then selected.Cancellation.Dispose()
                    selected.Cancellation = Nothing
                    If System.Object.ReferenceEquals(_running, selected) Then _running = Nothing
                End SyncLock
                If System.Threading.Volatile.Read(_disposed) <> 0 AndAlso selected.Shutdown IsNot Nothing Then
                    Try
                        selected.Shutdown.Invoke()
                    Catch ex As System.Exception
                        System.Diagnostics.Debug.WriteLine("Maintenance shutdown [" & selected.Id & "]: " & ex.Message)
                    End Try
                End If
            End Try
        End Function

        Public Sub CancelProvider(providerId As String)
            SyncLock _lock
                If _running IsNot Nothing AndAlso System.String.Equals(_running.Id, providerId, System.StringComparison.OrdinalIgnoreCase) Then
                    CancelRunningBatchCore()
                End If
            End SyncLock
        End Sub

        Private Sub CancelRunningBatch()
            SyncLock _lock
                CancelRunningBatchCore()
            End SyncLock
        End Sub

        Private Sub CancelRunningBatchCore()
            If _running IsNot Nothing AndAlso _running.Cancellation IsNot Nothing Then
                _running.Cancellation.Cancel()
            End If
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            If System.Threading.Interlocked.Exchange(_disposed, 1) <> 0 Then Return
            SyncLock InstancesLock
                Instances.Remove(Me)
            End SyncLock
            Dim stopped As New System.Collections.Generic.List(Of Provider)()
            SyncLock _lock
                CancelRunningBatchCore()
                For Each provider In _providers
                    If Not System.Object.ReferenceEquals(provider, _running) Then stopped.Add(provider)
                Next
            End SyncLock
            For Each provider In stopped
                If provider.Shutdown Is Nothing Then Continue For
                Try
                    provider.Shutdown.Invoke()
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("Maintenance shutdown [" & provider.Id & "]: " & ex.Message)
                End Try
            Next
        End Sub

        Private NotInheritable Class InteractiveLease
            Implements System.IDisposable
            Private _disposed As Integer
            Public Sub Dispose() Implements System.IDisposable.Dispose
                If System.Threading.Interlocked.Exchange(_disposed, 1) = 0 Then System.Threading.Interlocked.Decrement(_interactiveCount)
            End Sub
        End Class
    End Class

    ''' <summary>Strict processing-window grammar shared by providers that opt into it.</summary>
    Public NotInheritable Class BackgroundProcessingWindow
        Private Sub New()
        End Sub

        Public Shared Function IsValid(specification As String) As Boolean
            Dim diagnostic As String = Nothing
            Allows(specification, System.DateTime.Now, diagnostic)
            Return diagnostic Is Nothing
        End Function

        Public Shared Function Allows(specification As String, localNow As System.DateTime,
                                      Optional ByRef diagnostic As String = Nothing) As Boolean
            diagnostic = Nothing
            Dim spec = If(specification, "").Trim()
            If spec.Length = 0 Then Return True
            Dim allowMode As Boolean = True
            If spec.StartsWith("allow:", System.StringComparison.OrdinalIgnoreCase) Then
                spec = spec.Substring(6).Trim()
            ElseIf spec.StartsWith("deny:", System.StringComparison.OrdinalIgnoreCase) Then
                spec = spec.Substring(5).Trim()
                allowMode = False
            End If
            Dim matched As Boolean = False
            Dim count As Integer = 0
            For Each part In spec.Split(New Char() {";"c, ","c}, System.StringSplitOptions.RemoveEmptyEntries)
                Dim bounds = part.Trim().Split(New Char() {"-"c})
                Dim startTime As System.TimeSpan
                Dim endTime As System.TimeSpan
                Dim formats As String() = {"h\:mm", "hh\:mm", "h\:mm\:ss", "hh\:mm\:ss"}
                If bounds.Length <> 2 OrElse
                   Not System.TimeSpan.TryParseExact(bounds(0).Trim(), formats, System.Globalization.CultureInfo.InvariantCulture, startTime) OrElse
                   Not System.TimeSpan.TryParseExact(bounds(1).Trim(), formats, System.Globalization.CultureInfo.InvariantCulture, endTime) OrElse
                   startTime.TotalHours >= 24 OrElse endTime.TotalHours >= 24 Then
                    diagnostic = "Use allow:HH:mm-HH:mm or deny:HH:mm-HH:mm, with semicolons between ranges."
                    Return False
                End If
                count += 1
                Dim nowTime = localNow.TimeOfDay
                If startTime = endTime OrElse
                   (startTime < endTime AndAlso nowTime >= startTime AndAlso nowTime < endTime) OrElse
                   (startTime > endTime AndAlso (nowTime >= startTime OrElse nowTime < endTime)) Then matched = True
            Next
            If count = 0 Then
                diagnostic = "A processing window must contain at least one time range."
                Return False
            End If
            Return If(allowMode, matched, Not matched)
        End Function
    End Class
End Namespace
