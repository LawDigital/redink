' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: RedinkPythonAgentClient.vb
' Purpose: Defines the secure Python agent client, execution protocol models,
'          host-call contracts, limits, and safe result/error envelopes.
'
' Architecture / How it works:
'  - Declares the shared contract types used between the host broker and the
'    external Python worker: requests, responses, events, outputs, and errors.
'  - Centralizes configuration, executable-trust expectations, execution-time
'    options, and resource limits for secure worker startup and supervision.
'  - `RedInkPythonAgentClient` stages input files and request JSON, launches the
'    worker process, monitors heartbeat/events/host-calls, and parses the final
'    bounded response payload.
'  - Failures are normalized into typed execution/configuration/trust errors and
'    safe result objects instead of exposing arbitrary process diagnostics.
' =============================================================================

Option Explicit On
Option Strict On
Option Infer On

Namespace Agents

    Public Class RedInkPythonAgentException
        Inherits System.Exception
        Public Sub New(message As System.String)
            MyBase.New(message)
        End Sub
        Public Sub New(message As System.String, innerException As System.Exception)
            MyBase.New(message, innerException)
        End Sub
    End Class

    Public NotInheritable Class RedInkPythonAgentConfigurationException
        Inherits RedInkPythonAgentException
        Public Sub New(message As System.String)
            MyBase.New(message)
        End Sub
    End Class

    Public NotInheritable Class RedInkPythonAgentExecutableTrustException
        Inherits RedInkPythonAgentException
        Public ReadOnly Property Code As System.String
        Public Sub New(code As System.String, message As System.String)
            MyBase.New(message)
            Me.Code = code
        End Sub
        Public Sub New(message As System.String, innerException As System.Exception)
            MyBase.New(message, innerException)
            Me.Code = "EXECUTABLE_SIGNATURE_INVALID"
        End Sub
        Public Sub New(message As System.String)
            MyBase.New(message)
            Me.Code = "EXECUTABLE_SIGNATURE_INVALID"
        End Sub
    End Class

    Public NotInheritable Class RedInkPythonAgentExecutionException
        Inherits RedInkPythonAgentException
        Public Sub New(message As System.String)
            MyBase.New(message)
        End Sub
        Public Sub New(message As System.String, innerException As System.Exception)
            MyBase.New(message, innerException)
        End Sub
    End Class

    Public Interface IRedInkPythonAgentExecutableTrustValidator
        Sub Validate(executableFullPath As System.String)
    End Interface

    Public NotInheritable Class RedInkPythonAgentConfiguration
        Public Property ExecutablePath As System.String = System.String.Empty
        Public Property ExpectedSignerOrganization As System.String
        Public Property ExpectedSha256 As System.String
    End Class

    Public NotInheritable Class RedInkPythonAgentExecutionOptions
        Public Property OverallTimeout As System.TimeSpan = System.TimeSpan.FromMinutes(10)
        Public Property HeartbeatTimeout As System.TimeSpan = System.TimeSpan.FromSeconds(15)
        Public Property PollInterval As System.TimeSpan = System.TimeSpan.FromMilliseconds(100)
        Public Property CancellationGracePeriod As System.TimeSpan = System.TimeSpan.FromSeconds(3)
        Public Property HardKillWait As System.TimeSpan = System.TimeSpan.FromSeconds(5)
        Public Property HumanDiagnosticLog As System.Action(Of System.String)
    End Class

    Public NotInheritable Class RedInkPythonAgentInputFile
        Public Sub New(sourcePath As System.String, relativePath As System.String)
            Me.SourcePath = sourcePath
            Me.RelativePath = relativePath
        End Sub
        Public ReadOnly Property SourcePath As System.String
        Public ReadOnly Property RelativePath As System.String
    End Class

    Public NotInheritable Class RedInkPythonAgentLimits
        Public Property OverallWallTimeSeconds As System.Int32 = 2400
        Public Property StartupWallTimeSeconds As System.Int32 = 300
        Public Property ExecuteWallTimeSeconds As System.Int32 = 1800
        Public Property ExecutionInactivitySeconds As System.Int32 = 1800
        Public Property ValidatorWallTimeSeconds As System.Int32 = 300
        Public Property MemoryMiB As System.Int32 = 1536
        Public Property MaxOutputBytes As System.Int64 = 268435456L
        Public Property MaxOutputFiles As System.Int32 = 500
        Public Property MaxWorkingBytes As System.Int64 = 2147483648L
        Public Property MaxWorkingFiles As System.Int32 = 10000
        Public Property MaxResultBytes As System.Int64 = 67108864L
        Public Property MaxResultJsonDepth As System.Int32 = 128
        Public Property MaxResultJsonNodes As System.Int32 = 1000000
        Public Property AllowedOperations As System.Collections.Generic.IList(Of System.String) = New System.Collections.Generic.List(Of System.String)()
        Public Property MaximumCalls As System.Int32 = 20
        Public Property MaximumConcurrentCalls As System.Int32 = 2
        Public Property MaximumRequestBytes As System.Int32 = 67108864
        Public Property MaximumResponseBytes As System.Int32 = 134217728
        Public Property DefaultCallTimeoutSeconds As System.Int32 = 60
        Public Property MaximumCallTimeoutSeconds As System.Int32 = 180
    End Class

    Public Enum RedInkPythonAgentHostOperation
        LlmComplete
        WebGet
        WebSearch
    End Enum

    Public MustInherit Class RedInkPythonAgentHostCallArguments
    End Class
    Public NotInheritable Class RedInkPythonAgentLlmRequest
        Inherits RedInkPythonAgentHostCallArguments
        Public Property SystemPrompt As System.String = System.String.Empty
        Public Property UserPrompt As System.String = System.String.Empty
    End Class
    Public NotInheritable Class RedInkPythonAgentWebGetRequest
        Inherits RedInkPythonAgentHostCallArguments
        Public Property Url As System.String = System.String.Empty
        Public Property MaximumCharacters As System.Int32
    End Class
    Public NotInheritable Class RedInkPythonAgentWebSearchRequest
        Inherits RedInkPythonAgentHostCallArguments
        Public Property Query As System.String = System.String.Empty
        Public Property MaximumResults As System.Int32
    End Class
    Public NotInheritable Class RedInkPythonAgentWebSearchItem
        Public Property Title As System.String = System.String.Empty
        Public Property Url As System.String = System.String.Empty
        Public Property Snippet As System.String = System.String.Empty
    End Class
    Public NotInheritable Class RedInkPythonAgentHostCallRequest
        Public Property RequestId As System.Guid
        Public Property Operation As RedInkPythonAgentHostOperation
        Public Property Timeout As System.TimeSpan
        Public Property Arguments As RedInkPythonAgentHostCallArguments
    End Class
    Public NotInheritable Class RedInkPythonAgentHostCallResponse
        Public Property IsSuccess As System.Boolean
        Public Property Text As System.String
        Public Property SearchResults As System.Collections.Generic.IReadOnlyList(Of RedInkPythonAgentWebSearchItem)
        Public Property ErrorCode As System.String
        Public Property Retryable As System.Boolean
        Public Property RetryAfterMilliseconds As System.Nullable(Of System.Int32)
        Public Shared Function SuccessLlm(text As System.String) As RedInkPythonAgentHostCallResponse
            Return New RedInkPythonAgentHostCallResponse() With {.IsSuccess = True, .Text = text}
        End Function
        Public Shared Function SuccessWebGet(text As System.String) As RedInkPythonAgentHostCallResponse
            Return SuccessLlm(text)
        End Function
        Public Shared Function SuccessWebSearch(results As System.Collections.Generic.IReadOnlyList(Of RedInkPythonAgentWebSearchItem)) As RedInkPythonAgentHostCallResponse
            Return New RedInkPythonAgentHostCallResponse() With {.IsSuccess = True, .SearchResults = results}
        End Function
        Public Shared Function Failure(code As System.String, retryable As System.Boolean, retryAfterMilliseconds As System.Nullable(Of System.Int32)) As RedInkPythonAgentHostCallResponse
            Return New RedInkPythonAgentHostCallResponse() With {.IsSuccess = False, .ErrorCode = code, .Retryable = retryable, .RetryAfterMilliseconds = retryAfterMilliseconds}
        End Function
    End Class

    Public Interface IRedInkPythonAgentHostServiceHandler
        Function HandleAsync(request As RedInkPythonAgentHostCallRequest, cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of RedInkPythonAgentHostCallResponse)
    End Interface

    Public NotInheritable Class RedInkPythonAgentEvent
        Public Property Sequence As System.Int64
        Public Property EventCode As System.String = System.String.Empty
        Public Property Phase As System.String = System.String.Empty
        Public Property Current As System.Nullable(Of System.Int32)
        Public Property Total As System.Nullable(Of System.Int32)
        Public Property StepId As System.String
        Public Property HostOperation As System.String
        Public Property HostCallIndex As System.Nullable(Of System.Int32)
    End Class

    Public NotInheritable Class RedInkPythonAgentOutput
        Public Property RelativePath As System.String = System.String.Empty
        Public Property MediaType As System.String = System.String.Empty
        Public Property Size As System.Int64
        Public Property Sha256 As System.String = System.String.Empty
        Public Property FullPath As System.String = System.String.Empty
        ' Path of the output relative to results/<session-id>, set by the core only after the
        ' file passed containment/size/hash validation. Hosts use it to preserve subdirectories.
        Public Property PublishedSubPath As System.String = System.String.Empty
    End Class

    Public NotInheritable Class RedInkPythonAgentSafeSource
        Public Property File As System.String = System.String.Empty
        Public Property Line As System.Int32
        Public Property Column As System.Nullable(Of System.Int32)
        Public Property [Function] As System.String
        Public Property Symbol As System.String
    End Class

    Public NotInheritable Class RedInkPythonAgentSafeStackFrame
        Public Property File As System.String = System.String.Empty
        Public Property Line As System.Int32
        Public Property [Function] As System.String
    End Class

    Public NotInheritable Class RedInkPythonAgentSafeError
        Public Property Code As System.String = System.String.Empty
        Public Property Phase As System.String = System.String.Empty
        Public Property Retryable As System.Boolean
        Public Property Source As RedInkPythonAgentSafeSource
        Public Property HostOperation As System.String
        Public Property Limit As System.Nullable(Of System.Int64)
        Public Property Observed As System.Nullable(Of System.Int64)
        Public Property Stack As System.Collections.Generic.List(Of RedInkPythonAgentSafeStackFrame) = New System.Collections.Generic.List(Of RedInkPythonAgentSafeStackFrame)()
    End Class

    Public NotInheritable Class RedInkPythonAgentPublishedResult
        Public Property Kind As System.String = System.String.Empty
        Public Property Value As Newtonsoft.Json.Linq.JToken
    End Class

    Public NotInheritable Class RedInkPythonAgentRunResult
        Public Property ProtocolVersion As System.Int32
        Public Property SessionId As System.Guid
        Public Property Status As System.String = System.String.Empty
        Public Property ExitCode As System.Int32
        Public Property DiagnosticId As System.Guid
        Public Property HumanLogAvailable As System.Boolean
        Public Property Outputs As System.Collections.Generic.List(Of RedInkPythonAgentOutput) = New System.Collections.Generic.List(Of RedInkPythonAgentOutput)()
        Public Property Result As RedInkPythonAgentPublishedResult
        Public Property [Error] As RedInkPythonAgentSafeError
    End Class

    Public NotInheritable Class RedInkPythonAgentExecution
        Private ReadOnly Cancellation As System.Threading.CancellationTokenSource
        Private ReadOnly CompletionValue As System.Threading.Tasks.Task(Of RedInkPythonAgentRunResult)
        Friend Sub New(sessionId As System.Guid, cancellation As System.Threading.CancellationTokenSource, completion As System.Threading.Tasks.Task(Of RedInkPythonAgentRunResult))
            Me.SessionId = sessionId
            Me.Cancellation = cancellation
            Me.CompletionValue = completion
        End Sub
        Public ReadOnly Property SessionId As System.Guid
        Public ReadOnly Property Completion As System.Threading.Tasks.Task(Of RedInkPythonAgentRunResult)
            Get
                Return Me.CompletionValue
            End Get
        End Property
        Public Function CancelAsync() As System.Threading.Tasks.Task
            Me.Cancellation.Cancel()
            Return System.Threading.Tasks.Task.CompletedTask
        End Function
    End Class

    Public NotInheritable Class RedInkPythonAgentClient
        Private Const ProtocolVersion As System.Int32 = 2
        Private ReadOnly Options As RedInkPythonAgentExecutionOptions
        Public Sub New(Optional options As RedInkPythonAgentExecutionOptions = Nothing)
            Me.Options = If(options, New RedInkPythonAgentExecutionOptions())
            ValidateOptions(Me.Options)
        End Sub

        Public Shared Function ValidateExecutableConfiguration(configuration As RedInkPythonAgentConfiguration) As System.String
            If configuration Is Nothing Then Throw New System.ArgumentNullException(NameOf(configuration))
            Dim executable As System.String = ResolveExecutable(configuration.ExecutablePath)
            Using heldExecutable As New System.IO.FileStream(executable, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                VerifyConfiguration(configuration, executable, heldExecutable)
            End Using
            Return executable
        End Function

        Public Function CreateExecution(configuration As RedInkPythonAgentConfiguration, rootPath As System.String, pythonCode As System.String, inputFiles As System.Collections.Generic.IEnumerable(Of RedInkPythonAgentInputFile), limits As RedInkPythonAgentLimits, hostServiceHandler As IRedInkPythonAgentHostServiceHandler, progress As System.IProgress(Of RedInkPythonAgentEvent)) As RedInkPythonAgentExecution
            Dim sessionId As System.Guid = System.Guid.NewGuid()
            Dim cancellation As New System.Threading.CancellationTokenSource()
            Dim completion As System.Threading.Tasks.Task(Of RedInkPythonAgentRunResult) = System.Threading.Tasks.Task.Run(Function() RunAsync(configuration, rootPath, pythonCode, inputFiles, limits, hostServiceHandler, progress, sessionId, cancellation.Token))
            Return New RedInkPythonAgentExecution(sessionId, cancellation, completion)
        End Function

        Private Async Function RunAsync(configuration As RedInkPythonAgentConfiguration, rootPath As System.String, pythonCode As System.String, inputFiles As System.Collections.Generic.IEnumerable(Of RedInkPythonAgentInputFile), limits As RedInkPythonAgentLimits, hostServiceHandler As IRedInkPythonAgentHostServiceHandler, progress As System.IProgress(Of RedInkPythonAgentEvent), sessionId As System.Guid, userCancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of RedInkPythonAgentRunResult)
            ValidateInputs(configuration, rootPath, pythonCode, limits)
            Dim fullRoot As System.String = System.IO.Path.GetFullPath(rootPath)
            System.IO.Directory.CreateDirectory(fullRoot)
            Dim nonce As System.String = CreateNonce()
            Dim requestDirectory As System.String = System.IO.Path.Combine(fullRoot, "requests", sessionId.ToString("D"))
            Dim controlDirectory As System.String = System.IO.Path.Combine(fullRoot, "control", sessionId.ToString("D"))
            Dim requestPath As System.String = System.IO.Path.Combine(fullRoot, "request.json")
            Dim responsePath As System.String = System.IO.Path.Combine(fullRoot, "response.json")
            Dim process As System.Diagnostics.Process = Nothing
            Dim timeoutCancellation As New System.Threading.CancellationTokenSource(Me.Options.OverallTimeout)
            Dim linked As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(userCancellation, timeoutCancellation.Token)
            Dim operationCancelled As System.Boolean = False
            Try
                linked.Token.ThrowIfCancellationRequested()
                StageRequest(requestDirectory, requestPath, pythonCode, inputFiles, limits, sessionId, nonce, linked.Token)
                Dim executable As System.String = ResolveExecutable(configuration.ExecutablePath)
                Using heldExecutable As New System.IO.FileStream(executable, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                    VerifyConfiguration(configuration, executable, heldExecutable)
                    Dim startInfo As New System.Diagnostics.ProcessStartInfo() With {
                    .FileName = executable,
                    .Arguments = "--root """ & fullRoot & """",
                    .UseShellExecute = False,
                    .CreateNoWindow = True,
                    .WorkingDirectory = fullRoot,
                    .RedirectStandardOutput = False,
                    .RedirectStandardError = False
                }
                    process = System.Diagnostics.Process.Start(startInfo)
                    If process Is Nothing Then
                        Throw New RedInkPythonAgentExecutionException("BROKER_START_FAILED")
                    End If
                End Using

                Dim eventOffset As System.Int64 = 0
                Dim pending As New System.Collections.Generic.Dictionary(Of System.String, System.Threading.Tasks.Task)(System.StringComparer.OrdinalIgnoreCase)
                Dim lastHeartbeat As System.DateTimeOffset = System.DateTimeOffset.UtcNow
                Do
                    If linked.IsCancellationRequested Then
                        Await RequestCancellationAndKillAsync(process, controlDirectory, sessionId, nonce, If(timeoutCancellation.IsCancellationRequested, "SESSION_TIMEOUT", "HOST_CANCELLED"), pending).ConfigureAwait(False)
                        Return CreateCallerResult(sessionId, If(timeoutCancellation.IsCancellationRequested, "timeout", "cancelled"), If(timeoutCancellation.IsCancellationRequested, "SESSION_TIMEOUT", "SESSION_CANCELLED"))
                    End If
                    Dim exited As System.Boolean = False
                    Try
                        exited = process.HasExited
                    Catch ex As System.InvalidOperationException
                        Throw New RedInkPythonAgentExecutionException("BROKER_EXITED_WITHOUT_RESPONSE", ex)
                    End Try
                    eventOffset = ReadEvents(controlDirectory, sessionId, progress, eventOffset)
                    DispatchHostCalls(controlDirectory, sessionId, nonce, limits, hostServiceHandler, pending, linked.Token, Me.Options.HumanDiagnosticLog)
                    lastHeartbeat = ValidateHeartbeat(controlDirectory, sessionId, nonce, lastHeartbeat)
                    If System.DateTimeOffset.UtcNow - lastHeartbeat > Me.Options.HeartbeatTimeout Then
                        Await RequestCancellationAndKillAsync(process, controlDirectory, sessionId, nonce, "HOST_CANCELLED", pending).ConfigureAwait(False)
                        Return CreateCallerResult(sessionId, "failed", "BROKER_HEARTBEAT_LOST")
                    End If
                    If exited Then
                        Exit Do
                    End If
                    Await System.Threading.Tasks.Task.Delay(Me.Options.PollInterval, linked.Token).ConfigureAwait(False)
                Loop
                Await DrainPendingAsync(pending, Me.Options.CancellationGracePeriod).ConfigureAwait(False)
                If Not System.IO.File.Exists(responsePath) Then
                    Return CreateCallerResult(sessionId, "failed", "BROKER_EXITED_WITHOUT_RESPONSE")
                End If
                Return ParseResponse(responsePath, fullRoot, sessionId, nonce, limits)
            Catch ex As System.OperationCanceledException
                operationCancelled = True
            Finally
                linked.Dispose()
                timeoutCancellation.Dispose()
                If process IsNot Nothing AndAlso Not operationCancelled Then
                    process.Dispose()
                End If
            End Try

            If operationCancelled Then
                Try
                    If process IsNot Nothing Then
                        Await RequestCancellationAndKillAsync(process, controlDirectory, sessionId, nonce, "HOST_CANCELLED", New System.Collections.Generic.Dictionary(Of System.String, System.Threading.Tasks.Task)()).ConfigureAwait(False)
                    End If
                Finally
                    If process IsNot Nothing Then
                        process.Dispose()
                    End If
                End Try
                Return CreateCallerResult(sessionId, "cancelled", "SESSION_CANCELLED")
            End If

            Throw New RedInkPythonAgentExecutionException("INTERNAL_BROKER_ERROR")
        End Function

        Private Shared Sub StageRequest(requestDirectory As System.String, requestPath As System.String, pythonCode As System.String, inputFiles As System.Collections.Generic.IEnumerable(Of RedInkPythonAgentInputFile), limits As RedInkPythonAgentLimits, sessionId As System.Guid, nonce As System.String, cancellationToken As System.Threading.CancellationToken)
            System.IO.Directory.CreateDirectory(System.IO.Path.Combine(requestDirectory, "input"))
            cancellationToken.ThrowIfCancellationRequested()
            System.IO.File.WriteAllText(System.IO.Path.Combine(requestDirectory, "code.py"), pythonCode, New System.Text.UTF8Encoding(False, True))
            For Each item As RedInkPythonAgentInputFile In inputFiles
                cancellationToken.ThrowIfCancellationRequested()
                Dim relative As System.String = ValidateRelativePath(item.RelativePath)
                Dim destination As System.String = System.IO.Path.Combine(requestDirectory, "input", relative)
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(destination))
                Using source As New System.IO.FileStream(System.IO.Path.GetFullPath(item.SourcePath), System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                    Using target As New System.IO.FileStream(destination, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None)
                        Dim buffer(1048575) As System.Byte
                        Do
                            cancellationToken.ThrowIfCancellationRequested()
                            Dim read As System.Int32 = source.Read(buffer, 0, buffer.Length)
                            If read = 0 Then Exit Do
                            target.Write(buffer, 0, read)
                        Loop
                    End Using
                End Using
            Next
            Dim operations As New Newtonsoft.Json.Linq.JArray()
            For Each operation As System.String In limits.AllowedOperations
                operations.Add(operation)
            Next
            Dim request As New Newtonsoft.Json.Linq.JObject(
            New Newtonsoft.Json.Linq.JProperty("protocolVersion", ProtocolVersion),
            New Newtonsoft.Json.Linq.JProperty("sessionId", sessionId.ToString("D")),
            New Newtonsoft.Json.Linq.JProperty("nonce", nonce),
            New Newtonsoft.Json.Linq.JProperty("limits", New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("startupWallTimeSeconds", limits.StartupWallTimeSeconds),
                New Newtonsoft.Json.Linq.JProperty("executeWallTimeSeconds", limits.ExecuteWallTimeSeconds),
                New Newtonsoft.Json.Linq.JProperty("executionInactivitySeconds", limits.ExecutionInactivitySeconds),
                New Newtonsoft.Json.Linq.JProperty("validatorWallTimeSeconds", limits.ValidatorWallTimeSeconds),
                New Newtonsoft.Json.Linq.JProperty("overallWallTimeSeconds", limits.OverallWallTimeSeconds),
                New Newtonsoft.Json.Linq.JProperty("memoryMiB", limits.MemoryMiB),
                New Newtonsoft.Json.Linq.JProperty("maxOutputBytes", limits.MaxOutputBytes),
                New Newtonsoft.Json.Linq.JProperty("maxOutputFiles", limits.MaxOutputFiles),
                New Newtonsoft.Json.Linq.JProperty("maxWorkingBytes", limits.MaxWorkingBytes),
                New Newtonsoft.Json.Linq.JProperty("maxWorkingFiles", limits.MaxWorkingFiles),
                New Newtonsoft.Json.Linq.JProperty("maxResultBytes", limits.MaxResultBytes),
                New Newtonsoft.Json.Linq.JProperty("maxResultJsonDepth", limits.MaxResultJsonDepth),
                New Newtonsoft.Json.Linq.JProperty("maxResultJsonNodes", limits.MaxResultJsonNodes))),
            New Newtonsoft.Json.Linq.JProperty("hostServices", New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("allowedOperations", operations),
                New Newtonsoft.Json.Linq.JProperty("maximumCalls", If(operations.Count = 0, 0, limits.MaximumCalls)),
                New Newtonsoft.Json.Linq.JProperty("maximumConcurrentCalls", If(operations.Count = 0, 0, limits.MaximumConcurrentCalls)),
                New Newtonsoft.Json.Linq.JProperty("maximumRequestBytes", limits.MaximumRequestBytes),
                New Newtonsoft.Json.Linq.JProperty("maximumResponseBytes", limits.MaximumResponseBytes),
                New Newtonsoft.Json.Linq.JProperty("defaultCallTimeoutSeconds", limits.DefaultCallTimeoutSeconds),
                New Newtonsoft.Json.Linq.JProperty("maximumCallTimeoutSeconds", limits.MaximumCallTimeoutSeconds))))
            AtomicWrite(requestPath, request)
        End Sub

        Private Shared Sub DispatchHostCalls(controlDirectory As System.String, sessionId As System.Guid, nonce As System.String, limits As RedInkPythonAgentLimits, handler As IRedInkPythonAgentHostServiceHandler, pending As System.Collections.Generic.Dictionary(Of System.String, System.Threading.Tasks.Task), cancellationToken As System.Threading.CancellationToken, humanDiagnosticLog As System.Action(Of System.String))
            Dim requestDirectory As System.String = System.IO.Path.Combine(controlDirectory, "hostcalls", "requests")
            If Not System.IO.Directory.Exists(requestDirectory) Then Return
            For Each path As System.String In System.IO.Directory.GetFiles(requestDirectory, "*.request.json")
                Dim key As System.String = System.IO.Path.GetFileName(path)
                If pending.ContainsKey(key) Then Continue For
                Dim task As System.Threading.Tasks.Task = HandleHostCallAsync(path, System.IO.Path.Combine(controlDirectory, "hostcalls", "responses"), sessionId, nonce, limits, handler, cancellationToken, humanDiagnosticLog)
                pending.Add(key, task)
            Next
            Dim completed As New System.Collections.Generic.List(Of System.String)()
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Threading.Tasks.Task) In pending
                If pair.Value.IsCompleted Then completed.Add(pair.Key)
            Next
            For Each key As System.String In completed
                pending.Remove(key)
            Next
        End Sub

        Private Shared Async Function HandleHostCallAsync(path As System.String, responseDirectory As System.String, sessionId As System.Guid, nonce As System.String, limits As RedInkPythonAgentLimits, handler As IRedInkPythonAgentHostServiceHandler, cancellationToken As System.Threading.CancellationToken, humanDiagnosticLog As System.Action(Of System.String)) As System.Threading.Tasks.Task
            Dim requestId As System.Guid = System.Guid.Empty
            Dim operationText As System.String = System.String.Empty
            Dim response As RedInkPythonAgentHostCallResponse = Nothing
            Try
                Dim requestInformation As New System.IO.FileInfo(path)
                If requestInformation.Length < 1L OrElse requestInformation.Length > System.Convert.ToInt64(limits.MaximumRequestBytes) + 4096L Then Throw New RedInkPythonAgentExecutionException("HOST_CALL_REQUEST_INVALID")
                Dim obj As Newtonsoft.Json.Linq.JObject = ParseStrictObject(System.IO.File.ReadAllText(path, New System.Text.UTF8Encoding(False, True)))
                RequireFields(obj, New System.String() {"protocolVersion", "sessionId", "nonce", "requestId", "operation", "timeoutSeconds", "arguments"})
                If CInt(obj("protocolVersion")) <> 2 OrElse Not System.String.Equals(CStr(obj("sessionId")), sessionId.ToString("D"), System.StringComparison.Ordinal) OrElse Not FixedTimeEquals(CStr(obj("nonce")), nonce) Then Throw New RedInkPythonAgentExecutionException("HOST_CALL_REQUEST_INVALID")
                requestId = System.Guid.ParseExact(CStr(obj("requestId")), "D")
                operationText = CStr(obj("operation"))
                Dim timeoutSeconds As System.Int32 = CInt(obj("timeoutSeconds"))
                If timeoutSeconds < 1 OrElse timeoutSeconds > limits.MaximumCallTimeoutSeconds Then Throw New RedInkPythonAgentExecutionException("HOST_CALL_REQUEST_INVALID")
                If Not limits.AllowedOperations.Contains(operationText) Then
                    response = RedInkPythonAgentHostCallResponse.Failure("HOST_OPERATION_NOT_ALLOWED", False, Nothing)
                End If
                Dim arguments As Newtonsoft.Json.Linq.JObject = CType(obj("arguments"), Newtonsoft.Json.Linq.JObject)
                Dim request As New RedInkPythonAgentHostCallRequest() With {.RequestId = requestId, .Timeout = System.TimeSpan.FromSeconds(timeoutSeconds)}
                Select Case operationText
                    Case "llm.complete"
                        RequireFields(arguments, New System.String() {"systemPrompt", "userPrompt"})
                        request.Operation = RedInkPythonAgentHostOperation.LlmComplete
                        request.Arguments = New RedInkPythonAgentLlmRequest() With {.SystemPrompt = CStr(arguments("systemPrompt")), .UserPrompt = CStr(arguments("userPrompt"))}
                    Case "web.get"
                        RequireFields(arguments, New System.String() {"url", "maximumCharacters"})
                        request.Operation = RedInkPythonAgentHostOperation.WebGet
                        request.Arguments = New RedInkPythonAgentWebGetRequest() With {.Url = CStr(arguments("url")), .MaximumCharacters = CInt(arguments("maximumCharacters"))}
                    Case "web.search"
                        RequireFields(arguments, New System.String() {"query", "maximumResults"})
                        request.Operation = RedInkPythonAgentHostOperation.WebSearch
                        request.Arguments = New RedInkPythonAgentWebSearchRequest() With {.Query = CStr(arguments("query")), .MaximumResults = CInt(arguments("maximumResults"))}
                    Case Else
                        response = RedInkPythonAgentHostCallResponse.Failure("HOST_OPERATION_NOT_ALLOWED", False, Nothing)
                End Select
                If response Is Nothing Then
                    Using callCancellation As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                        callCancellation.CancelAfter(System.TimeSpan.FromSeconds(timeoutSeconds))
                        response = Await handler.HandleAsync(request, callCancellation.Token).ConfigureAwait(False)
                    End Using
                End If
            Catch ex As System.OperationCanceledException
                response = RedInkPythonAgentHostCallResponse.Failure("HOST_CALL_CANCELLED", True, Nothing)
            Catch ex As System.Exception
                ReportHumanDiagnostic(humanDiagnosticLog, ex.ToString())
                response = RedInkPythonAgentHostCallResponse.Failure("HOST_CALL_FAILED", False, Nothing)
            End Try
            If requestId = System.Guid.Empty Then Return
            Dim result As Newtonsoft.Json.Linq.JObject
            If response IsNot Nothing AndAlso response.IsSuccess Then
                Dim resultValue As Newtonsoft.Json.Linq.JObject
                If operationText = "web.search" Then
                    Dim array As New Newtonsoft.Json.Linq.JArray()
                    If response.SearchResults IsNot Nothing Then
                        For Each item As RedInkPythonAgentWebSearchItem In response.SearchResults
                            array.Add(New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("title", item.Title), New Newtonsoft.Json.Linq.JProperty("url", item.Url), New Newtonsoft.Json.Linq.JProperty("snippet", item.Snippet)))
                        Next
                    End If
                    resultValue = New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("results", array))
                Else
                    resultValue = New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("text", If(response.Text, System.String.Empty)))
                End If
                result = New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("protocolVersion", 2), New Newtonsoft.Json.Linq.JProperty("sessionId", sessionId.ToString("D")), New Newtonsoft.Json.Linq.JProperty("nonce", nonce), New Newtonsoft.Json.Linq.JProperty("requestId", requestId.ToString("D")), New Newtonsoft.Json.Linq.JProperty("status", "success"), New Newtonsoft.Json.Linq.JProperty("result", resultValue))
            Else
                Dim code As System.String = If(response Is Nothing OrElse Not IsSafeHostCallErrorCode(response.ErrorCode), "HOST_CALL_FAILED", response.ErrorCode)
                result = New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("protocolVersion", 2), New Newtonsoft.Json.Linq.JProperty("sessionId", sessionId.ToString("D")), New Newtonsoft.Json.Linq.JProperty("nonce", nonce), New Newtonsoft.Json.Linq.JProperty("requestId", requestId.ToString("D")), New Newtonsoft.Json.Linq.JProperty("status", "failed"), New Newtonsoft.Json.Linq.JProperty("error", New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("code", code), New Newtonsoft.Json.Linq.JProperty("retryable", response IsNot Nothing AndAlso response.Retryable), New Newtonsoft.Json.Linq.JProperty("retryAfterMilliseconds", If(response Is Nothing, Nothing, response.RetryAfterMilliseconds)))))
            End If
            Dim encoded As System.String = result.ToString(Newtonsoft.Json.Formatting.None)
            If System.Text.Encoding.UTF8.GetByteCount(encoded) > limits.MaximumResponseBytes Then
                Dim oversizedCode As System.String = If(operationText = "web.get" OrElse operationText = "web.search", "WEB_RESPONSE_TOO_LARGE", "LLM_RESPONSE_INVALID")
                result = New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("protocolVersion", 2),
                New Newtonsoft.Json.Linq.JProperty("sessionId", sessionId.ToString("D")),
                New Newtonsoft.Json.Linq.JProperty("nonce", nonce),
                New Newtonsoft.Json.Linq.JProperty("requestId", requestId.ToString("D")),
                New Newtonsoft.Json.Linq.JProperty("status", "failed"),
                New Newtonsoft.Json.Linq.JProperty("error", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("code", oversizedCode),
                    New Newtonsoft.Json.Linq.JProperty("retryable", False),
                    New Newtonsoft.Json.Linq.JProperty("retryAfterMilliseconds", Newtonsoft.Json.Linq.JValue.CreateNull()))))
            End If
            System.IO.Directory.CreateDirectory(responseDirectory)
            AtomicWriteBounded(System.IO.Path.Combine(responseDirectory, requestId.ToString("D") & ".response.json"), result, limits.MaximumResponseBytes)
        End Function

        Private Shared Function IsSafeHostCallErrorCode(value As System.String) As System.Boolean
            Select Case value
                Case "HOST_OPERATION_NOT_ALLOWED", "HOST_CALL_LIMIT_EXCEEDED", "HOST_CALL_REQUEST_INVALID", "HOST_CALL_TIMEOUT", "HOST_CALL_CANCELLED", "HOST_CALL_FAILED", "LLM_AUTHENTICATION_FAILED", "LLM_RATE_LIMITED", "LLM_PROVIDER_UNAVAILABLE", "LLM_RESPONSE_INVALID", "WEB_ACCESS_DENIED", "WEB_REQUEST_TIMEOUT", "WEB_RESPONSE_TOO_LARGE"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Sub AtomicWriteBounded(path As System.String, value As Newtonsoft.Json.Linq.JObject, maximumBytes As System.Int32)
            Dim encoded As System.String = value.ToString(Newtonsoft.Json.Formatting.None)
            If System.Text.Encoding.UTF8.GetByteCount(encoded) > maximumBytes Then Throw New RedInkPythonAgentExecutionException("HOST_CALL_FAILED")
            Dim directory As System.String = System.IO.Path.GetDirectoryName(path)
            If Not System.String.IsNullOrEmpty(directory) Then System.IO.Directory.CreateDirectory(directory)
            Dim temporary As System.String = path & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
            System.IO.File.WriteAllText(temporary, encoded, New System.Text.UTF8Encoding(False, True))
            System.IO.File.Move(temporary, path)
        End Sub

        Private Shared Function ReadEvents(controlDirectory As System.String, sessionId As System.Guid, progress As System.IProgress(Of RedInkPythonAgentEvent), offset As System.Int64) As System.Int64
            Dim path As System.String = System.IO.Path.Combine(controlDirectory, "events.ndjson")
            If Not System.IO.File.Exists(path) Then Return offset
            Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite Or System.IO.FileShare.Delete)
                If offset > stream.Length Then offset = 0
                stream.Position = offset
                Using reader As New System.IO.StreamReader(stream, New System.Text.UTF8Encoding(False, True), False, 4096, True)
                    Do While Not reader.EndOfStream
                        Dim line As System.String = reader.ReadLine()
                        If System.String.IsNullOrWhiteSpace(line) Then Continue Do
                        Try
                            Dim obj As Newtonsoft.Json.Linq.JObject = ParseStrictObject(line)
                            Dim value As New RedInkPythonAgentEvent() With {.Sequence = CLng(obj("sequence")), .EventCode = CStr(obj("eventCode")), .Phase = CStr(obj("phase")), .StepId = If(obj("stepId") Is Nothing OrElse obj("stepId").Type = Newtonsoft.Json.Linq.JTokenType.Null, Nothing, CStr(obj("stepId"))), .HostOperation = If(obj("hostOperation") Is Nothing OrElse obj("hostOperation").Type = Newtonsoft.Json.Linq.JTokenType.Null, Nothing, CStr(obj("hostOperation")))}
                            If progress IsNot Nothing Then
                                Try
                                    progress.Report(value)
                                Catch ex As System.Exception
                                End Try
                            End If
                        Catch ex As System.Exception
                        End Try
                    Loop
                End Using
                Return stream.Position
            End Using
        End Function

        Private Shared Function ValidateHeartbeat(controlDirectory As System.String, sessionId As System.Guid, nonce As System.String, previous As System.DateTimeOffset) As System.DateTimeOffset
            Dim path As System.String = System.IO.Path.Combine(controlDirectory, "heartbeat.json")
            If Not System.IO.File.Exists(path) Then Return previous
            Try
                Dim obj As Newtonsoft.Json.Linq.JObject = ParseStrictObject(System.IO.File.ReadAllText(path, New System.Text.UTF8Encoding(False, True)))
                RequireFields(obj, New System.String() {"protocolVersion", "sessionId", "nonce", "sequence", "phase", "writtenUtc", "mainActivityAgeMilliseconds"})
                If CInt(obj("protocolVersion")) <> 2 OrElse CStr(obj("sessionId")) <> sessionId.ToString("D") OrElse Not FixedTimeEquals(CStr(obj("nonce")), nonce) Then Return previous
                Dim written As System.DateTimeOffset
                If System.DateTimeOffset.TryParse(CStr(obj("writtenUtc")), written) Then Return written.ToUniversalTime()
            Catch ex As System.Exception
            End Try
            Return previous
        End Function

        Private Async Function RequestCancellationAndKillAsync(process As System.Diagnostics.Process, controlDirectory As System.String, sessionId As System.Guid, nonce As System.String, reason As System.String, pending As System.Collections.Generic.Dictionary(Of System.String, System.Threading.Tasks.Task)) As System.Threading.Tasks.Task
            System.IO.Directory.CreateDirectory(controlDirectory)
            AtomicWrite(System.IO.Path.Combine(controlDirectory, "cancel.request.json"), New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("protocolVersion", 2), New Newtonsoft.Json.Linq.JProperty("sessionId", sessionId.ToString("D")), New Newtonsoft.Json.Linq.JProperty("nonce", nonce), New Newtonsoft.Json.Linq.JProperty("reasonCode", "HOST_CANCELLED")))
            Await DrainPendingAsync(pending, Me.Options.CancellationGracePeriod).ConfigureAwait(False)
            If Not process.HasExited Then
                process.Kill()
                Dim deadline As System.DateTimeOffset = System.DateTimeOffset.UtcNow + Me.Options.HardKillWait
                Do While Not process.HasExited AndAlso System.DateTimeOffset.UtcNow < deadline
                    Await System.Threading.Tasks.Task.Delay(100).ConfigureAwait(False)
                Loop
            End If
        End Function

        Private Shared Async Function DrainPendingAsync(pending As System.Collections.Generic.Dictionary(Of System.String, System.Threading.Tasks.Task), timeout As System.TimeSpan) As System.Threading.Tasks.Task
            If pending.Count = 0 Then Return
            Dim all As System.Threading.Tasks.Task = System.Threading.Tasks.Task.WhenAll(pending.Values)
            Await System.Threading.Tasks.Task.WhenAny(all, System.Threading.Tasks.Task.Delay(timeout)).ConfigureAwait(False)
        End Function

        Private Shared Function ParseResponse(path As System.String, root As System.String, sessionId As System.Guid, nonce As System.String, limits As RedInkPythonAgentLimits) As RedInkPythonAgentRunResult
            Dim maximumFileBytes As System.Int64 = limits.MaxResultBytes + 4194304L
            Dim information As New System.IO.FileInfo(path)
            If information.Length < 1L OrElse information.Length > maximumFileBytes Then
                Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
            End If
            Dim obj As Newtonsoft.Json.Linq.JObject = ParseStrictObject(
            System.IO.File.ReadAllText(path, New System.Text.UTF8Encoding(False, True)),
            System.Math.Min(256, limits.MaxResultJsonDepth + 8))
            RequireFields(obj, New System.String() {"protocolVersion", "sessionId", "nonce", "status", "exitCode", "completedUtc", "durationMilliseconds", "diagnosticId", "humanLogAvailable", "outputs", "result", "error"})
            Dim result As New RedInkPythonAgentRunResult() With {
            .ProtocolVersion = CInt(obj("protocolVersion")),
            .SessionId = System.Guid.Parse(CStr(obj("sessionId"))),
            .Status = CStr(obj("status")),
            .ExitCode = CInt(obj("exitCode")),
            .DiagnosticId = System.Guid.Parse(CStr(obj("diagnosticId"))),
            .HumanLogAvailable = CBool(obj("humanLogAvailable"))
        }
            If result.ProtocolVersion <> 2 OrElse result.SessionId <> sessionId OrElse Not FixedTimeEquals(CStr(obj("nonce")), nonce) Then
                Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
            End If
            If Not IsSafeStatus(result.Status) Then
                Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
            End If
            Dim outputs As Newtonsoft.Json.Linq.JArray = TryCast(obj("outputs"), Newtonsoft.Json.Linq.JArray)
            If outputs Is Nothing Then
                Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
            End If
            For Each token As Newtonsoft.Json.Linq.JToken In outputs
                Dim item As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
                If item Is Nothing Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
                RequireFields(item, New System.String() {"relativePath", "mediaType", "size", "sha256"})
                Dim relative As System.String = CStr(item("relativePath"))
                result.Outputs.Add(New RedInkPythonAgentOutput() With {
                .RelativePath = relative,
                .MediaType = CStr(item("mediaType")),
                .Size = CLng(item("size")),
                .Sha256 = CStr(item("sha256")),
                .FullPath = System.IO.Path.GetFullPath(System.IO.Path.Combine(root, relative))
            })
            Next
            Dim publishedToken As Newtonsoft.Json.Linq.JToken = obj("result")
            If publishedToken IsNot Nothing AndAlso publishedToken.Type <> Newtonsoft.Json.Linq.JTokenType.Null Then
                Dim publishedObject As Newtonsoft.Json.Linq.JObject = TryCast(publishedToken, Newtonsoft.Json.Linq.JObject)
                If publishedObject Is Nothing Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
                RequireFields(publishedObject, New System.String() {"kind", "value"})
                Dim kind As System.String = CStr(publishedObject("kind"))
                If Not System.String.Equals(kind, "json", System.StringComparison.Ordinal) AndAlso Not System.String.Equals(kind, "text", System.StringComparison.Ordinal) Then
                    Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
                End If
                If System.String.Equals(kind, "text", System.StringComparison.Ordinal) AndAlso publishedObject("value").Type <> Newtonsoft.Json.Linq.JTokenType.String Then
                    Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
                End If
                ValidatePublishedResultValue(publishedObject("value"), limits.MaxResultJsonDepth, limits.MaxResultJsonNodes)
                result.Result = New RedInkPythonAgentPublishedResult() With {
                .Kind = kind,
                .Value = publishedObject("value").DeepClone()
            }
            End If
            Dim errorToken As Newtonsoft.Json.Linq.JToken = obj("error")
            If errorToken IsNot Nothing AndAlso errorToken.Type <> Newtonsoft.Json.Linq.JTokenType.Null Then
                result.Error = ParseSafeError(errorToken)
            End If
            Return result
        End Function

        Private Shared Sub ValidatePublishedResultValue(value As Newtonsoft.Json.Linq.JToken, maximumDepth As System.Int32, maximumNodes As System.Int32)
            Dim stack As New System.Collections.Generic.Stack(Of System.Tuple(Of Newtonsoft.Json.Linq.JToken, System.Int32))()
            stack.Push(System.Tuple.Create(value, 1))
            Dim nodes As System.Int32 = 0
            Do While stack.Count > 0
                Dim current As System.Tuple(Of Newtonsoft.Json.Linq.JToken, System.Int32) = stack.Pop()
                nodes += 1
                If nodes > maximumNodes OrElse current.Item2 > maximumDepth Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
                Dim propertyValue As Newtonsoft.Json.Linq.JProperty = TryCast(current.Item1, Newtonsoft.Json.Linq.JProperty)
                If propertyValue IsNot Nothing Then
                    stack.Push(System.Tuple.Create(propertyValue.Value, current.Item2))
                    Continue Do
                End If
                Dim container As Newtonsoft.Json.Linq.JContainer = TryCast(current.Item1, Newtonsoft.Json.Linq.JContainer)
                If container IsNot Nothing Then
                    For Each child As Newtonsoft.Json.Linq.JToken In container.Children()
                        stack.Push(System.Tuple.Create(child, current.Item2 + 1))
                    Next
                End If
            Loop
        End Sub

        Private Shared Function ParseSafeError(token As Newtonsoft.Json.Linq.JToken) As RedInkPythonAgentSafeError
            Dim value As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
            If value Is Nothing Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
            RequireFields(value, New System.String() {"code", "phase", "retryable", "source", "hostOperation", "limit", "observed", "stack"})
            Dim code As System.String = CStr(value("code"))
            Dim phase As System.String = CStr(value("phase"))
            If Not IsSafeErrorCode(code) OrElse Not IsSafePhase(phase) Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
            Dim result As New RedInkPythonAgentSafeError() With {
            .Code = code,
            .Phase = phase,
            .Retryable = CBool(value("retryable")),
            .HostOperation = If(value("hostOperation").Type = Newtonsoft.Json.Linq.JTokenType.Null, Nothing, CStr(value("hostOperation"))),
            .Limit = ReadOptionalInt64(value("limit")),
            .Observed = ReadOptionalInt64(value("observed"))
        }
            If result.HostOperation IsNot Nothing AndAlso result.HostOperation <> "llm.complete" AndAlso result.HostOperation <> "web.get" AndAlso result.HostOperation <> "web.search" Then
                Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
            End If
            If value("source").Type <> Newtonsoft.Json.Linq.JTokenType.Null Then
                Dim source As Newtonsoft.Json.Linq.JObject = TryCast(value("source"), Newtonsoft.Json.Linq.JObject)
                If source Is Nothing Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
                RequireFields(source, New System.String() {"file", "line", "column", "function", "symbol"})
                If CStr(source("file")) <> "code.py" Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
                result.Source = New RedInkPythonAgentSafeSource() With {
                .File = "code.py",
                .Line = CInt(source("line")),
                .Column = ReadOptionalInt32(source("column")),
                .Function = ReadOptionalIdentifier(source("function")),
                .Symbol = ReadOptionalIdentifier(source("symbol"))
            }
                If result.Source.Line < 1 OrElse (result.Source.Column.HasValue AndAlso result.Source.Column.Value < 1) Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
            End If
            Dim stack As Newtonsoft.Json.Linq.JArray = TryCast(value("stack"), Newtonsoft.Json.Linq.JArray)
            If stack Is Nothing OrElse stack.Count > 20 Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
            For Each frameToken As Newtonsoft.Json.Linq.JToken In stack
                Dim frame As Newtonsoft.Json.Linq.JObject = TryCast(frameToken, Newtonsoft.Json.Linq.JObject)
                If frame Is Nothing Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
                RequireFields(frame, New System.String() {"file", "line", "function"})
                If CStr(frame("file")) <> "code.py" OrElse CInt(frame("line")) < 1 Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
                result.Stack.Add(New RedInkPythonAgentSafeStackFrame() With {
                .File = "code.py",
                .Line = CInt(frame("line")),
                .Function = ReadOptionalIdentifier(frame("function"))
            })
            Next
            Return result
        End Function

        Private Shared Function ReadOptionalInt64(token As Newtonsoft.Json.Linq.JToken) As System.Nullable(Of System.Int64)
            If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return Nothing
            Dim value As System.Int64 = CLng(token)
            If value < 0L Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
            Return value
        End Function

        Private Shared Function ReadOptionalInt32(token As Newtonsoft.Json.Linq.JToken) As System.Nullable(Of System.Int32)
            If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return Nothing
            Return CInt(token)
        End Function

        Private Shared Function ReadOptionalIdentifier(token As Newtonsoft.Json.Linq.JToken) As System.String
            If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return Nothing
            Dim value As System.String = CStr(token)
            If Not System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-Za-z_][A-Za-z0-9_]{0,127}$", System.Text.RegularExpressions.RegexOptions.CultureInvariant) Then
                Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
            End If
            Return value
        End Function

        Private Shared Sub ReportHumanDiagnostic(callback As System.Action(Of System.String), message As System.String)
            Try
                If callback IsNot Nothing Then
                    callback(message)
                Else
                    System.Diagnostics.Trace.WriteLine(message)
                End If
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try
        End Sub

        Private Shared Function IsSafeStatus(value As System.String) As System.Boolean
            Return value = "success" OrElse value = "failed" OrElse value = "cancelled" OrElse value = "timeout"
        End Function

        Private Shared Function IsSafePhase(value As System.String) As System.Boolean
            Select Case value
                Case "initializing", "request_validation", "input_staging", "runtime_verification", "execute_startup", "execute", "host_call", "validation", "publication", "cleanup", "completed"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function IsSafeErrorCode(value As System.String) As System.Boolean
            Select Case value
                Case "CONFIGURATION_INVALID", "EXECUTABLE_NOT_FOUND", "EXECUTABLE_HASH_MISMATCH", "EXECUTABLE_SIGNATURE_INVALID", "EXECUTABLE_SIGNER_MISMATCH", "ROOT_BUSY", "REQUEST_INVALID", "SECURITY_INVARIANT_FAILED", "BROKER_START_FAILED", "BROKER_EXITED_WITHOUT_RESPONSE", "BROKER_HEARTBEAT_LOST", "SESSION_TIMEOUT", "SESSION_CANCELLED", "WORKER_TIMEOUT", "WORKER_MEMORY_LIMIT", "WORKER_PROCESS_LIMIT", "WORKER_STREAM_LIMIT", "WORKER_WORKING_BYTES_LIMIT", "WORKER_WORKING_FILES_LIMIT", "PYTHON_SYNTAX_ERROR", "PYTHON_NAME_ERROR", "PYTHON_IMPORT_ERROR", "PYTHON_ATTRIBUTE_ERROR", "PYTHON_TYPE_ERROR", "PYTHON_VALUE_ERROR", "PYTHON_KEY_ERROR", "PYTHON_INDEX_ERROR", "PYTHON_FILE_NOT_FOUND", "PYTHON_PERMISSION_ERROR", "PYTHON_RUNTIME_ERROR", "PYTHON_RESULT_INVALID", "PYTHON_RESULT_TOO_LARGE", "PYTHON_RESULT_TOO_DEEP", "PYTHON_RESULT_TOO_COMPLEX", "HOST_OPERATION_NOT_ALLOWED", "HOST_CALL_LIMIT_EXCEEDED", "HOST_CALL_REQUEST_INVALID", "HOST_CALL_TIMEOUT", "HOST_CALL_CANCELLED", "HOST_CALL_FAILED", "LLM_AUTHENTICATION_FAILED", "LLM_RATE_LIMITED", "LLM_PROVIDER_UNAVAILABLE", "LLM_RESPONSE_INVALID", "WEB_ACCESS_DENIED", "WEB_REQUEST_TIMEOUT", "WEB_RESPONSE_TOO_LARGE", "OUTPUT_VALIDATION_FAILED", "OUTPUT_PUBLICATION_FAILED", "INTERNAL_BROKER_ERROR"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function CreateCallerResult(sessionId As System.Guid, status As System.String, code As System.String) As RedInkPythonAgentRunResult
            Return New RedInkPythonAgentRunResult() With {.ProtocolVersion = 2, .SessionId = sessionId, .Status = status, .ExitCode = 1, .DiagnosticId = System.Guid.NewGuid(), .HumanLogAvailable = True, .Error = New RedInkPythonAgentSafeError() With {.Code = code, .Phase = "cleanup", .Retryable = False}}
        End Function

        Private Shared Sub VerifyConfiguration(configuration As RedInkPythonAgentConfiguration, executable As System.String, stream As System.IO.FileStream)
            If System.String.IsNullOrWhiteSpace(configuration.ExpectedSignerOrganization) AndAlso System.String.IsNullOrWhiteSpace(configuration.ExpectedSha256) Then Throw New RedInkPythonAgentConfigurationException("At least one executable trust criterion is required.")
            If Not System.String.IsNullOrWhiteSpace(configuration.ExpectedSha256) Then
                Dim actual As System.String = ComputeSha256(stream)
                If Not FixedTimeEquals(actual, NormalizeSha256(configuration.ExpectedSha256)) Then Throw New RedInkPythonAgentExecutableTrustException("EXECUTABLE_HASH_MISMATCH", "Executable SHA-256 mismatch.")
            End If
            If Not System.String.IsNullOrWhiteSpace(configuration.ExpectedSignerOrganization) Then
                Dim verified As RedInkAuthenticodeVerificationResult = RedInkAuthenticodeVerifier.Verify(executable)
                If Not System.StringComparer.OrdinalIgnoreCase.Equals(NormalizeSigner(verified.SignerOrganization), NormalizeSigner(configuration.ExpectedSignerOrganization)) Then Throw New RedInkPythonAgentExecutableTrustException("EXECUTABLE_SIGNER_MISMATCH", "Executable signer organization mismatch.")
            End If
        End Sub

        Private Shared Function ResolveExecutable(path As System.String) As System.String
            Dim expanded As System.String = System.Environment.ExpandEnvironmentVariables(path)
            Dim full As System.String = System.IO.Path.GetFullPath(expanded)
            If Not full.EndsWith(".exe", System.StringComparison.OrdinalIgnoreCase) OrElse Not System.IO.File.Exists(full) Then Throw New RedInkPythonAgentExecutableTrustException("EXECUTABLE_NOT_FOUND", "Executable was not found.")
            If (System.IO.File.GetAttributes(full) And System.IO.FileAttributes.ReparsePoint) <> 0 Then Throw New RedInkPythonAgentExecutableTrustException("EXECUTABLE_SIGNATURE_INVALID", "Executable must not be a reparse point.")
            Dim root As System.String = System.IO.Path.GetPathRoot(full)
            Dim drive As New System.IO.DriveInfo(root)
            If drive.DriveType <> System.IO.DriveType.Fixed Then Throw New RedInkPythonAgentExecutableTrustException("EXECUTABLE_SIGNATURE_INVALID", "Executable must be on a local fixed drive.")
            Return full
        End Function

        Public Shared Function ComputeSha256(stream As System.IO.Stream) As System.String
            stream.Position = 0
            Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Dim digest As System.Byte() = algorithm.ComputeHash(stream)
                Dim builder As New System.Text.StringBuilder(digest.Length * 2)
                For Each value As System.Byte In digest
                    builder.Append(value.ToString("x2", System.Globalization.CultureInfo.InvariantCulture))
                Next
                stream.Position = 0
                Return builder.ToString()
            End Using
        End Function

        Private Shared Sub ValidateInputs(configuration As RedInkPythonAgentConfiguration, root As System.String, code As System.String, limits As RedInkPythonAgentLimits)
            If configuration Is Nothing OrElse limits Is Nothing Then Throw New System.ArgumentNullException()
            If System.String.IsNullOrWhiteSpace(root) OrElse System.String.IsNullOrWhiteSpace(code) Then Throw New RedInkPythonAgentConfigurationException("Root and code are required.")
            If limits.StartupWallTimeSeconds < 1 OrElse limits.StartupWallTimeSeconds > 600 Then Throw New RedInkPythonAgentConfigurationException("StartupWallTimeSeconds is outside the supported range.")
            If limits.ExecuteWallTimeSeconds < 1 OrElse limits.ExecuteWallTimeSeconds > 14400 Then Throw New RedInkPythonAgentConfigurationException("ExecuteWallTimeSeconds is outside the supported range.")
            If limits.ExecutionInactivitySeconds < 1 OrElse limits.ExecutionInactivitySeconds > 7200 Then Throw New RedInkPythonAgentConfigurationException("ExecutionInactivitySeconds is outside the supported range.")
            If limits.ValidatorWallTimeSeconds < 1 OrElse limits.ValidatorWallTimeSeconds > 1800 Then Throw New RedInkPythonAgentConfigurationException("ValidatorWallTimeSeconds is outside the supported range.")
            If limits.OverallWallTimeSeconds < 1 OrElse limits.OverallWallTimeSeconds > 18000 Then Throw New RedInkPythonAgentConfigurationException("OverallWallTimeSeconds is outside the supported range.")
            If limits.OverallWallTimeSeconds < limits.StartupWallTimeSeconds + limits.ExecuteWallTimeSeconds + limits.ValidatorWallTimeSeconds Then Throw New RedInkPythonAgentConfigurationException("Overall timeout must cover startup, execute and validation timeouts.")
            If limits.MaxResultBytes < 1L OrElse limits.MaxResultBytes > 134217728L Then Throw New RedInkPythonAgentConfigurationException("MaxResultBytes is outside the supported range.")
            If limits.MaxResultJsonDepth < 1 OrElse limits.MaxResultJsonDepth > 256 Then Throw New RedInkPythonAgentConfigurationException("MaxResultJsonDepth is outside the supported range.")
            If limits.MaxResultJsonNodes < 1 OrElse limits.MaxResultJsonNodes > 4000000 Then Throw New RedInkPythonAgentConfigurationException("MaxResultJsonNodes is outside the supported range.")
        End Sub
        Private Shared Sub ValidateOptions(options As RedInkPythonAgentExecutionOptions)
            If options.OverallTimeout <= System.TimeSpan.Zero OrElse options.HeartbeatTimeout <= System.TimeSpan.Zero OrElse options.CancellationGracePeriod <= System.TimeSpan.Zero OrElse options.HardKillWait <= System.TimeSpan.Zero Then Throw New RedInkPythonAgentConfigurationException("Timeouts must be positive and finite.")
            If options.PollInterval < System.TimeSpan.FromMilliseconds(50) OrElse options.PollInterval > System.TimeSpan.FromMilliseconds(1000) Then Throw New RedInkPythonAgentConfigurationException("PollInterval must be 50-1000 ms.")
        End Sub
        Private Shared Function CreateNonce() As System.String
            Dim bytes(31) As System.Byte
            Using random As System.Security.Cryptography.RandomNumberGenerator = System.Security.Cryptography.RandomNumberGenerator.Create()
                random.GetBytes(bytes)
            End Using
            Return System.BitConverter.ToString(bytes).Replace("-", System.String.Empty).ToLowerInvariant()
        End Function
        Private Shared Function ValidateRelativePath(value As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(value) OrElse System.IO.Path.IsPathRooted(value) OrElse value.Contains("..") Then Throw New RedInkPythonAgentConfigurationException("Unsafe relative path.")
            Return value.Replace("/"c, System.IO.Path.DirectorySeparatorChar)
        End Function
        Private Shared Sub AtomicWrite(path As System.String, value As Newtonsoft.Json.Linq.JObject)
            Dim directory As System.String = System.IO.Path.GetDirectoryName(path)
            If Not System.String.IsNullOrEmpty(directory) Then System.IO.Directory.CreateDirectory(directory)
            Dim temporary As System.String = path & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
            System.IO.File.WriteAllText(temporary, value.ToString(Newtonsoft.Json.Formatting.None), New System.Text.UTF8Encoding(False, True))
            System.IO.File.Move(temporary, path)
        End Sub
        Private Shared Function ParseStrictObject(text As System.String, Optional maximumDepth As System.Int32 = 256) As Newtonsoft.Json.Linq.JObject
            Using reader As New Newtonsoft.Json.JsonTextReader(New System.IO.StringReader(text))
                reader.DateParseHandling = Newtonsoft.Json.DateParseHandling.None
                reader.MaxDepth = maximumDepth
                Dim settings As New Newtonsoft.Json.Linq.JsonLoadSettings() With {
                .CommentHandling = Newtonsoft.Json.Linq.CommentHandling.Load,
                .DuplicatePropertyNameHandling = Newtonsoft.Json.Linq.DuplicatePropertyNameHandling.Error,
                .LineInfoHandling = Newtonsoft.Json.Linq.LineInfoHandling.Ignore
            }
                Dim token As Newtonsoft.Json.Linq.JToken = Newtonsoft.Json.Linq.JToken.ReadFrom(reader, settings)
                If reader.Read() Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
                ValidateStrictJsonTokens(token)
                Dim obj As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
                If obj Is Nothing Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
                Return obj
            End Using
        End Function
        Private Shared Sub ValidateStrictJsonTokens(root As Newtonsoft.Json.Linq.JToken)
            Dim stack As New System.Collections.Generic.Stack(Of Newtonsoft.Json.Linq.JToken)()
            stack.Push(root)
            Do While stack.Count > 0
                Dim token As Newtonsoft.Json.Linq.JToken = stack.Pop()
                Select Case token.Type
                    Case Newtonsoft.Json.Linq.JTokenType.Object,
                     Newtonsoft.Json.Linq.JTokenType.Array,
                     Newtonsoft.Json.Linq.JTokenType.Property
                        Dim container As Newtonsoft.Json.Linq.JContainer = DirectCast(token, Newtonsoft.Json.Linq.JContainer)
                        For Each child As Newtonsoft.Json.Linq.JToken In container.Children()
                            stack.Push(child)
                        Next
                    Case Newtonsoft.Json.Linq.JTokenType.Integer,
                     Newtonsoft.Json.Linq.JTokenType.String,
                     Newtonsoft.Json.Linq.JTokenType.Boolean,
                     Newtonsoft.Json.Linq.JTokenType.Null
                        Continue Do
                    Case Newtonsoft.Json.Linq.JTokenType.Float
                        Dim number As System.Double = System.Convert.ToDouble(DirectCast(token, Newtonsoft.Json.Linq.JValue).Value, System.Globalization.CultureInfo.InvariantCulture)
                        If System.Double.IsNaN(number) OrElse System.Double.IsInfinity(number) Then
                            Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
                        End If
                    Case Else
                        Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
                End Select
            Loop
        End Sub

        Private Shared Sub RequireFields(obj As Newtonsoft.Json.Linq.JObject, expected As System.Collections.Generic.IEnumerable(Of System.String))
            Dim setValue As New System.Collections.Generic.HashSet(Of System.String)(expected, System.StringComparer.Ordinal)
            For Each propertyValue As Newtonsoft.Json.Linq.JProperty In obj.Properties()
                If Not setValue.Remove(propertyValue.Name) Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
            Next
            If setValue.Count <> 0 Then Throw New RedInkPythonAgentExecutionException("REQUEST_INVALID")
        End Sub
        Private Shared Function FixedTimeEquals(left As System.String, right As System.String) As System.Boolean
            If left Is Nothing OrElse right Is Nothing OrElse left.Length <> right.Length Then Return False
            Dim difference As System.Int32 = 0
            For index As System.Int32 = 0 To left.Length - 1
                difference = difference Or (AscW(left(index)) Xor AscW(right(index)))
            Next
            Return difference = 0
        End Function
        Friend Shared Function NormalizeSha256(value As System.String) As System.String
            Dim normalized As System.String = value.Replace(" ", System.String.Empty).Replace("-", System.String.Empty).Trim().ToLowerInvariant()
            If normalized.Length <> 64 Then Throw New RedInkPythonAgentConfigurationException("SHA-256 must contain 64 hexadecimal characters.")
            For Each character As System.Char In normalized
                If Not System.Uri.IsHexDigit(character) Then Throw New RedInkPythonAgentConfigurationException("SHA-256 contains a non-hexadecimal character.")
            Next
            Return normalized
        End Function
        Friend Shared Function NormalizeSigner(value As System.String) As System.String
            Return System.Text.RegularExpressions.Regex.Replace(value.Trim(), " +", " ")
        End Function
    End Class


    Public NotInheritable Class RedInkAuthenticodeVerificationResult
        Public Property SignerOrganization As System.String = System.String.Empty
        Public Property CertificateThumbprint As System.String = System.String.Empty
        Public Property CertificateSubject As System.String = System.String.Empty
    End Class

    Public NotInheritable Class RedInkAuthenticodeVerifier
        Private Shared ReadOnly WinTrustActionGenericVerifyV2 As New System.Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE")

        Private Const WtdUiNone As System.UInt32 = 2UI
        Private Const WtdRevokeNone As System.UInt32 = 0UI
        Private Const WtdChoiceFile As System.UInt32 = 1UI
        Private Const WtdStateActionVerify As System.UInt32 = 1UI
        Private Const WtdStateActionClose As System.UInt32 = 2UI
        Private Const WtdCacheOnlyUrlRetrieval As System.UInt32 = &H1000UI
        Private Const CertNameAttrType As System.UInt32 = 3UI
        Private Const OrganizationOid As System.String = "2.5.4.10"

        <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential, CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
        Private Structure WinTrustFileInfo
            Public CbStruct As System.UInt32
            Public PcwszFilePath As System.IntPtr
            Public HFile As System.IntPtr
            Public PgKnownSubject As System.IntPtr
        End Structure

        <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential, CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
        Private Structure WinTrustData
            Public CbStruct As System.UInt32
            Public PPolicyCallbackData As System.IntPtr
            Public PSipClientData As System.IntPtr
            Public DwUiChoice As System.UInt32
            Public FdwRevocationChecks As System.UInt32
            Public DwUnionChoice As System.UInt32
            Public PFile As System.IntPtr
            Public DwStateAction As System.UInt32
            Public HWvtStateData As System.IntPtr
            Public PwszUrlReference As System.IntPtr
            Public DwProvFlags As System.UInt32
            Public DwUiContext As System.UInt32
        End Structure

        <System.Runtime.InteropServices.DllImport("wintrust.dll", ExactSpelling:=True, CharSet:=System.Runtime.InteropServices.CharSet.Unicode, SetLastError:=True)>
        Private Shared Function WinVerifyTrust(
        hWnd As System.IntPtr,
        ByRef actionId As System.Guid,
        trustData As System.IntPtr
    ) As System.Int32
        End Function

        <System.Runtime.InteropServices.DllImport("crypt32.dll", EntryPoint:="CertGetNameStringW", ExactSpelling:=True, CharSet:=System.Runtime.InteropServices.CharSet.Unicode, SetLastError:=True)>
        Private Shared Function CertGetNameStringW(
        certificateContext As System.IntPtr,
        nameType As System.UInt32,
        flags As System.UInt32,
        typeParameter As System.IntPtr,
        nameBuilder As System.Text.StringBuilder,
        nameBuilderLength As System.UInt32
    ) As System.UInt32
        End Function

        Private Sub New()
        End Sub

        Public Shared Function Verify(executableFullPath As System.String) As RedInkAuthenticodeVerificationResult
            If System.String.IsNullOrWhiteSpace(executableFullPath) Then
                Throw New System.ArgumentException("Executable path is empty.", NameOf(executableFullPath))
            End If

            VerifyWindowsTrust(executableFullPath)

            Try
                Using nativeCertificate As System.Security.Cryptography.X509Certificates.X509Certificate = System.Security.Cryptography.X509Certificates.X509Certificate.CreateFromSignedFile(executableFullPath)
                    Using certificate As New System.Security.Cryptography.X509Certificates.X509Certificate2(nativeCertificate)
                        Dim organization As System.String = ReadOrganization(certificate)
                        If System.String.IsNullOrWhiteSpace(organization) Then
                            Throw New RedInkPythonAgentExecutableTrustException("The Authenticode signer certificate has no organization attribute.")
                        End If
                        Return New RedInkAuthenticodeVerificationResult() With {
                        .SignerOrganization = organization.Trim(),
                        .CertificateThumbprint = If(certificate.Thumbprint, System.String.Empty).Replace(" ", System.String.Empty).ToLowerInvariant(),
                        .CertificateSubject = If(certificate.Subject, System.String.Empty)
                    }
                    End Using
                End Using
            Catch ex As RedInkPythonAgentExecutableTrustException
                Throw
            Catch ex As System.Exception
                Throw New RedInkPythonAgentExecutableTrustException("The Authenticode signer certificate could not be read.", ex)
            End Try
        End Function

        Private Shared Sub VerifyWindowsTrust(executableFullPath As System.String)
            Dim filePathPointer As System.IntPtr = System.IntPtr.Zero
            Dim fileInfoPointer As System.IntPtr = System.IntPtr.Zero
            Dim trustDataPointer As System.IntPtr = System.IntPtr.Zero
            Dim stateWasOpened As System.Boolean = False

            Try
                filePathPointer = System.Runtime.InteropServices.Marshal.StringToCoTaskMemUni(executableFullPath)

                Dim fileInfo As New WinTrustFileInfo() With {
                .CbStruct = CUInt(System.Runtime.InteropServices.Marshal.SizeOf(GetType(WinTrustFileInfo))),
                .PcwszFilePath = filePathPointer,
                .HFile = System.IntPtr.Zero,
                .PgKnownSubject = System.IntPtr.Zero
            }
                fileInfoPointer = System.Runtime.InteropServices.Marshal.AllocCoTaskMem(System.Runtime.InteropServices.Marshal.SizeOf(GetType(WinTrustFileInfo)))
                System.Runtime.InteropServices.Marshal.StructureToPtr(fileInfo, fileInfoPointer, False)

                Dim trustData As New WinTrustData() With {
                .CbStruct = CUInt(System.Runtime.InteropServices.Marshal.SizeOf(GetType(WinTrustData))),
                .PPolicyCallbackData = System.IntPtr.Zero,
                .PSipClientData = System.IntPtr.Zero,
                .DwUiChoice = WtdUiNone,
                .FdwRevocationChecks = WtdRevokeNone,
                .DwUnionChoice = WtdChoiceFile,
                .PFile = fileInfoPointer,
                .DwStateAction = WtdStateActionVerify,
                .HWvtStateData = System.IntPtr.Zero,
                .PwszUrlReference = System.IntPtr.Zero,
                .DwProvFlags = WtdCacheOnlyUrlRetrieval,
                .DwUiContext = 0UI
            }
                trustDataPointer = System.Runtime.InteropServices.Marshal.AllocCoTaskMem(System.Runtime.InteropServices.Marshal.SizeOf(GetType(WinTrustData)))
                System.Runtime.InteropServices.Marshal.StructureToPtr(trustData, trustDataPointer, False)

                Dim actionId As System.Guid = WinTrustActionGenericVerifyV2
                Dim result As System.Int32 = WinVerifyTrust(System.IntPtr.Zero, actionId, trustDataPointer)
                stateWasOpened = True
                If result <> 0 Then
                    Throw New RedInkPythonAgentExecutableTrustException(
                    "Windows rejected the executable Authenticode signature. WinVerifyTrust=0x" &
                    result.ToString("X8", System.Globalization.CultureInfo.InvariantCulture) & ".")
                End If
            Finally
                If stateWasOpened AndAlso trustDataPointer <> System.IntPtr.Zero Then
                    Try
                        Dim closeData As WinTrustData = CType(System.Runtime.InteropServices.Marshal.PtrToStructure(trustDataPointer, GetType(WinTrustData)), WinTrustData)
                        closeData.DwStateAction = WtdStateActionClose
                        System.Runtime.InteropServices.Marshal.StructureToPtr(closeData, trustDataPointer, True)
                        Dim closeAction As System.Guid = WinTrustActionGenericVerifyV2
                        WinVerifyTrust(System.IntPtr.Zero, closeAction, trustDataPointer)
                    Catch ex As System.Exception
                        ' Closing the WinTrust state must not hide an earlier verification result.
                    End Try
                End If

                If trustDataPointer <> System.IntPtr.Zero Then
                    System.Runtime.InteropServices.Marshal.FreeCoTaskMem(trustDataPointer)
                End If
                If fileInfoPointer <> System.IntPtr.Zero Then
                    System.Runtime.InteropServices.Marshal.FreeCoTaskMem(fileInfoPointer)
                End If
                If filePathPointer <> System.IntPtr.Zero Then
                    System.Runtime.InteropServices.Marshal.FreeCoTaskMem(filePathPointer)
                End If
            End Try
        End Sub

        Private Shared Function ReadOrganization(certificate As System.Security.Cryptography.X509Certificates.X509Certificate2) As System.String
            Dim oidPointer As System.IntPtr = System.IntPtr.Zero
            Try
                oidPointer = System.Runtime.InteropServices.Marshal.StringToHGlobalAnsi(OrganizationOid)
                Dim required As System.UInt32 = CertGetNameStringW(
                certificate.Handle,
                CertNameAttrType,
                0UI,
                oidPointer,
                Nothing,
                0UI)
                If required <= 1UI Then
                    Return System.String.Empty
                End If

                Dim builder As New System.Text.StringBuilder(CInt(required))
                Dim written As System.UInt32 = CertGetNameStringW(
                certificate.Handle,
                CertNameAttrType,
                0UI,
                oidPointer,
                builder,
                required)
                If written <= 1UI Then
                    Return System.String.Empty
                End If
                Return builder.ToString()
            Finally
                If oidPointer <> System.IntPtr.Zero Then
                    System.Runtime.InteropServices.Marshal.FreeHGlobal(oidPointer)
                End If
            End Try
        End Function
    End Class

    Public NotInheritable Class RedInkPythonAgentConfigurationParser
        Private Sub New()
        End Sub

        Public Shared Function Parse(line As System.String) As RedInkPythonAgentConfiguration
            If System.String.IsNullOrWhiteSpace(line) Then
                Throw New RedInkPythonAgentConfigurationException("PythonAgentPath configuration is empty.")
            End If

            Dim equalsIndex As System.Int32 = line.IndexOf("="c)
            If equalsIndex <= 0 Then
                Throw New RedInkPythonAgentConfigurationException("PythonAgentPath configuration must contain '='.")
            End If
            Dim key As System.String = line.Substring(0, equalsIndex).Trim()
            If Not System.String.Equals(key, "PythonAgentPath", System.StringComparison.OrdinalIgnoreCase) Then
                Throw New RedInkPythonAgentConfigurationException("Unknown configuration key.")
            End If

            Dim fields As System.Collections.Generic.List(Of System.String) = SplitFields(line.Substring(equalsIndex + 1))
            If fields.Count < 2 OrElse fields.Count > 3 Then
                Throw New RedInkPythonAgentConfigurationException("PythonAgentPath requires an executable path and at least one trust criterion.")
            End If

            Dim result As New RedInkPythonAgentConfiguration() With {
            .ExecutablePath = Unquote(fields(0)).Trim()
        }

            For index As System.Int32 = 1 To fields.Count - 1
                Dim field As System.String = fields(index).Trim()
                Dim separator As System.Int32 = field.IndexOf("="c)
                If separator > 0 Then
                    Dim fieldName As System.String = field.Substring(0, separator).Trim()
                    Dim fieldValue As System.String = Unquote(field.Substring(separator + 1)).Trim()
                    If System.String.Equals(fieldName, "signer", System.StringComparison.OrdinalIgnoreCase) Then
                        If result.ExpectedSignerOrganization IsNot Nothing Then
                            Throw New RedInkPythonAgentConfigurationException("Duplicate signer criterion.")
                        End If
                        result.ExpectedSignerOrganization = NormalizeSigner(fieldValue)
                    ElseIf System.String.Equals(fieldName, "sha256", System.StringComparison.OrdinalIgnoreCase) Then
                        If result.ExpectedSha256 IsNot Nothing Then
                            Throw New RedInkPythonAgentConfigurationException("Duplicate SHA-256 criterion.")
                        End If
                        result.ExpectedSha256 = NormalizeSha256(fieldValue)
                    Else
                        Throw New RedInkPythonAgentConfigurationException("Unknown trust criterion: " & fieldName)
                    End If
                Else
                    Dim bare As System.String = Unquote(field).Trim()
                    If LooksLikeSha256(bare) Then
                        If result.ExpectedSha256 IsNot Nothing Then
                            Throw New RedInkPythonAgentConfigurationException("Duplicate SHA-256 criterion.")
                        End If
                        result.ExpectedSha256 = NormalizeSha256(bare)
                    Else
                        If result.ExpectedSignerOrganization IsNot Nothing Then
                            Throw New RedInkPythonAgentConfigurationException("Duplicate signer criterion.")
                        End If
                        result.ExpectedSignerOrganization = NormalizeSigner(bare)
                    End If
                End If
            Next

            If System.String.IsNullOrWhiteSpace(result.ExecutablePath) Then
                Throw New RedInkPythonAgentConfigurationException("Executable path is empty.")
            End If
            If result.ExpectedSignerOrganization Is Nothing AndAlso result.ExpectedSha256 Is Nothing Then
                Throw New RedInkPythonAgentConfigurationException("At least one trust criterion is required.")
            End If
            Return result
        End Function

        Private Shared Function SplitFields(value As System.String) As System.Collections.Generic.List(Of System.String)
            Dim result As New System.Collections.Generic.List(Of System.String)()
            Dim builder As New System.Text.StringBuilder()
            Dim quoted As System.Boolean = False
            Dim index As System.Int32 = 0
            While index < value.Length
                Dim character As System.Char = value(index)
                If character = """"c Then
                    If quoted AndAlso index + 1 < value.Length AndAlso value(index + 1) = """"c Then
                        builder.Append(""""c)
                        index += 2
                        Continue While
                    End If
                    quoted = Not quoted
                    builder.Append(character)
                ElseIf character = ";"c AndAlso Not quoted Then
                    result.Add(builder.ToString().Trim())
                    builder.Clear()
                Else
                    builder.Append(character)
                End If
                index += 1
            End While
            If quoted Then
                Throw New RedInkPythonAgentConfigurationException("Unterminated quoted configuration value.")
            End If
            result.Add(builder.ToString().Trim())
            If result.Exists(Function(item As System.String) item.Length = 0) Then
                Throw New RedInkPythonAgentConfigurationException("Configuration contains an empty field.")
            End If
            Return result
        End Function

        Private Shared Function Unquote(value As System.String) As System.String
            Dim trimmed As System.String = value.Trim()
            If trimmed.Length >= 2 AndAlso trimmed(0) = """"c AndAlso trimmed(trimmed.Length - 1) = """"c Then
                Return trimmed.Substring(1, trimmed.Length - 2).Replace("""""", """")
            End If
            If trimmed.Contains(""""c) Then
                Throw New RedInkPythonAgentConfigurationException("A quote may only occur around a complete value.")
            End If
            Return trimmed
        End Function

        Private Shared Function LooksLikeSha256(value As System.String) As System.Boolean
            Try
                NormalizeSha256(value)
                Return True
            Catch ex As RedInkPythonAgentConfigurationException
                Return False
            End Try
        End Function

        Private Shared Function NormalizeSha256(value As System.String) As System.String
            Dim normalized As System.String = value.Replace(" ", System.String.Empty).Replace("-", System.String.Empty).Trim().ToLowerInvariant()
            If normalized.Length <> 64 Then
                Throw New RedInkPythonAgentConfigurationException("SHA-256 must contain exactly 64 hexadecimal characters.")
            End If
            For Each character As System.Char In normalized
                If Not ((character >= "0"c AndAlso character <= "9"c) OrElse (character >= "a"c AndAlso character <= "f"c)) Then
                    Throw New RedInkPythonAgentConfigurationException("SHA-256 contains a non-hexadecimal character.")
                End If
            Next
            Return normalized
        End Function

        Private Shared Function NormalizeSigner(value As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(value) Then
                Throw New RedInkPythonAgentConfigurationException("Signer organization is empty.")
            End If
            Dim parts As System.String() = value.Trim().Split(New System.Char() {" "c}, System.StringSplitOptions.RemoveEmptyEntries)
            Dim normalized As System.String = System.String.Join(" ", parts)
            If normalized.Length > 256 Then
                Throw New RedInkPythonAgentConfigurationException("Signer organization is too long.")
            End If
            Return normalized
        End Function
    End Class

End Namespace
