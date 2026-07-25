Option Explicit On
Option Strict On
Option Infer On

Namespace Agents
    Public NotInheritable Class PythonExecuteToolArgumentException
        Inherits System.Exception

        Public Sub New(message As System.String)
            MyBase.New(message)
        End Sub
    End Class

    Public NotInheritable Class PythonExecuteToolCoreOptions
        Public Property AgentConfiguration As RedInkPythonAgentConfiguration
        Public Property RootDirectory As System.String = System.String.Empty
        Public Property InputFileResolver As System.Func(Of System.String, RedInkPythonAgentInputFile)
        Public Property HostServiceHandler As IRedInkPythonAgentHostServiceHandler
        Public Property DefaultTimeoutSeconds As System.Int32 = 30
        Public Property MinimumTimeoutSeconds As System.Int32 = 5
        Public Property MaximumTimeoutSeconds As System.Int32 = 180
        Public Property MaximumCodeBytes As System.Int32 = 1048576
        Public Property MaximumStdinBytes As System.Int32 = 1048576
        Public Property MaximumInputFiles As System.Int32 = 100
        Public Property MaximumOutputBytes As System.Int64 = 67108864L
        Public Property MaximumOutputFiles As System.Int32 = 10
        Public Property MaximumWorkingBytes As System.Int64 = 2147483648L
        Public Property MaximumWorkingFiles As System.Int32 = 10000
        Public Property MaximumResultBytes As System.Int64 = 67108864L
        Public Property MaximumResultJsonDepth As System.Int32 = 128
        Public Property MaximumResultJsonNodes As System.Int32 = 1000000
        Public Property AllowedOperations As System.Collections.Generic.IList(Of System.String) = New System.Collections.Generic.List(Of System.String)(New System.String() {"llm.complete", "web.get", "web.search"})
        Public Property MaximumHostCalls As System.Int32 = 20
        Public Property MaximumConcurrentHostCalls As System.Int32 = 2
        Public Property MaximumHostRequestBytes As System.Int32 = 67108864
        Public Property MaximumHostResponseBytes As System.Int32 = 134217728
        Public Property DefaultHostCallTimeoutSeconds As System.Int32 = 60
        Public Property MaximumHostCallTimeoutSeconds As System.Int32 = 180
        Public Property MemoryMiB As System.Int32 = 1536
        Public Property HeartbeatTimeout As System.TimeSpan = System.TimeSpan.FromSeconds(15)
        Public Property PollInterval As System.TimeSpan = System.TimeSpan.FromMilliseconds(100)
        Public Property CancellationGracePeriod As System.TimeSpan = System.TimeSpan.FromSeconds(3)
        Public Property HardKillWait As System.TimeSpan = System.TimeSpan.FromSeconds(5)
        Public Property LogAgentToolCall As System.Action(Of System.Object, System.Object)
        Public Property LogRawResponse As System.Action(Of System.String, System.String)
        Public Property PublishOutputFile As System.Action(Of RedInkPythonAgentOutput)

        ''' <summary>
        ''' Creates a shallow copy so a single execution can apply per-call capability overrides
        ''' (allowed operations and host-service handler) without mutating the shared configured
        ''' options, which would otherwise race across concurrent tooling loops / sub-agents.
        ''' </summary>
        Public Function Clone() As PythonExecuteToolCoreOptions
            Dim copy As New PythonExecuteToolCoreOptions() With {
                .AgentConfiguration = Me.AgentConfiguration,
                .RootDirectory = Me.RootDirectory,
                .InputFileResolver = Me.InputFileResolver,
                .HostServiceHandler = Me.HostServiceHandler,
                .DefaultTimeoutSeconds = Me.DefaultTimeoutSeconds,
                .MinimumTimeoutSeconds = Me.MinimumTimeoutSeconds,
                .MaximumTimeoutSeconds = Me.MaximumTimeoutSeconds,
                .MaximumCodeBytes = Me.MaximumCodeBytes,
                .MaximumStdinBytes = Me.MaximumStdinBytes,
                .MaximumInputFiles = Me.MaximumInputFiles,
                .MaximumOutputBytes = Me.MaximumOutputBytes,
                .MaximumOutputFiles = Me.MaximumOutputFiles,
                .MaximumWorkingBytes = Me.MaximumWorkingBytes,
                .MaximumWorkingFiles = Me.MaximumWorkingFiles,
                .MaximumResultBytes = Me.MaximumResultBytes,
                .MaximumResultJsonDepth = Me.MaximumResultJsonDepth,
                .MaximumResultJsonNodes = Me.MaximumResultJsonNodes,
                .MaximumHostCalls = Me.MaximumHostCalls,
                .MaximumConcurrentHostCalls = Me.MaximumConcurrentHostCalls,
                .MaximumHostRequestBytes = Me.MaximumHostRequestBytes,
                .MaximumHostResponseBytes = Me.MaximumHostResponseBytes,
                .DefaultHostCallTimeoutSeconds = Me.DefaultHostCallTimeoutSeconds,
                .MaximumHostCallTimeoutSeconds = Me.MaximumHostCallTimeoutSeconds,
                .MemoryMiB = Me.MemoryMiB,
                .HeartbeatTimeout = Me.HeartbeatTimeout,
                .PollInterval = Me.PollInterval,
                .CancellationGracePeriod = Me.CancellationGracePeriod,
                .HardKillWait = Me.HardKillWait,
                .LogAgentToolCall = Me.LogAgentToolCall,
                .LogRawResponse = Me.LogRawResponse,
                .PublishOutputFile = Me.PublishOutputFile
            }
            copy.AllowedOperations = New System.Collections.Generic.List(Of System.String)(Me.AllowedOperations)
            Return copy
        End Function
    End Class

    Public NotInheritable Class PythonExecuteToolCoreResult
        Public Property Payload As System.String = System.String.Empty
        Public Property Success As System.Boolean
        Public Property Status As System.String = System.String.Empty
        Public Property ErrorCode As System.String = System.String.Empty
        Public Property ErrorMessage As System.String = System.String.Empty
        Public Property ExitCode As System.Int32
        Public Property DurationMilliseconds As System.Int64
        Public Property DiagnosticId As System.Guid
        Public Property HumanLogAvailable As System.Boolean
        Public Property RunResult As RedInkPythonAgentRunResult
    End Class

    Public NotInheritable Class PythonExecuteToolCore
        Public Const ToolName As System.String = "python_execute"
        Private Const ReservedStdinRelativePath As System.String = "__redink_tool/stdin.txt"

        Public Shared ReadOnly Property ToolInstructionsPrompt As System.String
            Get
                Return "Use python_execute for calculations, parsing, structured transformations, deterministic data processing, document editing, and generating output artifacts from self-contained Python code. When a task involves several steps or several files, combine them all into a single script and call python_execute once; do not issue multiple python_execute calls for one logical task. Access every input file through redink_pythonagent.agent_api.input_path(name), passing the same relative name you listed in input_files (for example: from redink_pythonagent import agent_api; doc = docx.Document(agent_api.input_path('Schreiben.docx'))). Never open an input by a bare filename or absolute path; staged inputs are not in the working directory and a bare open will fail with a not-found error. Write every produced document to a path obtained from agent_api.output_path(name); do not write to arbitrary or absolute paths. Publish a direct JSON result with agent_api.publish_result(...) or text with agent_api.publish_result_text(...); use output_path(...) for large or binary documents. The worker has no direct network access. Depending on how the host is configured for this task, it may additionally have access to host-mediated capabilities such as language-model assistance and web retrieval/search; when such a capability is not enabled, any attempt to use it fails with a typed host error, so treat these as optional and check availability at runtime rather than assuming them. Only explicitly supplied input files are visible. Execution and all relays are time- and size-bounded. Raw stdout, stderr, tracebacks, and arbitrary exception text remain human diagnostics; structured safe errors are returned to the model."
            End Get
        End Property

        Public Shared ReadOnly Property ToolDefinitionJson As System.String
            Get
                Dim definition As New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("name", ToolName),
                New Newtonsoft.Json.Linq.JProperty("description", "Executes a self-contained Python script in a sandboxed, network-isolated process. It may return an explicit bounded JSON/text result, create validated output documents, and, depending on host configuration for the task, may additionally have access to optional host-mediated capabilities such as language-model assistance and web retrieval/search. Use for calculations, parsing, transformations, document editing, and file generation. The worker has no direct network access and sees only files explicitly passed in."),
                New Newtonsoft.Json.Linq.JProperty("parameters", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("type", "object"),
                    New Newtonsoft.Json.Linq.JProperty("properties", New Newtonsoft.Json.Linq.JObject(
                        New Newtonsoft.Json.Linq.JProperty("code", New Newtonsoft.Json.Linq.JObject(
                            New Newtonsoft.Json.Linq.JProperty("type", "string"),
                            New Newtonsoft.Json.Linq.JProperty("description", "The complete Python source to execute. Must be self-contained."))),
                        New Newtonsoft.Json.Linq.JProperty("stdin", New Newtonsoft.Json.Linq.JObject(
                            New Newtonsoft.Json.Linq.JProperty("type", "string"),
                            New Newtonsoft.Json.Linq.JProperty("description", "Optional text exposed to the script through sys.stdin."))),
                        New Newtonsoft.Json.Linq.JProperty("input_files", New Newtonsoft.Json.Linq.JObject(
                            New Newtonsoft.Json.Linq.JProperty("type", "array"),
                            New Newtonsoft.Json.Linq.JProperty("items", New Newtonsoft.Json.Linq.JObject(
                                New Newtonsoft.Json.Linq.JProperty("type", "string"))),
                            New Newtonsoft.Json.Linq.JProperty("description", "Optional relative names made available read-only inside the sandbox. Inside the script, open each one via redink_pythonagent.agent_api.input_path(name) using the same name listed here; do not open it by a bare filename or absolute path, because staged inputs are not in the working directory."))),
                        New Newtonsoft.Json.Linq.JProperty("timeout_seconds", New Newtonsoft.Json.Linq.JObject(
                            New Newtonsoft.Json.Linq.JProperty("type", "integer"),
                            New Newtonsoft.Json.Linq.JProperty("description", "Optional absolute wall-clock limit, clamped to host policy; default 30."))))),
                    New Newtonsoft.Json.Linq.JProperty("required", New Newtonsoft.Json.Linq.JArray("code")))))
                Return definition.ToString(Newtonsoft.Json.Formatting.None)
            End Get
        End Property

        Private Sub New()
        End Sub

        Public Shared Function ResolveAndValidateAvailability(options As PythonExecuteToolCoreOptions) As System.String
            ValidateOptions(options)
            Dim configuration As RedInkPythonAgentConfiguration = ResolveConfigurationRelativeToAssembly(options.AgentConfiguration)
            Return RedInkPythonAgentClient.ValidateExecutableConfiguration(configuration)
        End Function

        Public Shared Async Function ExecuteAsync(
        options As PythonExecuteToolCoreOptions,
        arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
        cancellationToken As System.Threading.CancellationToken,
        Optional logStep As System.Action(Of System.String) = Nothing,
        Optional logInfo As System.Action(Of System.String) = Nothing,
        Optional logWarn As System.Action(Of System.String) = Nothing,
        Optional logDiag As System.Action(Of System.String) = Nothing
    ) As System.Threading.Tasks.Task(Of PythonExecuteToolCoreResult)
            Dim result As PythonExecuteToolCoreResult = Nothing
            Dim argumentFailure As PythonExecuteToolArgumentException = Nothing
            Dim configurationFailure As RedInkPythonAgentConfigurationException = Nothing
            Dim unexpectedFailure As System.Exception = Nothing
            Try
                result = Await ExecuteValidatedAsync(options, arguments, cancellationToken, logStep, logInfo, logWarn, logDiag).ConfigureAwait(False)
            Catch ex As PythonExecuteToolArgumentException
                argumentFailure = ex
            Catch ex As RedInkPythonAgentConfigurationException
                configurationFailure = ex
            Catch ex As System.OperationCanceledException
                Throw
            Catch ex As System.Exception
                unexpectedFailure = ex
            End Try
            If argumentFailure IsNot Nothing Then
                SafeLog(logDiag, argumentFailure.ToString())
                Return CreateLocalFailure("REQUEST_INVALID", "failed", "The Python execution request is invalid.")
            End If
            If configurationFailure IsNot Nothing Then
                SafeLog(logDiag, configurationFailure.ToString())
                Return CreateLocalFailure("CONFIGURATION_INVALID", "failed", FriendlyMessage("CONFIGURATION_INVALID", "failed"))
            End If
            If unexpectedFailure IsNot Nothing Then
                SafeLog(logDiag, unexpectedFailure.ToString())
                Return CreateLocalFailure("INTERNAL_BROKER_ERROR", "failed", FriendlyMessage("INTERNAL_BROKER_ERROR", "failed"))
            End If
            Return result
        End Function

        Private Shared Async Function ExecuteValidatedAsync(
        options As PythonExecuteToolCoreOptions,
        arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object),
        cancellationToken As System.Threading.CancellationToken,
        logStep As System.Action(Of System.String),
        logInfo As System.Action(Of System.String),
        logWarn As System.Action(Of System.String),
        logDiag As System.Action(Of System.String)
    ) As System.Threading.Tasks.Task(Of PythonExecuteToolCoreResult)

            ValidateOptions(options)
            If arguments Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(arguments))
            End If
            cancellationToken.ThrowIfCancellationRequested()

            Dim code As System.String = GetRequiredString(arguments, "code")
            Dim codeBytes As System.Int32 = System.Text.Encoding.UTF8.GetByteCount(code)
            If codeBytes > options.MaximumCodeBytes Then
                Return CreateLocalFailure("REQUEST_INVALID", "failed", "Python source exceeds the configured limit.")
            End If

            Dim stdinText As System.String = GetOptionalString(arguments, "stdin")
            If System.Text.Encoding.UTF8.GetByteCount(stdinText) > options.MaximumStdinBytes Then
                Return CreateLocalFailure("REQUEST_INVALID", "failed", "Standard input exceeds the configured limit.")
            End If

            Dim requestedTimeout As System.Int32 = GetTimeoutSeconds(arguments, options.DefaultTimeoutSeconds)
            Dim effectiveTimeout As System.Int32 = Clamp(requestedTimeout, options.MinimumTimeoutSeconds, options.MaximumTimeoutSeconds)
            If effectiveTimeout <> requestedTimeout Then
                SafeLog(logWarn, "Python timeout was clamped to host policy.")
            End If

            Dim requestedInputFiles As System.Collections.Generic.List(Of System.String) = GetInputFiles(arguments, options.MaximumInputFiles)
            Dim resolvedInputFiles As New System.Collections.Generic.List(Of RedInkPythonAgentInputFile)()
            Dim seenInputFiles As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each requested As System.String In requestedInputFiles
                cancellationToken.ThrowIfCancellationRequested()
                Dim relative As System.String = ValidateWorkspaceRelativePath(requested)
                If System.String.Equals(relative, ReservedStdinRelativePath, System.StringComparison.OrdinalIgnoreCase) Then
                    Return CreateLocalFailure("REQUEST_INVALID", "failed", "The requested input path is reserved.")
                End If
                If Not seenInputFiles.Add(relative) Then
                    Return CreateLocalFailure("REQUEST_INVALID", "failed", "Duplicate input file path.")
                End If
                If options.InputFileResolver Is Nothing Then
                    Return CreateLocalFailure("CONFIGURATION_INVALID", "failed", "Input-file resolution is not configured.")
                End If
                Dim resolved As RedInkPythonAgentInputFile = Nothing
                Try
                    resolved = options.InputFileResolver(relative)
                Catch ex As System.Exception
                    SafeLog(logDiag, ex.ToString())
                    Return CreateLocalFailure("REQUEST_INVALID", "failed", "Workspace input file could not be resolved.")
                End Try
                If resolved Is Nothing OrElse System.String.IsNullOrWhiteSpace(resolved.SourcePath) OrElse Not System.IO.File.Exists(System.IO.Path.GetFullPath(resolved.SourcePath)) Then
                    Return CreateLocalFailure("REQUEST_INVALID", "failed", "Workspace input file was not found.")
                End If
                resolvedInputFiles.Add(New RedInkPythonAgentInputFile(System.IO.Path.GetFullPath(resolved.SourcePath), relative))
            Next

            Dim temporaryInputDirectory As System.String = Nothing
            If stdinText.Length <> 0 Then
                temporaryInputDirectory = CreateTemporaryInputDirectory()
                Dim stdinSource As System.String = System.IO.Path.Combine(temporaryInputDirectory, "stdin.txt")
                System.IO.File.WriteAllText(stdinSource, stdinText, New System.Text.UTF8Encoding(False, True))
                resolvedInputFiles.Add(New RedInkPythonAgentInputFile(stdinSource, ReservedStdinRelativePath))
                code = WrapCodeForStandardInput(code)
            End If

            Dim callRoot As System.String = CreateCallRoot(options.RootDirectory)
            Dim retainDiagnostics As System.Boolean = False
            Try
                Dim configuration As RedInkPythonAgentConfiguration = ResolveConfigurationRelativeToAssembly(options.AgentConfiguration)
                Dim limits As RedInkPythonAgentLimits = CreateLimits(options, effectiveTimeout)
                Dim executionOptions As New RedInkPythonAgentExecutionOptions() With {
                .OverallTimeout = System.TimeSpan.FromSeconds(effectiveTimeout) + options.CancellationGracePeriod + options.HardKillWait + System.TimeSpan.FromSeconds(2),
                .HeartbeatTimeout = options.HeartbeatTimeout,
                .PollInterval = options.PollInterval,
                .CancellationGracePeriod = options.CancellationGracePeriod,
                .HardKillWait = options.HardKillWait,
                .HumanDiagnosticLog = logDiag
            }
                Dim client As New RedInkPythonAgentClient(executionOptions)
                Dim progress As System.IProgress(Of RedInkPythonAgentEvent) = New PythonExecuteProgressAdapter(logStep, logInfo, logWarn, logDiag)
                Dim execution As RedInkPythonAgentExecution = Nothing
                Dim runResult As RedInkPythonAgentRunResult = Nothing
                Dim started As System.DateTimeOffset = System.DateTimeOffset.UtcNow
                Dim cancelled As System.Boolean = False
                Dim unexpected As System.Exception = Nothing

                SafeLog(logStep, "Running secure Python script...")
                Try
                    Dim executable As System.String = RedInkPythonAgentClient.ValidateExecutableConfiguration(configuration)
                    SafeLog(logDiag, "Verified PythonAgent executable: " & executable)
                    SafeLog(logDiag, "Python request source bytes: " & codeBytes.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    execution = client.CreateExecution(configuration, callRoot, code, resolvedInputFiles, limits, options.HostServiceHandler, progress)
                    Using registration As System.Threading.CancellationTokenRegistration = cancellationToken.Register(
                    Sub()
                        If execution IsNot Nothing Then
                            execution.CancelAsync().GetAwaiter().GetResult()
                        End If
                    End Sub)
                        runResult = Await execution.Completion.ConfigureAwait(False)
                    End Using
                Catch ex As System.OperationCanceledException
                    cancelled = True
                Catch ex As System.Exception
                    unexpected = ex
                Finally
                    DeleteDirectoryBestEffort(temporaryInputDirectory, logDiag)
                End Try

                If cancelled OrElse cancellationToken.IsCancellationRequested Then
                    SafeLog(logWarn, "Python execution was cancelled.")
                    Return CreateLocalFailure("SESSION_CANCELLED", "cancelled", "Operation was cancelled.")
                End If
                If unexpected IsNot Nothing Then
                    retainDiagnostics = True
                    SafeLog(logDiag, unexpected.ToString())
                    Dim trustException As RedInkPythonAgentExecutableTrustException = TryCast(unexpected, RedInkPythonAgentExecutableTrustException)
                    If trustException IsNot Nothing Then
                        Return CreateLocalFailure(trustException.Code, "failed", FriendlyMessage(trustException.Code, "failed"))
                    End If
                    If TypeOf unexpected Is RedInkPythonAgentConfigurationException Then
                        Return CreateLocalFailure("CONFIGURATION_INVALID", "failed", FriendlyMessage("CONFIGURATION_INVALID", "failed"))
                    End If
                    Return CreateLocalFailure("INTERNAL_BROKER_ERROR", "failed", FriendlyMessage("INTERNAL_BROKER_ERROR", "failed"))
                End If
                If runResult Is Nothing Then
                    retainDiagnostics = True
                    Return CreateLocalFailure("BROKER_EXITED_WITHOUT_RESPONSE", "failed", FriendlyMessage("BROKER_EXITED_WITHOUT_RESPONSE", "failed"))
                End If

                Dim duration As System.Int64 = System.Convert.ToInt64((System.DateTimeOffset.UtcNow - started).TotalMilliseconds)
                Dim result As PythonExecuteToolCoreResult = CreateResultFromRun(runResult, duration)
                If options.PublishOutputFile IsNot Nothing Then
                    Dim expectedResultsRoot As System.String = System.IO.Path.GetFullPath(System.IO.Path.Combine(callRoot, "results", runResult.SessionId.ToString("D")))
                    For Each output As RedInkPythonAgentOutput In runResult.Outputs
                        Try
                            If Not ValidatePublishedOutput(output, expectedResultsRoot, logDiag) Then
                                Continue For
                            End If
                            options.PublishOutputFile.Invoke(output)
                        Catch ex As System.Exception
                            SafeLog(logDiag, ex.ToString())
                        End Try
                    Next
                End If
                If Not result.Success Then
                    retainDiagnostics = True
                End If
                If result.Success Then
                    SafeLog(logStep, "Python script finished (exit " & result.ExitCode.ToString(System.Globalization.CultureInfo.InvariantCulture) & ", " & duration.ToString(System.Globalization.CultureInfo.InvariantCulture) & " ms).")
                Else
                    SafeLog(logWarn, "Python script finished with status " & result.Status & ".")
                End If
                SafeLog(logDiag, "PythonAgent model-safe response bytes: " & System.Text.Encoding.UTF8.GetByteCount(result.Payload).ToString(System.Globalization.CultureInfo.InvariantCulture))
                Return result
            Finally
                If retainDiagnostics Then
                    RetainDiagnosticsBestEffort(callRoot, options.RootDirectory, logDiag)
                End If
                DeleteDirectoryBestEffort(callRoot, logDiag)
            End Try
        End Function

        Private Shared Function ValidatePublishedOutput(output As RedInkPythonAgentOutput, expectedResultsRoot As System.String, logDiag As System.Action(Of System.String)) As System.Boolean
            If output Is Nothing Then Return False
            Dim relative As System.String = If(output.RelativePath, System.String.Empty)
            If relative.Length = 0 Then
                SafeLog(logDiag, "Rejected python_execute output with an empty relative path.")
                Return False
            End If
            If System.IO.Path.IsPathRooted(relative) Then
                SafeLog(logDiag, "Rejected rooted python_execute output path: " & relative)
                Return False
            End If
            For Each part As System.String In relative.Replace("\"c, "/"c).Split("/"c)
                If part = "." OrElse part = ".." Then
                    SafeLog(logDiag, "Rejected python_execute output path with relative navigation: " & relative)
                    Return False
                End If
            Next
            Dim fullPath As System.String
            Try
                fullPath = System.IO.Path.GetFullPath(output.FullPath)
            Catch ex As System.Exception
                SafeLog(logDiag, ex.ToString())
                Return False
            End Try
            Dim containmentPrefix As System.String = expectedResultsRoot.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar) & System.IO.Path.DirectorySeparatorChar
            If Not fullPath.StartsWith(containmentPrefix, System.StringComparison.OrdinalIgnoreCase) Then
                SafeLog(logDiag, "Rejected python_execute output outside the results directory: " & fullPath)
                Return False
            End If
            Dim information As System.IO.FileInfo
            Try
                information = New System.IO.FileInfo(fullPath)
            Catch ex As System.Exception
                SafeLog(logDiag, ex.ToString())
                Return False
            End Try
            If Not information.Exists Then
                SafeLog(logDiag, "python_execute output file is missing: " & fullPath)
                Return False
            End If
            If (information.Attributes And System.IO.FileAttributes.Directory) = System.IO.FileAttributes.Directory OrElse
               (information.Attributes And System.IO.FileAttributes.ReparsePoint) = System.IO.FileAttributes.ReparsePoint Then
                SafeLog(logDiag, "python_execute output is not a regular file: " & fullPath)
                Return False
            End If
            If information.Length <> output.Size Then
                SafeLog(logDiag, "python_execute output size mismatch: " & fullPath)
                Return False
            End If
            If Not VerifyFileSha256(fullPath, output.Sha256, logDiag) Then
                SafeLog(logDiag, "python_execute output SHA-256 mismatch: " & fullPath)
                Return False
            End If
            output.PublishedSubPath = fullPath.Substring(containmentPrefix.Length).Replace(System.IO.Path.DirectorySeparatorChar, "/"c)
            Return True
        End Function

        Private Shared Function VerifyFileSha256(fullPath As System.String, expectedHex As System.String, logDiag As System.Action(Of System.String)) As System.Boolean
            If System.String.IsNullOrWhiteSpace(expectedHex) Then Return False
            Try
                Using stream As New System.IO.FileStream(fullPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                    Using sha As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                        Dim hashBytes As System.Byte() = sha.ComputeHash(stream)
                        Dim builder As New System.Text.StringBuilder(hashBytes.Length * 2)
                        For Each value As System.Byte In hashBytes
                            builder.Append(value.ToString("x2", System.Globalization.CultureInfo.InvariantCulture))
                        Next
                        Return System.String.Equals(builder.ToString(), expectedHex.Trim(), System.StringComparison.OrdinalIgnoreCase)
                    End Using
                End Using
            Catch ex As System.Exception
                SafeLog(logDiag, ex.ToString())
                Return False
            End Try
        End Function

        Private Shared Sub RetainDiagnosticsBestEffort(callRoot As System.String, baseRoot As System.String, logDiag As System.Action(Of System.String))
            Try
                If System.String.IsNullOrWhiteSpace(callRoot) OrElse Not System.IO.Directory.Exists(callRoot) Then Return
                Dim diagnosticsRoot As System.String = System.IO.Path.Combine(System.IO.Path.GetFullPath(System.Environment.ExpandEnvironmentVariables(baseRoot)), "diagnostics")
                System.IO.Directory.CreateDirectory(diagnosticsRoot)
                PruneDiagnostics(diagnosticsRoot, 20, logDiag)
                Dim target As System.String = System.IO.Path.Combine(diagnosticsRoot, System.DateTime.UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'", System.Globalization.CultureInfo.InvariantCulture) & "_" & System.Guid.NewGuid().ToString("N"))
                System.IO.Directory.CreateDirectory(target)
                Dim responsePath As System.String = System.IO.Path.Combine(callRoot, "response.json")
                If System.IO.File.Exists(responsePath) Then
                    CopyFileSharedRead(responsePath, System.IO.Path.Combine(target, "response.json"), logDiag)
                End If
                For Each logPath As System.String In System.IO.Directory.EnumerateFiles(callRoot, "*.log", System.IO.SearchOption.AllDirectories)
                    CopyFileSharedRead(logPath, System.IO.Path.Combine(target, System.IO.Path.GetFileName(logPath)), logDiag)
                Next
                SafeLog(logDiag, "python_execute diagnostics retained under: " & target)
            Catch ex As System.Exception
                SafeLog(logDiag, ex.ToString())
            End Try
        End Sub

        Private Shared Sub CopyFileSharedRead(sourcePath As System.String, destinationPath As System.String, logDiag As System.Action(Of System.String))
            Try
                Using source As New System.IO.FileStream(sourcePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite Or System.IO.FileShare.Delete)
                    Using destination As New System.IO.FileStream(destinationPath, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None)
                        source.CopyTo(destination)
                    End Using
                End Using
            Catch ex As System.Exception
                SafeLog(logDiag, ex.ToString())
            End Try
        End Sub

        Private Shared Sub PruneDiagnostics(diagnosticsRoot As System.String, maximumEntries As System.Int32, logDiag As System.Action(Of System.String))
            Try
                Dim directories As New System.Collections.Generic.List(Of System.String)(System.IO.Directory.EnumerateDirectories(diagnosticsRoot))
                If directories.Count < maximumEntries Then Return
                directories.Sort(Function(a, b) System.IO.Directory.GetCreationTimeUtc(a).CompareTo(System.IO.Directory.GetCreationTimeUtc(b)))
                Dim removeCount As System.Int32 = directories.Count - maximumEntries + 1
                For index As System.Int32 = 0 To removeCount - 1
                    DeleteDirectoryBestEffort(directories(index), logDiag)
                Next
            Catch ex As System.Exception
                SafeLog(logDiag, ex.ToString())
            End Try
        End Sub

        Private Shared Function CreateLimits(options As PythonExecuteToolCoreOptions, timeoutSeconds As System.Int32) As RedInkPythonAgentLimits
            Dim validatorReserve As System.Int32 = System.Math.Max(1, System.Math.Min(30, timeoutSeconds \ 4))
            Dim executeSeconds As System.Int32 = System.Math.Max(1, timeoutSeconds - validatorReserve)
            Dim operations As New System.Collections.Generic.List(Of System.String)()
            If options.HostServiceHandler IsNot Nothing Then
                For Each operation As System.String In options.AllowedOperations
                    operations.Add(operation)
                Next
            End If
            Return New RedInkPythonAgentLimits() With {
            .OverallWallTimeSeconds = timeoutSeconds,
            .ExecuteWallTimeSeconds = executeSeconds,
            .ValidatorWallTimeSeconds = validatorReserve,
            .MemoryMiB = options.MemoryMiB,
            .MaxOutputBytes = options.MaximumOutputBytes,
            .MaxOutputFiles = options.MaximumOutputFiles,
            .MaxWorkingBytes = options.MaximumWorkingBytes,
            .MaxWorkingFiles = options.MaximumWorkingFiles,
            .MaxResultBytes = options.MaximumResultBytes,
            .MaxResultJsonDepth = options.MaximumResultJsonDepth,
            .MaxResultJsonNodes = options.MaximumResultJsonNodes,
            .AllowedOperations = operations,
            .MaximumCalls = If(operations.Count = 0, 0, options.MaximumHostCalls),
            .MaximumConcurrentCalls = If(operations.Count = 0, 0, options.MaximumConcurrentHostCalls),
            .MaximumRequestBytes = options.MaximumHostRequestBytes,
            .MaximumResponseBytes = options.MaximumHostResponseBytes,
            .DefaultCallTimeoutSeconds = options.DefaultHostCallTimeoutSeconds,
            .MaximumCallTimeoutSeconds = options.MaximumHostCallTimeoutSeconds
        }
        End Function

        Private Shared Function CreateResultFromRun(runResult As RedInkPythonAgentRunResult, durationMilliseconds As System.Int64) As PythonExecuteToolCoreResult
            Dim success As System.Boolean = System.String.Equals(runResult.Status, "success", System.StringComparison.Ordinal)
            Dim code As System.String = If(runResult.Error Is Nothing, System.String.Empty, runResult.Error.Code)
            Dim payload As New Newtonsoft.Json.Linq.JObject(
            New Newtonsoft.Json.Linq.JProperty("status", runResult.Status),
            New Newtonsoft.Json.Linq.JProperty("exit_code", runResult.ExitCode),
            New Newtonsoft.Json.Linq.JProperty("duration_ms", durationMilliseconds),
            New Newtonsoft.Json.Linq.JProperty("diagnostic_id", runResult.DiagnosticId.ToString("D")),
            New Newtonsoft.Json.Linq.JProperty("human_log_available", runResult.HumanLogAvailable))
            If runResult.Result IsNot Nothing Then
                payload.Add("result", New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("kind", runResult.Result.Kind),
                New Newtonsoft.Json.Linq.JProperty("value", runResult.Result.Value.DeepClone())))
            Else
                payload.Add("result", Newtonsoft.Json.Linq.JValue.CreateNull())
            End If
            Dim outputs As New Newtonsoft.Json.Linq.JArray()
            For Each output As RedInkPythonAgentOutput In runResult.Outputs
                outputs.Add(New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("name", output.RelativePath),
                New Newtonsoft.Json.Linq.JProperty("media_type", output.MediaType),
                New Newtonsoft.Json.Linq.JProperty("bytes", output.Size),
                New Newtonsoft.Json.Linq.JProperty("sha256", output.Sha256)))
            Next
            payload.Add("output_files", outputs)
            If runResult.Error IsNot Nothing Then
                payload.Add("error", CreateSafeErrorJson(runResult.Error))
            Else
                payload.Add("error", Newtonsoft.Json.Linq.JValue.CreateNull())
            End If
            Return New PythonExecuteToolCoreResult() With {
            .Payload = payload.ToString(Newtonsoft.Json.Formatting.None),
            .Success = success,
            .Status = runResult.Status,
            .ErrorCode = code,
            .ErrorMessage = If(success, System.String.Empty, FriendlyMessage(code, runResult.Status)),
            .ExitCode = runResult.ExitCode,
            .DurationMilliseconds = durationMilliseconds,
            .DiagnosticId = runResult.DiagnosticId,
            .HumanLogAvailable = runResult.HumanLogAvailable,
            .RunResult = runResult
        }
        End Function

        Private Shared Function CreateSafeErrorJson(errorValue As RedInkPythonAgentSafeError) As Newtonsoft.Json.Linq.JObject
            Dim sourceToken As Newtonsoft.Json.Linq.JToken = Newtonsoft.Json.Linq.JValue.CreateNull()
            If errorValue.Source IsNot Nothing Then
                sourceToken = New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("file", errorValue.Source.File),
                New Newtonsoft.Json.Linq.JProperty("line", errorValue.Source.Line),
                New Newtonsoft.Json.Linq.JProperty("column", If(errorValue.Source.Column.HasValue, New Newtonsoft.Json.Linq.JValue(errorValue.Source.Column.Value), Newtonsoft.Json.Linq.JValue.CreateNull())),
                New Newtonsoft.Json.Linq.JProperty("function", If(errorValue.Source.Function Is Nothing, Newtonsoft.Json.Linq.JValue.CreateNull(), New Newtonsoft.Json.Linq.JValue(errorValue.Source.Function))),
                New Newtonsoft.Json.Linq.JProperty("symbol", If(errorValue.Source.Symbol Is Nothing, Newtonsoft.Json.Linq.JValue.CreateNull(), New Newtonsoft.Json.Linq.JValue(errorValue.Source.Symbol))))
            End If
            Dim stack As New Newtonsoft.Json.Linq.JArray()
            For Each frame As RedInkPythonAgentSafeStackFrame In errorValue.Stack
                stack.Add(New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("file", frame.File),
                New Newtonsoft.Json.Linq.JProperty("line", frame.Line),
                New Newtonsoft.Json.Linq.JProperty("function", If(frame.Function Is Nothing, Newtonsoft.Json.Linq.JValue.CreateNull(), New Newtonsoft.Json.Linq.JValue(frame.Function)))))
            Next
            Return New Newtonsoft.Json.Linq.JObject(
            New Newtonsoft.Json.Linq.JProperty("code", errorValue.Code),
            New Newtonsoft.Json.Linq.JProperty("phase", errorValue.Phase),
            New Newtonsoft.Json.Linq.JProperty("retryable", errorValue.Retryable),
            New Newtonsoft.Json.Linq.JProperty("source", sourceToken),
            New Newtonsoft.Json.Linq.JProperty("host_operation", If(errorValue.HostOperation Is Nothing, Newtonsoft.Json.Linq.JValue.CreateNull(), New Newtonsoft.Json.Linq.JValue(errorValue.HostOperation))),
            New Newtonsoft.Json.Linq.JProperty("limit", If(errorValue.Limit.HasValue, New Newtonsoft.Json.Linq.JValue(errorValue.Limit.Value), Newtonsoft.Json.Linq.JValue.CreateNull())),
            New Newtonsoft.Json.Linq.JProperty("observed", If(errorValue.Observed.HasValue, New Newtonsoft.Json.Linq.JValue(errorValue.Observed.Value), Newtonsoft.Json.Linq.JValue.CreateNull())),
            New Newtonsoft.Json.Linq.JProperty("stack", stack))
        End Function

        Private Shared Function CreateLocalFailure(code As System.String, status As System.String, message As System.String) As PythonExecuteToolCoreResult
            Dim diagnosticId As System.Guid = System.Guid.NewGuid()
            Dim payload As New Newtonsoft.Json.Linq.JObject(
            New Newtonsoft.Json.Linq.JProperty("status", status),
            New Newtonsoft.Json.Linq.JProperty("exit_code", 1),
            New Newtonsoft.Json.Linq.JProperty("duration_ms", 0),
            New Newtonsoft.Json.Linq.JProperty("diagnostic_id", diagnosticId.ToString("D")),
            New Newtonsoft.Json.Linq.JProperty("human_log_available", False),
            New Newtonsoft.Json.Linq.JProperty("result", Newtonsoft.Json.Linq.JValue.CreateNull()),
            New Newtonsoft.Json.Linq.JProperty("output_files", New Newtonsoft.Json.Linq.JArray()),
            New Newtonsoft.Json.Linq.JProperty("error", New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("code", code),
                New Newtonsoft.Json.Linq.JProperty("phase", "initializing"),
                New Newtonsoft.Json.Linq.JProperty("retryable", False),
                New Newtonsoft.Json.Linq.JProperty("source", Newtonsoft.Json.Linq.JValue.CreateNull()),
                New Newtonsoft.Json.Linq.JProperty("host_operation", Newtonsoft.Json.Linq.JValue.CreateNull()),
                New Newtonsoft.Json.Linq.JProperty("limit", Newtonsoft.Json.Linq.JValue.CreateNull()),
                New Newtonsoft.Json.Linq.JProperty("observed", Newtonsoft.Json.Linq.JValue.CreateNull()),
                New Newtonsoft.Json.Linq.JProperty("stack", New Newtonsoft.Json.Linq.JArray()))))
            Return New PythonExecuteToolCoreResult() With {
            .Payload = payload.ToString(Newtonsoft.Json.Formatting.None),
            .Success = False,
            .Status = status,
            .ErrorCode = code,
            .ErrorMessage = message,
            .ExitCode = 1,
            .DurationMilliseconds = 0,
            .DiagnosticId = diagnosticId,
            .HumanLogAvailable = False
        }
        End Function

        Public Shared Function FriendlyMessage(code As System.String, status As System.String) As System.String
            Select Case code
                Case "SESSION_CANCELLED"
                    Return "Operation was cancelled."
                Case "SESSION_TIMEOUT", "WORKER_TIMEOUT", "HOST_CALL_TIMEOUT"
                    Return "Python execution timed out."
                Case "EXECUTABLE_NOT_FOUND"
                    Return "The secure Python executor is unavailable."
                Case "EXECUTABLE_HASH_MISMATCH", "EXECUTABLE_SIGNATURE_INVALID", "EXECUTABLE_SIGNER_MISMATCH"
                    Return "The secure Python executor failed authenticity verification."
                Case "CONFIGURATION_INVALID", "REQUEST_INVALID"
                    Return "The Python execution request is invalid."
                Case "PYTHON_RESULT_TOO_LARGE"
                    Return "The direct Python result exceeded the configured byte limit; write large content to an output file instead."
                Case "PYTHON_RESULT_TOO_DEEP", "PYTHON_RESULT_TOO_COMPLEX", "PYTHON_RESULT_INVALID"
                    Return "The direct Python result was not valid within the configured JSON limits."
                Case "OUTPUT_VALIDATION_FAILED", "OUTPUT_PUBLICATION_FAILED"
                    Return "Python output validation failed."
                Case "BROKER_HEARTBEAT_LOST", "BROKER_EXITED_WITHOUT_RESPONSE", "BROKER_START_FAILED"
                    Return "The secure Python executor stopped unexpectedly."
                Case Else
                    If System.String.Equals(status, "cancelled", System.StringComparison.Ordinal) Then
                        Return "Operation was cancelled."
                    End If
                    If System.String.Equals(status, "timeout", System.StringComparison.Ordinal) Then
                        Return "Python execution timed out."
                    End If
                    Return "Python execution failed."
            End Select
        End Function

        Private Shared Function ResolveConfigurationRelativeToAssembly(configuration As RedInkPythonAgentConfiguration) As RedInkPythonAgentConfiguration
            If configuration Is Nothing Then
                Throw New RedInkPythonAgentConfigurationException("PythonAgent configuration is required.")
            End If
            Dim executablePath As System.String = configuration.ExecutablePath
            If System.String.IsNullOrWhiteSpace(executablePath) Then
                Throw New RedInkPythonAgentConfigurationException("PythonAgent executable path is required.")
            End If
            executablePath = System.Environment.ExpandEnvironmentVariables(executablePath)
            If Not System.IO.Path.IsPathRooted(executablePath) Then
                Dim assemblyPath As System.String = System.Reflection.Assembly.GetExecutingAssembly().Location
                Dim assemblyDirectory As System.String = System.IO.Path.GetDirectoryName(assemblyPath)
                executablePath = System.IO.Path.Combine(assemblyDirectory, executablePath)
            End If
            Return New RedInkPythonAgentConfiguration() With {
            .executablePath = System.IO.Path.GetFullPath(executablePath),
            .ExpectedSignerOrganization = configuration.ExpectedSignerOrganization,
            .ExpectedSha256 = configuration.ExpectedSha256
        }
        End Function

        Private Shared Sub ValidateOptions(options As PythonExecuteToolCoreOptions)
            If options Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(options))
            End If
            If options.AgentConfiguration Is Nothing Then
                Throw New RedInkPythonAgentConfigurationException("PythonAgent configuration is required.")
            End If
            If System.String.IsNullOrWhiteSpace(options.RootDirectory) Then
                options.RootDirectory = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "RedInk", "PythonAgent")
            End If
            If options.MinimumTimeoutSeconds < 1 OrElse options.MaximumTimeoutSeconds < options.MinimumTimeoutSeconds OrElse options.DefaultTimeoutSeconds < options.MinimumTimeoutSeconds OrElse options.DefaultTimeoutSeconds > options.MaximumTimeoutSeconds Then
                Throw New RedInkPythonAgentConfigurationException("Python execute timeout policy is invalid.")
            End If
            If options.MaximumCodeBytes < 1 OrElse options.MaximumStdinBytes < 0 OrElse options.MaximumInputFiles < 0 OrElse options.MaximumOutputBytes < 1 OrElse options.MaximumOutputFiles < 1 Then
                Throw New RedInkPythonAgentConfigurationException("Python execute size policy is invalid.")
            End If
            If options.MaximumResultBytes < 1L OrElse options.MaximumResultBytes > 134217728L OrElse options.MaximumResultJsonDepth < 1 OrElse options.MaximumResultJsonDepth > 256 OrElse options.MaximumResultJsonNodes < 1 OrElse options.MaximumResultJsonNodes > 4000000 Then
                Throw New RedInkPythonAgentConfigurationException("Python execute result policy is invalid.")
            End If
            If options.AllowedOperations Is Nothing Then Throw New RedInkPythonAgentConfigurationException("Python host-service policy is invalid.")
            Dim seenOperations As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each operation As System.String In options.AllowedOperations
                If operation <> "llm.complete" AndAlso operation <> "web.get" AndAlso operation <> "web.search" Then Throw New RedInkPythonAgentConfigurationException("Python host-service operation is invalid.")
                If Not seenOperations.Add(operation) Then Throw New RedInkPythonAgentConfigurationException("Python host-service operations contain duplicates.")
            Next
            If options.MaximumHostCalls < 0 OrElse options.MaximumHostCalls > 100 OrElse options.MaximumConcurrentHostCalls < 0 OrElse options.MaximumConcurrentHostCalls > 4 OrElse options.MaximumConcurrentHostCalls > options.MaximumHostCalls Then
                Throw New RedInkPythonAgentConfigurationException("Python host-service call limits are invalid.")
            End If
            If options.MaximumHostRequestBytes < 1 OrElse options.MaximumHostRequestBytes > 67108864 OrElse options.MaximumHostResponseBytes < 1 OrElse options.MaximumHostResponseBytes > 134217728 Then
                Throw New RedInkPythonAgentConfigurationException("Python host-service size limits are invalid.")
            End If
            If options.DefaultHostCallTimeoutSeconds < 1 OrElse options.MaximumHostCallTimeoutSeconds < options.DefaultHostCallTimeoutSeconds OrElse options.MaximumHostCallTimeoutSeconds > 600 Then
                Throw New RedInkPythonAgentConfigurationException("Python host-service timeout policy is invalid.")
            End If
        End Sub

        Private Shared Function GetRequiredString(arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object), name As System.String) As System.String
            Dim raw As System.Object = Nothing
            If Not arguments.TryGetValue(name, raw) Then
                Throw New PythonExecuteToolArgumentException("Missing required argument: " & name)
            End If
            Dim value As System.String = TryCast(raw, System.String)
            If System.String.IsNullOrWhiteSpace(value) Then
                Throw New PythonExecuteToolArgumentException("Missing required argument: " & name)
            End If
            Return value
        End Function

        Private Shared Function GetOptionalString(arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object), name As System.String) As System.String
            Dim raw As System.Object = Nothing
            If Not arguments.TryGetValue(name, raw) OrElse raw Is Nothing Then
                Return System.String.Empty
            End If
            Dim value As System.String = TryCast(raw, System.String)
            If value Is Nothing Then
                Throw New PythonExecuteToolArgumentException("Invalid " & name & " value.")
            End If
            Return value
        End Function

        Private Shared Function GetTimeoutSeconds(arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object), defaultValue As System.Int32) As System.Int32
            Dim raw As System.Object = Nothing
            If Not arguments.TryGetValue("timeout_seconds", raw) OrElse raw Is Nothing Then
                Return defaultValue
            End If
            If TypeOf raw Is System.Boolean Then
                Throw New PythonExecuteToolArgumentException("Invalid timeout_seconds value.")
            End If
            Dim token As Newtonsoft.Json.Linq.JValue = TryCast(raw, Newtonsoft.Json.Linq.JValue)
            If token IsNot Nothing Then
                raw = token.Value
            End If
            Try
                Dim converted As System.Int64 = System.Convert.ToInt64(raw, System.Globalization.CultureInfo.InvariantCulture)
                If converted < System.Int32.MinValue OrElse converted > System.Int32.MaxValue Then
                    Throw New System.OverflowException()
                End If
                Return System.Convert.ToInt32(converted)
            Catch ex As System.Exception
                Throw New PythonExecuteToolArgumentException("Invalid timeout_seconds value.")
            End Try
        End Function

        Private Shared Function GetInputFiles(arguments As System.Collections.Generic.IDictionary(Of System.String, System.Object), maximumCount As System.Int32) As System.Collections.Generic.List(Of System.String)
            Dim result As New System.Collections.Generic.List(Of System.String)()
            Dim raw As System.Object = Nothing
            If Not arguments.TryGetValue("input_files", raw) OrElse raw Is Nothing Then
                Return result
            End If
            If TypeOf raw Is System.String Then
                Throw New PythonExecuteToolArgumentException("Invalid input_files value.")
            End If
            Dim enumerable As System.Collections.IEnumerable = TryCast(raw, System.Collections.IEnumerable)
            If enumerable Is Nothing Then
                Throw New PythonExecuteToolArgumentException("Invalid input_files value.")
            End If
            For Each item As System.Object In enumerable
                If result.Count >= maximumCount Then
                    Throw New PythonExecuteToolArgumentException("Too many input files.")
                End If
                Dim token As Newtonsoft.Json.Linq.JValue = TryCast(item, Newtonsoft.Json.Linq.JValue)
                Dim value As System.String = If(token Is Nothing, TryCast(item, System.String), TryCast(token.Value, System.String))
                If System.String.IsNullOrWhiteSpace(value) Then
                    Throw New PythonExecuteToolArgumentException("Invalid input_files value.")
                End If
                result.Add(value)
            Next
            Return result
        End Function

        Private Shared Function ValidateWorkspaceRelativePath(value As System.String) As System.String
            Dim normalized As System.String = value.Replace("\"c, "/"c).Trim()
            If normalized.Length = 0 OrElse System.IO.Path.IsPathRooted(normalized) OrElse normalized.IndexOf(":"c) >= 0 Then
                Throw New PythonExecuteToolArgumentException("Invalid workspace path.")
            End If
            Dim parts As System.String() = normalized.Split("/"c)
            For Each part As System.String In parts
                If part.Length = 0 OrElse part = "." OrElse part = ".." OrElse part.EndsWith(".", System.StringComparison.Ordinal) OrElse part.EndsWith(" ", System.StringComparison.Ordinal) Then
                    Throw New PythonExecuteToolArgumentException("Invalid workspace path.")
                End If
                If part.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) >= 0 Then
                    Throw New PythonExecuteToolArgumentException("Invalid workspace path.")
                End If
            Next
            Return System.String.Join("/", parts)
        End Function

        Private Shared Function WrapCodeForStandardInput(code As System.String) As System.String
            Dim encoded As System.String = System.Convert.ToBase64String(New System.Text.UTF8Encoding(False, True).GetBytes(code))
            Dim builder As New System.Text.StringBuilder()
            builder.AppendLine("import base64 as __redink_base64")
            builder.AppendLine("import io as __redink_io")
            builder.AppendLine("import sys as __redink_sys")
            builder.AppendLine("from redink_pythonagent import agent_api as __redink_agent_api")
            builder.AppendLine("__redink_sys.stdin = __redink_io.StringIO(__redink_agent_api.input_path('__redink_tool/stdin.txt').read_text(encoding='utf-8'))")
            builder.Append("__redink_source = __redink_base64.b64decode('")
            builder.Append(encoded)
            builder.AppendLine("').decode('utf-8')")
            builder.AppendLine("__redink_globals = {'__name__': '__main__', '__file__': 'code.py', '__package__': None, '__cached__': None}")
            builder.AppendLine("exec(compile(__redink_source, 'code.py', 'exec'), __redink_globals, __redink_globals)")
            Return builder.ToString()
        End Function

        Private Shared Function CreateCallRoot(baseRoot As System.String) As System.String
            Dim fullBase As System.String = System.IO.Path.GetFullPath(System.Environment.ExpandEnvironmentVariables(baseRoot))
            System.IO.Directory.CreateDirectory(fullBase)
            Dim callParent As System.String = System.IO.Path.Combine(fullBase, "python_execute")
            System.IO.Directory.CreateDirectory(callParent)
            ' Best-effort sweep of stale per-call roots left behind when a prior broker process
            ' crashed or was force-killed with locked files (the normal Finally cleanup could not
            ' delete them). These transient folders are never needed after their run completes.
            PruneStaleCallRoots(callParent)
            Return System.IO.Path.Combine(callParent, System.Guid.NewGuid().ToString("D"))
        End Function

        Private Shared Sub PruneStaleCallRoots(callParent As System.String)
            Try
                Dim cutoff As System.DateTime = System.DateTime.UtcNow.AddHours(-1)
                For Each directory As System.String In System.IO.Directory.EnumerateDirectories(callParent)
                    Try
                        If System.IO.Directory.GetLastWriteTimeUtc(directory) < cutoff Then
                            System.IO.Directory.Delete(directory, True)
                        End If
                    Catch
                        ' Still locked or in use; skip and retry on a later run.
                    End Try
                Next
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try
        End Sub

        Private Shared Function CreateTemporaryInputDirectory() As System.String
            Dim directory As System.String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "RedInkPythonExecute", System.Guid.NewGuid().ToString("D"))
            System.IO.Directory.CreateDirectory(directory)
            Return directory
        End Function

        Private Shared Sub DeleteDirectoryBestEffort(path As System.String, logDiag As System.Action(Of System.String))
            If System.String.IsNullOrWhiteSpace(path) Then
                Return
            End If
            Try
                If System.IO.Directory.Exists(path) Then
                    System.IO.Directory.Delete(path, True)
                End If
            Catch ex As System.Exception
                SafeLog(logDiag, ex.ToString())
            End Try
        End Sub

        Private Shared Function Clamp(value As System.Int32, minimum As System.Int32, maximum As System.Int32) As System.Int32
            If value < minimum Then
                Return minimum
            End If
            If value > maximum Then
                Return maximum
            End If
            Return value
        End Function

        Private Shared Sub SafeLog(logger As System.Action(Of System.String), message As System.String)
            If logger Is Nothing Then
                Return
            End If
            Try
                logger(message)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try
        End Sub

        Private NotInheritable Class PythonExecuteProgressAdapter
            Implements System.IProgress(Of RedInkPythonAgentEvent)

            Private ReadOnly LogStepValue As System.Action(Of System.String)
            Private ReadOnly LogInfoValue As System.Action(Of System.String)
            Private ReadOnly LogWarnValue As System.Action(Of System.String)
            Private ReadOnly LogDiagValue As System.Action(Of System.String)

            Public Sub New(logStep As System.Action(Of System.String), logInfo As System.Action(Of System.String), logWarn As System.Action(Of System.String), logDiag As System.Action(Of System.String))
                Me.LogStepValue = logStep
                Me.LogInfoValue = logInfo
                Me.LogWarnValue = logWarn
                Me.LogDiagValue = logDiag
            End Sub

            Public Sub Report(value As RedInkPythonAgentEvent) Implements System.IProgress(Of RedInkPythonAgentEvent).Report
                If value Is Nothing Then
                    Return
                End If
                Select Case value.EventCode
                    Case "BROKER_STARTED"
                        SafeLog(Me.LogStepValue, "Preparing secure Python execution...")
                    Case "REQUEST_VALIDATED"
                        SafeLog(Me.LogDiagValue, "Python request validated.")
                    Case "INPUT_STAGING_STARTED"
                        SafeLog(Me.LogStepValue, "Preparing sandbox input files...")
                    Case "INPUT_STAGING_COMPLETED"
                        SafeLog(Me.LogDiagValue, "Sandbox input files prepared.")
                    Case "RUNTIME_VERIFICATION_STARTED"
                        SafeLog(Me.LogStepValue, "Verifying secure Python runtime...")
                    Case "RUNTIME_VERIFICATION_COMPLETED"
                        SafeLog(Me.LogDiagValue, "Secure Python runtime verified.")
                    Case "EXECUTE_SANDBOX_STARTING", "EXECUTE_SANDBOX_RUNNING"
                        SafeLog(Me.LogStepValue, "Running Python script...")
                    Case "PYTHON_PROGRESS"
                        SafeLog(Me.LogStepValue, FormatProgress(value))
                    Case "VALIDATION_STARTED"
                        SafeLog(Me.LogStepValue, "Validating Python output files...")
                    Case "VALIDATION_COMPLETED"
                        SafeLog(Me.LogDiagValue, "Python output files validated.")
                    Case "PUBLICATION_STARTED"
                        SafeLog(Me.LogStepValue, "Publishing Python output files...")
                    Case "PUBLICATION_COMPLETED"
                        SafeLog(Me.LogDiagValue, "Python output files published.")
                    Case "CANCELLATION_REQUESTED", "SESSION_CANCELLED"
                        SafeLog(Me.LogWarnValue, "Python execution cancellation requested.")
                    Case "SESSION_TIMED_OUT"
                        SafeLog(Me.LogWarnValue, "Python execution timed out.")
                    Case "SESSION_FAILED"
                        SafeLog(Me.LogWarnValue, "Python execution failed.")
                    Case "LOG_TRUNCATED"
                        SafeLog(Me.LogWarnValue, "Python diagnostic logging was truncated by policy.")
                    Case Else
                        SafeLog(Me.LogDiagValue, "PythonAgent event: " & value.EventCode)
                End Select
            End Sub

            Private Shared Function FormatProgress(value As RedInkPythonAgentEvent) As System.String
                Dim builder As New System.Text.StringBuilder("Python progress")
                If value.Current.HasValue AndAlso value.Total.HasValue Then
                    builder.Append(": ")
                    builder.Append(value.Current.Value.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    builder.Append("/")
                    builder.Append(value.Total.Value.ToString(System.Globalization.CultureInfo.InvariantCulture))
                End If
                If Not System.String.IsNullOrWhiteSpace(value.StepId) Then
                    builder.Append(" (")
                    builder.Append(value.StepId)
                    builder.Append(")")
                End If
                builder.Append(".")
                Return builder.ToString()
            End Function
        End Class
    End Class

End Namespace
