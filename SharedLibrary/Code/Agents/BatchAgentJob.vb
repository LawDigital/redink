' Part of Red Ink (SharedLibrary). Copyright (c) LawDigital Ltd., Switzerland.
' File queue/checkpoints only. Each item delegates to the EXISTING host tooling harness.
' This class contains no model calling, tool dispatch, tool recovery or finalization loop.
Option Strict On
Option Explicit On

Namespace Agents
    Public NotInheritable Class BatchJobOptions
        Public Property SourceDirectory As System.String = ""
        Public Property OutputBaseDirectory As System.String = ""
        Public Property Instruction As System.String = ""
        Public Property ModelKey As System.String = ""
        Public Property ModelDisplayName As System.String = ""
        Public Property ToolNames As New System.Collections.Generic.List(Of System.String)()
        Public Property Extensions As System.String = ".txt;.md;.docx;.doc;.rtf;.pdf;.xlsx;.xls;.xlsm;.csv;.tsv;.msg;.eml;.pptx;.ppt;.html;.json;.xml"
        Public Property Recursive As System.Boolean = False
        Public Property CarryForward As System.Boolean = False
        Public Property StopOnError As System.Boolean = False
        Public Property RequireArtifact As System.Boolean = False
        Public Property ShowDashboard As System.Boolean = False
        Public Property MaxRetries As System.Int32 = 2
        Public Property RetryDelaySeconds As System.Int32 = 3
        Public Property AttemptTimeoutMinutes As System.Int32 = 30
        Public Property MaxToolIterations As System.Int32 = 100
        Public Property MaxInputMegabytes As System.Int32 = 100
    End Class

    Public NotInheritable Class BatchJobOutput
        Public Property RelativePath As System.String = ""
        Public Property Sha256 As System.String = ""
    End Class

    Public NotInheritable Class BatchJobAttempt
        Public Property Number As System.Int32
        Public Property StartedUtc As System.DateTime = System.DateTime.UtcNow
        Public Property EndedUtc As System.DateTime?
        Public Property Status As System.String = "running"
        Public Property ErrorDetail As System.String = ""
        Public Property DirectoryRelativePath As System.String = ""
    End Class

    Public NotInheritable Class BatchJobItem
        Public Property SourcePath As System.String = ""
        Public Property SourceSha256 As System.String = ""
        Public Property Status As System.String = "pending"
        Public Property ErrorMessage As System.String = ""
        Public Property Attempts As New System.Collections.Generic.List(Of BatchJobAttempt)()
        Public Property Outputs As New System.Collections.Generic.List(Of BatchJobOutput)()
        Public Property Response As BatchJobOutput
    End Class

    Public NotInheritable Class BatchJobManifest
        Public Property Schema As System.Int32 = 1
        Public Property JobId As System.String = System.Guid.NewGuid().ToString("N")
        Public Property CreatedUtc As System.DateTime = System.DateTime.UtcNow
        Public Property UpdatedUtc As System.DateTime = System.DateTime.UtcNow
        Public Property Status As System.String = "ready"
        Public Property Options As New BatchJobOptions()
        Public Property Items As New System.Collections.Generic.List(Of BatchJobItem)()
        Public Property DiscoveryWarnings As New System.Collections.Generic.List(Of System.String)()
    End Class

    Public NotInheritable Class BatchJobProgress
        Public Property Total As System.Int32
        Public Property Completed As System.Int32
        Public Property ItemIndex As System.Int32
        Public Property Attempt As System.Int32
        Public Property FileName As System.String = ""
        Public Property Message As System.String = ""
    End Class

    Public NotInheritable Class BatchAttemptRequest
        Public Property Options As BatchJobOptions
        Public Property JobId As System.String
        Public Property ItemIndex As System.Int32
        Public Property Total As System.Int32
        Public Property Attempt As System.Int32
        Public Property StagingDirectory As System.String
        Public Property InputPath As System.String
        Public Property OriginalSourceName As System.String
        Public Property CarryPaths As New System.Collections.Generic.List(Of System.String)()
        Public Property PreviousResponsePath As System.String
        Public Property LogPath As System.String
    End Class

    Public NotInheritable Class BatchAttemptResult
        Public Property Response As System.String = ""
        Public Property OutputPaths As New System.Collections.Generic.List(Of System.String)()
    End Class

    Public NotInheritable Class BatchJobPermanentException
        Inherits System.IO.IOException
        Public Sub New(message As System.String)
            MyBase.New(message)
        End Sub
    End Class

    Public NotInheritable Class BatchJobQueue
        Public Const OutputContainerName As System.String = "RedInk_Batch_Jobs"
        Public Const ManifestFileName As System.String = "batch-job.json"
        Private Sub New()
        End Sub

        Public Shared Sub ValidateOptions(options As BatchJobOptions)
            If options Is Nothing Then Throw New System.ArgumentNullException(NameOf(options))
            If System.String.IsNullOrWhiteSpace(options.Instruction) Then Throw New System.ArgumentException("Enter a task instruction.")
            If System.String.IsNullOrWhiteSpace(options.ModelKey) Then Throw New System.ArgumentException("Select a tooling-capable model.")
            If options.ToolNames Is Nothing OrElse options.ToolNames.Count = 0 Then Throw New System.ArgumentException("Select at least one tool.")
            If options.MaxRetries < 0 OrElse options.MaxRetries > 10 Then Throw New System.ArgumentException("Retries must be between 0 and 10.")
            If options.RetryDelaySeconds < 0 OrElse options.RetryDelaySeconds > 300 Then Throw New System.ArgumentException("Invalid retry delay.")
            If options.AttemptTimeoutMinutes < 1 OrElse options.AttemptTimeoutMinutes > 1440 Then Throw New System.ArgumentException("Invalid attempt timeout.")
            If options.MaxToolIterations < 1 OrElse options.MaxToolIterations > 1000 Then Throw New System.ArgumentException("Invalid tooling iteration limit.")
            If options.MaxInputMegabytes < 1 OrElse options.MaxInputMegabytes > 2048 Then Throw New System.ArgumentException("Invalid input size limit.")
            options.SourceDirectory = System.IO.Path.GetFullPath(options.SourceDirectory)
            options.OutputBaseDirectory = System.IO.Path.GetFullPath(options.OutputBaseDirectory)
            If Not System.IO.Directory.Exists(options.SourceDirectory) Then Throw New System.IO.DirectoryNotFoundException(options.SourceDirectory)
            If Not System.IO.Directory.Exists(options.OutputBaseDirectory) Then Throw New System.IO.DirectoryNotFoundException(options.OutputBaseDirectory)
            RejectReparseAncestors(options.SourceDirectory)
            RejectReparseAncestors(options.OutputBaseDirectory)
        End Sub

        Public Shared Function Create(options As BatchJobOptions,
                                      progress As System.IProgress(Of BatchJobProgress),
                                      ct As System.Threading.CancellationToken) As System.String
            ' Never persist ModelConfig, credentials, API headers or the global workspace.
            ValidateOptions(options)
            Dim job As New BatchJobManifest With {.Options = options}
            Dim paths As New System.Collections.Generic.List(Of System.String)()
            Discover(options.SourceDirectory, options, paths, job.DiscoveryWarnings, ct)
            paths.Sort(System.StringComparer.OrdinalIgnoreCase)
            For Each path As System.String In paths
                ct.ThrowIfCancellationRequested()
                Dim item As New BatchJobItem With {.SourcePath = path}
                Try
                    If New System.IO.FileInfo(path).Length > CLng(options.MaxInputMegabytes) * 1024L * 1024L Then
                        Throw New BatchJobPermanentException("Input exceeds the configured size limit.")
                    End If
                    item.SourceSha256 = HashFile(path, ct)
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    ' A temporarily unavailable file remains an item; never silently skip it.
                    item.ErrorMessage = "Input snapshot: " & ex.Message
                End Try
                job.Items.Add(item)
                If progress IsNot Nothing Then progress.Report(New BatchJobProgress With {
                    .Total = paths.Count, .Message = "Capturing input snapshot: " & System.IO.Path.GetFileName(path)})
            Next
            If job.Items.Count = 0 Then Throw New System.InvalidOperationException("No matching source files were found. Check the extension filter.")
            Dim root As System.String = System.IO.Path.Combine(options.OutputBaseDirectory, OutputContainerName,
                System.DateTime.Now.ToString("yyyyMMdd_HHmmss", System.Globalization.CultureInfo.InvariantCulture) & "_" & job.JobId)
            System.IO.Directory.CreateDirectory(root)
            RejectReparseAncestors(root)
            Dim manifestPath As System.String = System.IO.Path.Combine(root, ManifestFileName)
            Save(manifestPath, job)
            Return manifestPath
        End Function

        Private Shared Sub Discover(directory As System.String, options As BatchJobOptions,
                                    paths As System.Collections.Generic.List(Of System.String),
                                    warnings As System.Collections.Generic.List(Of System.String), ct As System.Threading.CancellationToken)
            ct.ThrowIfCancellationRequested()
            Try
                RejectReparseAncestors(directory)
                For Each path As System.String In System.IO.Directory.GetFiles(directory)
                    ct.ThrowIfCancellationRequested()
                    Dim attributes As System.IO.FileAttributes = System.IO.File.GetAttributes(path)
                    If (attributes And (System.IO.FileAttributes.ReparsePoint Or System.IO.FileAttributes.Hidden Or System.IO.FileAttributes.System)) <> 0 Then
                        warnings.Add("Excluded hidden/system/reparse file: " & path)
                        Continue For
                    End If
                    If System.IO.Path.GetFileName(path).StartsWith("~$", System.StringComparison.Ordinal) Then Continue For
                    If MatchesExtension(path, options.Extensions) Then paths.Add(System.IO.Path.GetFullPath(path))
                Next
                If options.Recursive Then
                    For Each child As System.String In System.IO.Directory.GetDirectories(directory)
                        If System.String.Equals(System.IO.Path.GetFileName(child), OutputContainerName, System.StringComparison.OrdinalIgnoreCase) Then Continue For
                        If (System.IO.File.GetAttributes(child) And (System.IO.FileAttributes.ReparsePoint Or System.IO.FileAttributes.Hidden Or System.IO.FileAttributes.System)) <> 0 Then
                            warnings.Add("Excluded hidden/system/reparse directory: " & child)
                            Continue For
                        End If
                        Discover(child, options, paths, warnings, ct)
                    Next
                End If
            Catch ex As System.OperationCanceledException
                Throw
            Catch ex As System.Exception
                warnings.Add("Cannot completely enumerate '" & directory & "': " & ex.Message)
            End Try
        End Sub

        Private Shared Function MatchesExtension(path As System.String, filter As System.String) As System.Boolean
            Dim extension As System.String = System.IO.Path.GetExtension(path)
            For Each part As System.String In If(filter, "").Split(New System.Char() {";"c, ","c, " "c}, System.StringSplitOptions.RemoveEmptyEntries)
                Dim candidate As System.String = part.Trim()
                If candidate = "*" OrElse candidate = "*.*" Then Return True
                candidate = candidate.TrimStart("*"c)
                If Not candidate.StartsWith(".", System.StringComparison.Ordinal) Then candidate = "." & candidate
                If System.String.Equals(extension, candidate, System.StringComparison.OrdinalIgnoreCase) Then Return True
            Next
            Return False
        End Function

        Public Shared Function Load(manifestPath As System.String) As BatchJobManifest
            RejectReparseAncestors(manifestPath)
            Dim job As BatchJobManifest = Newtonsoft.Json.JsonConvert.DeserializeObject(Of BatchJobManifest)(
                System.IO.File.ReadAllText(manifestPath, System.Text.Encoding.UTF8))
            If job Is Nothing OrElse job.Schema <> 1 OrElse job.Options Is Nothing OrElse job.Items Is Nothing OrElse job.DiscoveryWarnings Is Nothing Then
                Throw New System.IO.InvalidDataException("Unsupported or incomplete batch manifest.")
            End If
            ValidateOptions(job.Options)
            For Each item As BatchJobItem In job.Items
                If item Is Nothing OrElse item.Attempts Is Nothing OrElse item.Outputs Is Nothing OrElse
                   Not IsWithin(item.SourcePath, job.Options.SourceDirectory) Then
                    Throw New System.IO.InvalidDataException("Invalid input entry in batch manifest.")
                End If
                Select Case item.Status
                    Case "pending", "running", "interrupted", "succeeded", "failed", "cancelled"
                    Case Else
                        Throw New System.IO.InvalidDataException("Invalid item status in batch manifest.")
                End Select
                If item.Status = "succeeded" AndAlso item.Response Is Nothing Then
                    Throw New System.IO.InvalidDataException("A successful item is missing its committed answer.")
                End If
            Next
            Return job
        End Function

        Public Shared Async Function RunAsync(manifestPath As System.String,
            runner As System.Func(Of BatchAttemptRequest, System.IProgress(Of System.String), System.Threading.CancellationToken, System.Threading.Tasks.Task(Of BatchAttemptResult)),
            progress As System.IProgress(Of BatchJobProgress),
            ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of BatchJobManifest)

            If runner Is Nothing Then Throw New System.ArgumentNullException(NameOf(runner))
            Dim root As System.String = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(manifestPath))
            RejectReparseAncestors(root)
            ' One owner per job, including different Outlook processes and double clicks.
            Using lease As New System.IO.FileStream(System.IO.Path.Combine(root, ".run.lock"), System.IO.FileMode.OpenOrCreate,
                                                    System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
                Dim job As BatchJobManifest = Load(manifestPath)
                Dim previous As BatchJobItem = Nothing
                Dim gapSeen As System.Boolean = False
                For Each item As BatchJobItem In job.Items
                    If item.Status = "running" Then
                        item.Status = "interrupted"
                        item.ErrorMessage = "The previous execution stopped before a result was committed."
                        If item.Attempts.Count > 0 Then
                            Dim interrupted As BatchJobAttempt = item.Attempts(item.Attempts.Count - 1)
                            interrupted.Status = "interrupted"
                            interrupted.EndedUtc = System.DateTime.UtcNow
                        End If
                    End If
                    If item.Status <> "succeeded" Then
                        gapSeen = True
                    ElseIf gapSeen AndAlso job.Options.CarryForward Then
                        Throw New System.IO.InvalidDataException("A dependent batch has successes after an unresolved gap. Start a new job; replaying only the gap would use an inconsistent accumulator.")
                    End If
                Next
                job.Status = "running"
                Save(manifestPath, job)
                Try
                    For index As System.Int32 = 0 To job.Items.Count - 1
                        ct.ThrowIfCancellationRequested()
                        Dim item As BatchJobItem = job.Items(index)
                        If item.Status = "succeeded" Then
                            VerifyStoredOutput(root, item.Response, ct)
                            For Each output As BatchJobOutput In item.Outputs
                                VerifyStoredOutput(root, output, ct)
                            Next
                            previous = item
                            Report(progress, job, index, 0, "Keeping the already committed result.")
                            Continue For
                        End If
                        For retry As System.Int32 = 0 To job.Options.MaxRetries
                            ct.ThrowIfCancellationRequested()
                            Dim attempt As New BatchJobAttempt With {
                                .Number = item.Attempts.Count + 1,
                                .DirectoryRelativePath = System.IO.Path.Combine("attempts", (index + 1).ToString("D5", System.Globalization.CultureInfo.InvariantCulture), System.Guid.NewGuid().ToString("N"))}
                            item.Attempts.Add(attempt)
                            item.Status = "running"
                            item.ErrorMessage = ""
                            Save(manifestPath, job)
                            Report(progress, job, index, attempt.Number, "Preparing isolated attempt.")
                            Dim succeeded As System.Boolean = False
                            Dim permanentFailure As System.Boolean = False
                            Using attemptCts As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(ct)
                                attemptCts.CancelAfter(System.TimeSpan.FromMinutes(job.Options.AttemptTimeoutMinutes))
                                Try
                                    Dim request As BatchAttemptRequest = PrepareAttempt(root, job, item, previous, index, attempt, attemptCts.Token)
                                    Save(manifestPath, job) ' Persist any newly captured input hash before tool execution.
                                    Dim messages As System.IProgress(Of System.String) = New InlineProgress(
                                        Sub(message As System.String) Report(progress, job, index, attempt.Number, message))
                                    Dim result As BatchAttemptResult = Await runner(request, messages, attemptCts.Token).ConfigureAwait(False)
                                    attemptCts.Token.ThrowIfCancellationRequested()
                                    If result Is Nothing OrElse System.String.IsNullOrWhiteSpace(result.Response) Then
                                        Throw New System.IO.InvalidDataException("The agent returned no final response.")
                                    End If
                                    If result.OutputPaths Is Nothing Then Throw New System.IO.InvalidDataException("Missing output collection.")
                                    If job.Options.RequireArtifact AndAlso result.OutputPaths.Count = 0 Then
                                        Throw New System.IO.InvalidDataException("This job requires at least one final result file.")
                                    End If
                                    If job.Options.CarryForward AndAlso request.CarryPaths.Count > 0 AndAlso result.OutputPaths.Count = 0 Then
                                        Throw New System.IO.InvalidDataException("Carry-forward files were supplied but no current final result file was returned. The accumulator must not silently disappear.")
                                    End If
                                    Commit(root, item, request, result, attemptCts.Token)
                                    item.Status = "succeeded"
                                    attempt.Status = "succeeded"
                                    succeeded = True
                                Catch ex As System.OperationCanceledException When ct.IsCancellationRequested
                                    item.Status = "cancelled"
                                    item.ErrorMessage = "Cancelled by the user; no attempt result was committed."
                                    attempt.Status = "cancelled"
                                    attempt.ErrorDetail = item.ErrorMessage
                                    Throw
                                Catch ex As System.Exception
                                    item.Status = "failed"
                                    item.ErrorMessage = If(TypeOf ex Is System.OperationCanceledException, "Attempt timeout expired.", ex.Message)
                                    attempt.Status = "failed"
                                    attempt.ErrorDetail = ex.ToString()
                                    permanentFailure = TypeOf ex Is BatchJobPermanentException
                                Finally
                                    attempt.EndedUtc = System.DateTime.UtcNow
                                End Try
                            End Using
                            ' Checkpoint/report failures abort the job, never retry a committed model action.
                            Save(manifestPath, job)
                            Report(progress, job, index, attempt.Number, If(succeeded, "Result committed.", item.ErrorMessage))
                            If succeeded OrElse permanentFailure OrElse retry = job.Options.MaxRetries Then Exit For
                            Await System.Threading.Tasks.Task.Delay(System.TimeSpan.FromSeconds(job.Options.RetryDelaySeconds), ct).ConfigureAwait(False)
                        Next
                        If item.Status = "succeeded" Then
                            previous = item
                        ElseIf job.Options.CarryForward OrElse job.Options.StopOnError Then
                            Exit For
                        End If
                    Next
                    Dim failed As System.Boolean = False
                    For Each item As BatchJobItem In job.Items
                        If item.Status <> "succeeded" Then failed = True
                    Next
                    job.Status = If(failed, "incomplete", If(job.DiscoveryWarnings.Count > 0, "completed_with_warnings", "completed"))
                Catch ex As System.OperationCanceledException
                    job.Status = "cancelled"
                    Throw
                Catch ex As System.Exception
                    job.Status = "interrupted"
                    Throw
                Finally
                    Save(manifestPath, job)
                    Report(progress, job, -1, 0, "Batch status: " & job.Status)
                End Try
                Return job
            End Using
        End Function

        Private Shared Function PrepareAttempt(root As System.String, job As BatchJobManifest, item As BatchJobItem,
                                               previous As BatchJobItem, index As System.Int32, attempt As BatchJobAttempt,
                                               ct As System.Threading.CancellationToken) As BatchAttemptRequest
            Dim attemptDir As System.String = ResolveStoredPath(root, attempt.DirectoryRelativePath)
            Dim stage As System.String = System.IO.Path.Combine(attemptDir, "files")
            System.IO.Directory.CreateDirectory(stage)
            RejectReparseAncestors(item.SourcePath)
            If New System.IO.FileInfo(item.SourcePath).Length > CLng(job.Options.MaxInputMegabytes) * 1024L * 1024L Then
                Throw New BatchJobPermanentException("Input exceeds the configured size limit.")
            End If
            Dim request As New BatchAttemptRequest With {
                .Options = job.Options, .JobId = job.JobId, .ItemIndex = index + 1, .Total = job.Items.Count,
                .Attempt = attempt.Number, .StagingDirectory = stage,
                .OriginalSourceName = System.IO.Path.GetFileName(item.SourcePath),
                .LogPath = System.IO.Path.Combine(attemptDir, "tooling.log")}
            If job.Options.CarryForward AndAlso previous IsNot Nothing Then
                For Each output As BatchJobOutput In previous.Outputs
                    Dim source As System.String = VerifyStoredOutput(root, output, ct)
                    Dim destination As System.String = UniquePath(stage, System.IO.Path.GetFileName(source))
                    System.IO.File.Copy(source, destination, False)
                    If HashFile(destination, ct) <> output.Sha256 Then Throw New BatchJobPermanentException("Carried output changed while copying.")
                    request.CarryPaths.Add(destination)
                Next
                If previous.Response IsNot Nothing Then
                    Dim source As System.String = VerifyStoredOutput(root, previous.Response, ct)
                    request.PreviousResponsePath = UniquePath(stage, "previous_response.md")
                    System.IO.File.Copy(source, request.PreviousResponsePath, False)
                    If HashFile(request.PreviousResponsePath, ct) <> previous.Response.Sha256 Then Throw New BatchJobPermanentException("Previous response changed while copying.")
                End If
            End If
            request.InputPath = UniquePath(stage, "input_" & System.IO.Path.GetFileName(item.SourcePath))
            System.IO.File.Copy(item.SourcePath, request.InputPath, False)
            Dim hash As System.String = HashFile(request.InputPath, ct)
            If item.SourceSha256 <> "" AndAlso item.SourceSha256 <> hash Then
                Throw New BatchJobPermanentException("The source file changed since this job's input snapshot. Start a new job instead of mixing versions.")
            End If
            item.SourceSha256 = hash
            Return request
        End Function

        Private Shared Sub Commit(root As System.String, item As BatchJobItem, request As BatchAttemptRequest,
                                  result As BatchAttemptResult, ct As System.Threading.CancellationToken)
            Dim relative As System.String = System.IO.Path.Combine("results", request.ItemIndex.ToString("D5", System.Globalization.CultureInfo.InvariantCulture) & "_" & System.Guid.NewGuid().ToString("N"))
            Dim finalDir As System.String = ResolveStoredPath(root, relative)
            Dim pending As System.String = finalDir & ".pending"
            System.IO.Directory.CreateDirectory(pending)
            Dim outputs As New System.Collections.Generic.List(Of BatchJobOutput)()
            For Each output As System.String In result.OutputPaths
                ct.ThrowIfCancellationRequested()
                If Not IsWithin(output, request.StagingDirectory) Then
                    Throw New BatchJobPermanentException("A result is outside this attempt's isolated staging directory: " & output)
                End If
                RejectReparseAncestors(output)
                If Not System.IO.File.Exists(output) Then Throw New System.IO.FileNotFoundException("A declared result file is missing.", output)
                If New System.IO.FileInfo(output).Length = 0 Then Throw New System.IO.InvalidDataException("A declared result file is empty: " & output)
                Dim destination As System.String = UniquePath(pending, System.IO.Path.GetFileName(output))
                Dim expectedHash As System.String = HashFile(output, ct)
                System.IO.File.Copy(output, destination, False)
                If HashFile(destination, ct) <> expectedHash Then Throw New System.IO.IOException("Result changed during commit: " & output)
                outputs.Add(New BatchJobOutput With {.RelativePath = System.IO.Path.Combine(relative, System.IO.Path.GetFileName(destination)), .Sha256 = expectedHash})
            Next
            Dim responsePath As System.String = UniquePath(pending, "_agent_response.md")
            System.IO.File.WriteAllText(responsePath, result.Response, New System.Text.UTF8Encoding(False))
            Dim response As New BatchJobOutput With {.RelativePath = System.IO.Path.Combine(relative, System.IO.Path.GetFileName(responsePath)), .Sha256 = HashFile(responsePath, ct)}
            ct.ThrowIfCancellationRequested()
            System.IO.Directory.Move(pending, finalDir)
            item.Outputs = outputs
            item.Response = response
        End Sub

        Public Shared Function VerifyStoredOutput(root As System.String, output As BatchJobOutput,
                                                 ct As System.Threading.CancellationToken) As System.String
            If output Is Nothing OrElse System.String.IsNullOrWhiteSpace(output.Sha256) Then Throw New BatchJobPermanentException("A recovery output has no fingerprint.")
            Dim path As System.String = ResolveStoredPath(root, output.RelativePath)
            RejectReparseAncestors(path)
            If Not System.IO.File.Exists(path) OrElse HashFile(path, ct) <> output.Sha256 Then
                Throw New BatchJobPermanentException("A committed output was removed or modified: " & path)
            End If
            Return path
        End Function

        Public Shared Function ResolveStoredPath(root As System.String, relative As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(relative) OrElse System.IO.Path.IsPathRooted(relative) Then Throw New System.IO.InvalidDataException("Invalid relative job path.")
            Dim full As System.String = System.IO.Path.GetFullPath(System.IO.Path.Combine(root, relative))
            If Not IsWithin(full, root) Then Throw New System.IO.InvalidDataException("Job path escapes its directory.")
            RejectReparseAncestors(full)
            Return full
        End Function

        Public Shared Function IsWithin(path As System.String, root As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(path) OrElse System.String.IsNullOrWhiteSpace(root) Then Return False
            Dim prefix As System.String = System.IO.Path.GetFullPath(root).TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar) & System.IO.Path.DirectorySeparatorChar
            Return System.IO.Path.GetFullPath(path).StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Sub RejectReparseAncestors(path As System.String)
            Dim current As System.String = System.IO.Path.GetFullPath(path)
            Do While Not System.String.IsNullOrWhiteSpace(current)
                If (System.IO.File.Exists(current) OrElse System.IO.Directory.Exists(current)) AndAlso
                   (System.IO.File.GetAttributes(current) And System.IO.FileAttributes.ReparsePoint) <> 0 Then
                    Throw New BatchJobPermanentException("Reparse points/symbolic links are not accepted in batch paths: " & current)
                End If
                Dim parent As System.IO.DirectoryInfo = System.IO.Directory.GetParent(current)
                current = If(parent Is Nothing, Nothing, parent.FullName)
            Loop
        End Sub

        Private Shared Function UniquePath(directory As System.String, leaf As System.String) As System.String
            leaf = System.IO.Path.GetFileName(leaf)
            Dim result As System.String = System.IO.Path.Combine(directory, leaf)
            Dim number As System.Int32 = 1
            While System.IO.File.Exists(result) OrElse System.IO.Directory.Exists(result)
                result = System.IO.Path.Combine(directory, System.IO.Path.GetFileNameWithoutExtension(leaf) & "_" & number.ToString(System.Globalization.CultureInfo.InvariantCulture) & System.IO.Path.GetExtension(leaf))
                number += 1
            End While
            Return result
        End Function

        Private Shared Function HashFile(path As System.String, ct As System.Threading.CancellationToken) As System.String
            Using hash As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                    Dim buffer(81919) As System.Byte
                    Dim count As System.Int32 = stream.Read(buffer, 0, buffer.Length)
                    While count > 0
                        ct.ThrowIfCancellationRequested()
                        hash.TransformBlock(buffer, 0, count, buffer, 0)
                        count = stream.Read(buffer, 0, buffer.Length)
                    End While
                    hash.TransformFinalBlock(System.Array.Empty(Of System.Byte)(), 0, 0)
                End Using
                Return System.BitConverter.ToString(hash.Hash).Replace("-", "").ToLowerInvariant()
            End Using
        End Function

        Public Shared Sub WriteAtomic(path As System.String, content As System.String)
            Dim temporary As System.String = path & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
            Try
                Dim bytes As System.Byte() = New System.Text.UTF8Encoding(False).GetBytes(content)
                Using stream As New System.IO.FileStream(temporary, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None)
                    stream.Write(bytes, 0, bytes.Length)
                    stream.Flush(True)
                End Using
                If System.IO.File.Exists(path) Then
                    System.IO.File.Replace(temporary, path, Nothing)
                Else
                    System.IO.File.Move(temporary, path)
                End If
            Finally
                If System.IO.File.Exists(temporary) Then System.IO.File.Delete(temporary)
            End Try
        End Sub

        Private Shared Sub Save(manifestPath As System.String, job As BatchJobManifest)
            job.UpdatedUtc = System.DateTime.UtcNow
            WriteAtomic(manifestPath, Newtonsoft.Json.JsonConvert.SerializeObject(job, Newtonsoft.Json.Formatting.Indented))
            WriteReport(System.IO.Path.GetDirectoryName(manifestPath), job)
        End Sub

        Private Shared Sub WriteReport(root As System.String, job As BatchJobManifest)
            Dim csv As New System.Text.StringBuilder("Item,Source,Status,Attempt,Error,AttemptDirectory" & System.Environment.NewLine)
            Dim summary As New System.Text.StringBuilder("# Batch Agent Job" & System.Environment.NewLine & System.Environment.NewLine)
            summary.AppendLine("Status: " & job.Status)
            summary.AppendLine("Model: " & job.Options.ModelDisplayName)
            summary.AppendLine("Carry forward: " & job.Options.CarryForward.ToString())
            summary.AppendLine()
            For index As System.Int32 = 0 To job.Items.Count - 1
                Dim item As BatchJobItem = job.Items(index)
                summary.AppendLine((index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) & ". " & System.IO.Path.GetFileName(item.SourcePath) & " — " & item.Status & "; attempts=" & item.Attempts.Count.ToString(System.Globalization.CultureInfo.InvariantCulture))
                If item.ErrorMessage <> "" Then summary.AppendLine("   " & item.ErrorMessage.Replace(System.Environment.NewLine, " "))
                For Each attempt As BatchJobAttempt In item.Attempts
                    If attempt.Status = "succeeded" Then Continue For
                    csv.AppendLine(System.String.Join(",", New System.String() {
                        (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), CsvField(item.SourcePath), CsvField(attempt.Status),
                        attempt.Number.ToString(System.Globalization.CultureInfo.InvariantCulture), CsvField(attempt.ErrorDetail), CsvField(attempt.DirectoryRelativePath)}))
                Next
                If item.Attempts.Count = 0 AndAlso item.ErrorMessage <> "" Then csv.AppendLine(System.String.Join(",", New System.String() {
                    (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), CsvField(item.SourcePath), CsvField(item.Status), "0", CsvField(item.ErrorMessage), ""}))
            Next
            If job.DiscoveryWarnings.Count > 0 Then
                summary.AppendLine()
                summary.AppendLine("## Discovery warnings (not processed)")
                For Each warning As System.String In job.DiscoveryWarnings
                    summary.AppendLine(warning)
                Next
            End If
            summary.AppendLine()
            summary.AppendLine("Results and attempts are immutable snapshots. Resume uses this manifest and verifies source/carried-output fingerprints. Failed attempts are retained for diagnosis, never carried forward.")
            WriteAtomic(System.IO.Path.Combine(root, "summary.md"), summary.ToString())
            WriteAtomic(System.IO.Path.Combine(root, "errors.csv"), System.Char.ConvertFromUtf32(&HFEFF) & csv.ToString())
        End Sub

        Private Shared Function CsvField(value As System.String) As System.String
            value = If(value, "")
            ' Do not let error/source text become a formula when the report is opened in Excel.
            Dim trimmed As System.String = value.TrimStart()
            If trimmed.Length > 0 AndAlso "=+-@".IndexOf(trimmed(0)) >= 0 Then value = "'" & value
            Return """" & value.Replace("""", """""") & """"
        End Function

        Private Shared Sub Report(progress As System.IProgress(Of BatchJobProgress), job As BatchJobManifest,
                                  index As System.Int32, attempt As System.Int32, message As System.String)
            If progress Is Nothing Then Return
            Dim completed As System.Int32 = 0
            For Each item As BatchJobItem In job.Items
                If item.Status = "succeeded" OrElse item.Status = "failed" Then completed += 1
            Next
            progress.Report(New BatchJobProgress With {.Total = job.Items.Count, .Completed = completed,
                .ItemIndex = index + 1, .Attempt = attempt,
                .FileName = If(index >= 0, System.IO.Path.GetFileName(job.Items(index).SourcePath), ""), .Message = message})
        End Sub

        Private NotInheritable Class InlineProgress
            Implements System.IProgress(Of System.String)
            Private ReadOnly _report As System.Action(Of System.String)
            Public Sub New(report As System.Action(Of System.String))
                _report = report
            End Sub
            Public Sub Report(value As System.String) Implements System.IProgress(Of System.String).Report
                _report(value)
            End Sub
        End Class
    End Class
End Namespace
