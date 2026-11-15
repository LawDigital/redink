' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Durable batch inputs, committed checkpoints and publication. No model/provider logic.
Option Explicit On
Option Strict On
Option Infer On
Imports System.Linq

Namespace BatchJobs
    Friend NotInheritable Class Settings
        Public Property Schema As System.Int32 = 1
        Public Property SourceDirectory As System.String = ""
        Public Property Prompt As System.String = ""
        Public Property LastPrompt As System.String = ""
        Public Property ModelKey As System.String = ""
        Public Property Retries As System.Int32 = 1
        Public Property MaxIterations As System.Int32 = 100
        Public Property OutputOnDesktop As System.Boolean = True
        Public Property CarryForward As System.Boolean = False
        Public Property ConsistentLayout As System.Boolean = True
        Public Property IncludeSubdirectories As System.Boolean = False
        Public Property StopOnError As System.Boolean = False
        Public Property OpenResults As System.Boolean = True
        ' Retain the setting for older settings files; selection now defaults to Local Chat.
        Public Property IncludeSelectedTools As System.Boolean = True
        Public Property UseLocalChatSelection As System.Boolean = True
        Public Property SelectedMainToolNames As New System.Collections.Generic.List(Of System.String)()
        Public Property SelectedAdvancedToolNames As New System.Collections.Generic.List(Of System.String)()
        Public Property Extensions As System.String = ""
        Public Property LastManifestPath As System.String = ""
    End Class

    Friend NotInheritable Class Item
        Public Property Number As System.Int32
        Public Property SourcePath As System.String = ""
        Public Property RelativeName As System.String = ""
        ' Persist physical history names independently from the complete source identity.
        Public Property StorageDirectoryName As System.String = ""
        Public Property SourceSha256 As System.String = ""
        Public Property Status As System.String = "pending"
        Public Property AttemptCount As System.Int32
        Public Property Snapshot As System.String = ""
        Public Property ErrorMessage As System.String = ""
        Public Property Outputs As New System.Collections.Generic.List(Of OutputFile)()
        Public Property Attempts As New System.Collections.Generic.List(Of Attempt)()
    End Class

    Friend NotInheritable Class OutputFile
        Public Property Name As System.String = ""
        Public Property Sha256 As System.String = ""
        ' Optional for schema-1 checkpoints created before stable slot binding existed.
        Public Property LogicalDeliverableId As System.String = ""
        Public Property OutputSlotId As System.String = ""
    End Class

    Friend NotInheritable Class Attempt
        Public Property Number As System.Int32
        Public Property StartedUtc As System.DateTime
        Public Property FinishedUtc As System.DateTime
        Public Property Status As System.String = "running"
        Public Property ErrorMessage As System.String = ""
        Public Property LogPath As System.String = ""
        Public Property Phase As System.String = "input_preparation"
        ' Keep the actual user answer even when a later output check fails.
        Public Property UserResponse As System.String = ""
        Public Property FailureType As System.String = ""
        Public Property ToolingStatus As System.String = ""
        Public Property IssueCode As System.String = ""
    End Class

    Friend NotInheritable Class Manifest
        Public Property Schema As System.Int32 = 1
        ' Older manifests omit this marker and keep their original physical layout.
        Public Property StorageLayoutVersion As System.Int32
        Public Property CreatedUtc As System.DateTime = System.DateTime.UtcNow
        Public Property RootDirectory As System.String = ""
        Public Property StorageNotice As System.String = ""
        Public Property Options As Settings = New Settings()
        Public Property Items As New System.Collections.Generic.List(Of Item)()
        Public Property ContractJson As System.String = ""
        Public Property ModelFingerprint As System.String = ""
        Public Property ToolNames As New System.Collections.Generic.List(Of System.String)()
        Public Property EnvironmentVersion As System.Int32
        Public Property EnvironmentFingerprint As System.String = ""
        Public Property SourceChatId As System.Int32
        Public Property ReadOnlyRoots As New System.Collections.Generic.List(Of System.String)()
        Public Property ResourceFiles As New System.Collections.Generic.List(Of ResourceFile)()
        Public Property FinishedUtc As System.DateTime?
        Public Property RunStatus As System.String = "ready"
        Public Property PublicationStatus As System.String = "pending"
        Public Property RunError As System.String = ""
        <Newtonsoft.Json.JsonIgnore>
        Public Property ReportWarning As System.String = ""
    End Class

    ' Only identities/checksums enter Job.json. Keep prompt text and executable tool
    ' configurations in memory so the checkpoint never stores tool credentials.
    Friend NotInheritable Class ResourceFile
        Public Property ToolName As System.String = ""
        Public Property Path As System.String = ""
        Public Property Sha256 As System.String = ""
    End Class

    Friend NotInheritable Class LocalChatEnvironment
        Public Property SourceChatId As System.Int32
        Public Property DeselectedMainToolNames As New System.Collections.Generic.List(Of System.String)()
        Public Property DeselectedAdvancedToolNames As New System.Collections.Generic.List(Of System.String)()
        Public Property PreviouslySelectedToolNames As New System.Collections.Generic.List(Of System.String)()
        Public Property Tools As New System.Collections.Generic.List(Of Global.SharedLibrary.SharedLibrary.ModelConfig)()
        Public Property SystemPrompt As System.String = ""
        Public Property UserLanguage As System.String = ""
        Public Property CentralResourceRoot As System.String = ""
        Public Property LocalResourceRoot As System.String = ""
        Public Property ReferenceWorkspaceRoot As System.String = ""
        Public Property ReadOnlyRoots As New System.Collections.Generic.List(Of System.String)()
        Public Property ResourceFiles As New System.Collections.Generic.List(Of ResourceFile)()
        Public Property Fingerprint As System.String = ""
        Public Property UseDefaultSemanticArchives As System.Boolean = False
        Public Property SemanticArchiveScope As Global.SharedLibrary.SharedLibrary.SemanticArchiveRunScope
    End Class

    Friend NotInheritable Class ProgressInfo
        Public Property Completed As System.Int32
        Public Property Total As System.Int32
        Public Property Failed As System.Int32
        Public Property Current As System.String = ""
        Public Property Message As System.String = ""
        Public Property RootDirectory As System.String = ""
    End Class

    Friend NotInheritable Class ModelChoice
        Public Property Key As System.String
        Public Property Caption As System.String
        Public Property Configuration As Global.SharedLibrary.SharedLibrary.ModelConfig
        Public Overrides Function ToString() As System.String
            Return Caption
        End Function
    End Class

    ' One immutable naming policy serves creation, budgets, retries and report links.
    Friend NotInheritable Class StorageLayout
        Public ReadOnly Property IsCompact As System.Boolean
        Public ReadOnly Property HistoryDirectoryName As System.String
        Public ReadOnly Property WorkDirectoryName As System.String
        Public ReadOnly Property InputsDirectoryName As System.String
        Public ReadOnly Property CandidateDirectoryName As System.String
        Public ReadOnly Property CommittedDirectoryName As System.String

        Friend Sub New(compact As System.Boolean)
            IsCompact = compact
            HistoryDirectoryName = If(compact, "h", "History")
            WorkDirectoryName = If(compact, "w", "Work")
            InputsDirectoryName = If(compact, "in", "Inputs")
            CandidateDirectoryName = If(compact, "c", "Candidate")
            CommittedDirectoryName = If(compact, "ok", "Committed")
        End Sub

        Public Function JobDirectoryName(createdLocal As System.DateTime) As System.String
            Return createdLocal.ToString(If(IsCompact, "yyMMdd_HHmmss", "yyMMdd_HH-mm-ss"), System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Public Function AttemptDirectoryName(number As System.Int32) As System.String
            If number < 1 Then Throw New System.ArgumentOutOfRangeException(NameOf(number))
            Return If(IsCompact, "a", "Attempt ") & number.ToString("00", System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Public Function CollisionDirectoryName(stem As System.String, number As System.Int32) As System.String
            If number < 2 Then Throw New System.ArgumentOutOfRangeException(NameOf(number))
            Dim suffix As System.String = number.ToString(System.Globalization.CultureInfo.InvariantCulture)
            Return stem & If(IsCompact, "_" & suffix, " (" & suffix & ")")
        End Function

        Public ReadOnly Property MaximumAttemptDirectoryName As System.String
            Get
                Return CollisionDirectoryName(AttemptDirectoryName(System.Int32.MaxValue), System.Int32.MaxValue)
            End Get
        End Property

        Public Function InputCopyFileName(sourcePath As System.String) As System.String
            Return If(IsCompact, "src" & System.IO.Path.GetExtension(sourcePath), "Input - " & System.IO.Path.GetFileName(sourcePath))
        End Function

        Public Function ReservedInputFileName(sourcePath As System.String) As System.String
            ' UniqueCopyPath allocates collision suffixes through _99999; preserve the whole extension.
            Return If(IsCompact, "src_99999", "Input") & System.IO.Path.GetExtension(sourcePath)
        End Function
    End Class

    Friend NotInheritable Class Store
        Public Const ContractFileName As System.String = "batch-output-contract.json"
        Public Const ContextFileName As System.String = "batch-context.json"
        Public Const StatusDirectoryName As System.String = "Batch status"
        Private Const OutputTemporarySuffix As System.String = ".native-append.pending"
        Private Shared ReadOnly LegacyLayout As New StorageLayout(False)
        Private Shared ReadOnly CompactLayout As New StorageLayout(True)
        Private Sub New()
        End Sub

        Public Shared Function LayoutFor(manifest As Manifest) As StorageLayout
            If manifest IsNot Nothing Then
                If manifest.Schema = 1 AndAlso manifest.StorageLayoutVersion = 0 Then Return LegacyLayout
                If manifest.Schema = 2 AndAlso manifest.StorageLayoutVersion = 1 Then Return CompactLayout
            End If
            Throw New System.IO.InvalidDataException("This batch uses an unsupported storage layout. Use the Red Ink version that created it, or choose Start batch. The host keeps its existing checkpoints.")
        End Function

        Public Shared ReadOnly Property SettingsPath As System.String
            Get
                Return System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "Red Ink", "BatchAgentJob", "settings.json")
            End Get
        End Property

        Public Shared Function LoadSettings() As Settings
            If Not System.IO.File.Exists(SettingsPath) Then Return New Settings()
            Dim result = ReadJson(Of Settings)(SettingsPath)
            If result Is Nothing OrElse result.Schema <> 1 Then Throw New System.IO.InvalidDataException("Unsupported Batch Agent Job settings.")
            ' Migrate the historical opt-in: a saved False must not drop selected skills.
            result.IncludeSelectedTools = True
            If result.SelectedMainToolNames Is Nothing Then result.SelectedMainToolNames = New System.Collections.Generic.List(Of System.String)()
            If result.SelectedAdvancedToolNames Is Nothing Then result.SelectedAdvancedToolNames = New System.Collections.Generic.List(Of System.String)()
            Return result
        End Function

        Public Shared Sub SaveSettings(value As Settings)
            WriteJson(SettingsPath, value)
        End Sub

        Public Shared Function ReadJson(Of T)(path As System.String) As T
            Using reader As New System.IO.StreamReader(path, System.Text.Encoding.UTF8, True)
                Using json As New Newtonsoft.Json.JsonTextReader(reader) With {.MaxDepth = 64}
                    Dim serializer = Newtonsoft.Json.JsonSerializer.Create(New Newtonsoft.Json.JsonSerializerSettings With {.TypeNameHandling = Newtonsoft.Json.TypeNameHandling.None})
                    Return serializer.Deserialize(Of T)(json)
                End Using
            End Using
        End Function

        Public Shared Sub WriteJson(path As System.String, value As System.Object)
            AtomicText(path, Newtonsoft.Json.JsonConvert.SerializeObject(value, Newtonsoft.Json.Formatting.Indented))
        End Sub

        Public Shared Sub AtomicText(path As System.String, value As System.String)
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path))
            Dim pending = path & ".pending"
            Using stream As New System.IO.FileStream(pending, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None)
                Dim bytes = (New System.Text.UTF8Encoding(False)).GetBytes(value)
                stream.Write(bytes, 0, bytes.Length)
                stream.Flush(True)
            End Using
            If System.IO.File.Exists(path) Then
                System.IO.File.Replace(pending, path, path & ".bak")
            Else
                System.IO.File.Move(pending, path)
            End If
        End Sub

        Public Shared Function Hash(path As System.String) As System.String
            Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                Using sha = System.Security.Cryptography.SHA256.Create()
                    Return System.BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant()
                End Using
            End Using
        End Function

        Public Shared Function ContainedPath(root As System.String, relative As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(relative) OrElse System.IO.Path.IsPathRooted(relative) Then Throw New System.IO.InvalidDataException("A batch path must be relative.")
            Dim fullRoot = System.IO.Path.GetFullPath(root).TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
            Dim full = System.IO.Path.GetFullPath(System.IO.Path.Combine(fullRoot & System.IO.Path.DirectorySeparatorChar, relative))
            If Not full.StartsWith(fullRoot & System.IO.Path.DirectorySeparatorChar, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.UnauthorizedAccessException("Batch path escaped its job directory.")
            Dim current = full
            While current.Length >= fullRoot.Length
                If (System.IO.File.Exists(current) OrElse System.IO.Directory.Exists(current)) AndAlso
                   (System.IO.File.GetAttributes(current) And System.IO.FileAttributes.ReparsePoint) <> 0 Then Throw New System.UnauthorizedAccessException("Batch paths cannot traverse symbolic links or junctions.")
                current = System.IO.Path.GetDirectoryName(current)
                If current Is Nothing Then Exit While
            End While
            Return full
        End Function

        Public Shared Function SafeName(value As System.String) As System.String
            Dim invalid = System.IO.Path.GetInvalidFileNameChars()
            Dim clean = (New System.String(If(value, "").Select(Function(c) If(invalid.Contains(c) OrElse System.Char.IsControl(c), "_"c, c)).ToArray())).Trim().TrimEnd("."c)
            If clean.Length > 75 Then clean = Global.SharedLibrary.Agents.FilePathBudget.UnicodePrefix(clean, 75).TrimEnd("."c, " "c)
            Return If(clean.Length = 0, "Document", clean)
        End Function

        Public Shared Function ItemDirectory(item As Item) As System.String
            If Not System.String.IsNullOrWhiteSpace(item.StorageDirectoryName) Then Return item.StorageDirectoryName
            ' Older checkpoints retain their recorded snapshots and report locations.
            Return item.Number.ToString("0000", System.Globalization.CultureInfo.InvariantCulture) & " - " & SafeName(System.IO.Path.GetFileNameWithoutExtension(item.RelativeName))
        End Function

        Private Shared Function StorageDirectoryBudget(root As System.String, item As Item, layout As StorageLayout) As System.Int32
            Dim longestChild As System.Int32 = System.Math.Max(
                System.IO.Path.Combine(layout.MaximumAttemptDirectoryName, layout.WorkDirectoryName, ContractFileName).Length,
                System.IO.Path.Combine(layout.MaximumAttemptDirectoryName, layout.WorkDirectoryName, layout.InputsDirectoryName, layout.ReservedInputFileName(item.SourcePath)).Length)
            longestChild = System.Math.Max(longestChild, System.IO.Path.Combine(layout.MaximumAttemptDirectoryName, layout.CandidateDirectoryName, "Response.md" & OutputTemporarySuffix).Length)
            Dim historyPrefix As System.Int32 = System.IO.Path.Combine(root, layout.HistoryDirectoryName, "x").Length - 1
            Dim reportPrefix As System.Int32 = System.IO.Path.Combine(root, "Results.previous", StatusDirectoryName, "x").Length - 1
            Dim reportChild As System.Int32 = System.IO.Path.Combine(layout.MaximumAttemptDirectoryName, "Response.md.pending").Length
            Dim prefixLength As System.Int32 = item.Number.ToString("0000", System.Globalization.CultureInfo.InvariantCulture).Length
            ' Compact jobs use only the inventory number; legacy jobs retain their bounded source label.
            Return System.Math.Min(If(layout.IsCompact, prefixLength, 24 + prefixLength + 3),
                System.Math.Min(Global.SharedLibrary.Agents.FilePathBudget.MaximumGeneratedPathLength - historyPrefix - 1 - longestChild,
                                Global.SharedLibrary.Agents.FilePathBudget.MaximumGeneratedPathLength - reportPrefix - 1 - reportChild))
        End Function

        Public Shared Sub PrepareItemStorage(manifest As Manifest, item As Item)
            Dim layout As StorageLayout = LayoutFor(manifest)
            If Not System.String.IsNullOrWhiteSpace(item.StorageDirectoryName) Then Return
            Dim prefix As System.String = item.Number.ToString("0000", System.Globalization.CultureInfo.InvariantCulture)
            Dim budget As System.Int32 = StorageDirectoryBudget(manifest.RootDirectory, item, layout)
            If budget < prefix.Length Then Throw New System.IO.PathTooLongException("The batch job folder leaves too little room for its working files. Start a new batch with Desktop output or select a source folder closer to the drive root. The host keeps existing checkpoints.")
            Dim name As System.String = prefix
            If Not layout.IsCompact AndAlso budget > prefix.Length + 3 Then
                Dim stem As System.String = SafeName(System.IO.Path.GetFileNameWithoutExtension(item.RelativeName))
                stem = Global.SharedLibrary.Agents.FilePathBudget.UnicodePrefix(stem, budget - prefix.Length - 3).TrimEnd("."c, " "c)
                If stem.Length > 0 Then name &= " - " & stem
            End If
            item.StorageDirectoryName = name
        End Sub

        Private Shared Function FitsJobStorage(root As System.String, items As System.Collections.Generic.IEnumerable(Of Item), layout As StorageLayout) As System.Boolean
            If Not Global.SharedLibrary.Agents.FilePathBudget.FitsGeneratedPath(root) Then Return False
            For Each item As Item In items
                If StorageDirectoryBudget(root, item, layout) < item.Number.ToString("0000", System.Globalization.CultureInfo.InvariantCulture).Length Then Return False
            Next
            Return True
        End Function

        Public Shared Function OutputFileNameLimit(manifest As Manifest, attemptRoot As System.String) As System.Int32
            Dim layout As StorageLayout = LayoutFor(manifest)
            Dim maximum As System.Int32 = Global.SharedLibrary.Agents.FilePathBudget.FileNameBudget(System.IO.Path.Combine(attemptRoot, layout.CandidateDirectoryName), OutputTemporarySuffix.Length)
            maximum = System.Math.Min(maximum, Global.SharedLibrary.Agents.FilePathBudget.FileNameBudget(System.IO.Path.Combine(attemptRoot, layout.CommittedDirectoryName)))
            For Each item As Item In manifest.Items
                If item.Status = "succeeded" Then Continue For
                Dim futureAttempt As System.String = System.IO.Path.Combine(manifest.RootDirectory, layout.HistoryDirectoryName, ItemDirectory(item), layout.MaximumAttemptDirectoryName)
                maximum = System.Math.Min(maximum, Global.SharedLibrary.Agents.FilePathBudget.FileNameBudget(System.IO.Path.Combine(futureAttempt, layout.WorkDirectoryName)))
                maximum = System.Math.Min(maximum, Global.SharedLibrary.Agents.FilePathBudget.FileNameBudget(System.IO.Path.Combine(futureAttempt, layout.CandidateDirectoryName), OutputTemporarySuffix.Length))
                maximum = System.Math.Min(maximum, Global.SharedLibrary.Agents.FilePathBudget.FileNameBudget(System.IO.Path.Combine(futureAttempt, layout.CommittedDirectoryName)))
            Next
            Dim prefixLength As System.Int32 = If(manifest.Options.CarryForward, 0, manifest.Items.Count.ToString("0000", System.Globalization.CultureInfo.InvariantCulture).Length + 3)
            maximum = System.Math.Min(maximum, Global.SharedLibrary.Agents.FilePathBudget.FileNameBudget(System.IO.Path.Combine(manifest.RootDirectory, "Results"), ".publishing".Length) - prefixLength)
            maximum = System.Math.Min(maximum, Global.SharedLibrary.Agents.FilePathBudget.FileNameBudget(System.IO.Path.Combine(manifest.RootDirectory, "Results.pending")) - prefixLength)
            maximum = System.Math.Min(maximum, Global.SharedLibrary.Agents.FilePathBudget.FileNameBudget(System.IO.Path.Combine(manifest.RootDirectory, "Results.previous")) - prefixLength)
            Return maximum
        End Function

        Public Shared Sub RequireOutputFileName(manifest As Manifest, attemptRoot As System.String, name As System.String)
            Dim maximum As System.Int32 = OutputFileNameLimit(manifest, attemptRoot)
            If name.Length > maximum Then Throw New System.IO.InvalidDataException("Use a final output filename of at most " & maximum.ToString(System.Globalization.CultureInfo.InvariantCulture) & " characters including its extension. The batch must preserve that name through validation, later inputs and publication; it cannot silently rename a committed output.")
        End Sub

        Public Shared Sub ValidateOptions(value As Settings)
            If value Is Nothing OrElse Not System.IO.Directory.Exists(value.SourceDirectory) Then Throw New System.IO.DirectoryNotFoundException("Select an existing source directory.")
            If System.String.IsNullOrWhiteSpace(value.Prompt) OrElse value.Prompt.Length > 100000 Then Throw New System.IO.InvalidDataException("Enter an instruction of 1 to 100,000 characters.")
            If value.Retries < 0 OrElse value.Retries > 10 OrElse value.MaxIterations < 1 OrElse value.MaxIterations > 1000 Then Throw New System.IO.InvalidDataException("Retries must be 0..10; maximum tool iterations must be 1..1000.")
            If System.String.IsNullOrWhiteSpace(value.ModelKey) Then Throw New System.IO.InvalidDataException("Select a tooling-capable model.")
            If (System.IO.File.GetAttributes(value.SourceDirectory) And System.IO.FileAttributes.ReparsePoint) <> 0 Then Throw New System.IO.InvalidDataException("The source root cannot be a junction or symbolic link.")
        End Sub

        Public Shared Function Create(options As Settings, ct As System.Threading.CancellationToken) As Manifest
            ValidateOptions(options)
            Dim sourceRoot = System.IO.Path.GetFullPath(options.SourceDirectory).TrimEnd(System.IO.Path.DirectorySeparatorChar) & System.IO.Path.DirectorySeparatorChar
            Dim desktopDirectory As System.String = System.IO.Path.GetFullPath(System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory)).TrimEnd(System.IO.Path.DirectorySeparatorChar)
            Dim desktopRoot As System.String = desktopDirectory & System.IO.Path.DirectorySeparatorChar & "Inky_Batch_Jobs"
            Dim sourceOutputRoot As System.String = sourceRoot & "Inky_Batch_Jobs"
            Dim destinationRoot As System.String = If(options.OutputOnDesktop, desktopRoot, sourceOutputRoot)
            ' The parent paths are already canonical. Do not normalize a fabricated long
            ' output child before the budget can select a shorter output root.
            Dim excludedRoots = New System.String() {
                desktopRoot & System.IO.Path.DirectorySeparatorChar,
                sourceOutputRoot & System.IO.Path.DirectorySeparatorChar
            }
            Dim extensions As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each part In If(options.Extensions, "").Split(New System.Char() {";"c, ","c, " "c}, System.StringSplitOptions.RemoveEmptyEntries)
                Dim ext = part.Trim().TrimStart("*"c)
                If Not ext.StartsWith(".", System.StringComparison.Ordinal) Then ext = "." & ext
                extensions.Add(ext)
            Next
            Dim files As New System.Collections.Generic.List(Of System.String)()
            Dim pending As New System.Collections.Generic.Stack(Of System.String)()
            pending.Push(sourceRoot)
            While pending.Count > 0
                ct.ThrowIfCancellationRequested()
                Dim directory = pending.Pop()
                For Each path In System.IO.Directory.GetFiles(directory)
                    Dim attributes = System.IO.File.GetAttributes(path)
                    If (attributes And (System.IO.FileAttributes.Hidden Or System.IO.FileAttributes.System Or System.IO.FileAttributes.ReparsePoint)) <> 0 OrElse System.IO.Path.GetFileName(path).StartsWith("~$", System.StringComparison.Ordinal) Then Continue For
                    If extensions.Count = 0 OrElse extensions.Contains(System.IO.Path.GetExtension(path)) Then files.Add(path)
                Next
                If options.IncludeSubdirectories Then
                    For Each child In System.IO.Directory.GetDirectories(directory)
                        If (System.IO.File.GetAttributes(child) And (System.IO.FileAttributes.Hidden Or System.IO.FileAttributes.System Or System.IO.FileAttributes.ReparsePoint)) <> 0 Then Continue For
                        Dim childRoot = System.IO.Path.GetFullPath(child).TrimEnd(System.IO.Path.DirectorySeparatorChar) & System.IO.Path.DirectorySeparatorChar
                        If excludedRoots.Any(Function(excluded) childRoot.StartsWith(excluded, System.StringComparison.OrdinalIgnoreCase)) Then Continue For
                        pending.Push(child)
                    Next
                End If
            End While
            files.Sort(System.StringComparer.OrdinalIgnoreCase)
            If files.Count = 0 Then Throw New System.IO.InvalidDataException("No matching source files were found.")
            ' Schema 2 prevents older builds from rewriting a compact job as a legacy layout.
            Dim manifest As New Manifest With {.Schema = 2, .StorageLayoutVersion = 1, .Options = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Settings)(Newtonsoft.Json.JsonConvert.SerializeObject(options))}
            Dim layout As StorageLayout = LayoutFor(manifest)
            manifest.Options.SourceDirectory = System.IO.Path.GetFullPath(options.SourceDirectory)
            For Each path In files
                ct.ThrowIfCancellationRequested()
                Dim item As New Item With {.Number = manifest.Items.Count + 1, .SourcePath = path, .RelativeName = path.Substring(sourceRoot.Length)}
                Try
                    item.SourceSha256 = Hash(path)
                Catch ex As System.IO.IOException
                    item.ErrorMessage = "Inventory could not read this file; its first readable attempt will pin its checksum: " & ex.Message
                Catch ex As System.UnauthorizedAccessException
                    item.ErrorMessage = "Inventory could not read this file; its first readable attempt will pin its checksum: " & ex.Message
                End Try
                manifest.Items.Add(item)
            Next
            Dim stamp As System.String = layout.JobDirectoryName(System.DateTime.Now)
            For Each candidateRoot As System.String In (New System.String() {destinationRoot, desktopRoot}).Distinct(System.StringComparer.OrdinalIgnoreCase)
                Dim candidate As System.String = candidateRoot & System.IO.Path.DirectorySeparatorChar & stamp
                Dim suffix As System.Int32 = 2
                While System.IO.Directory.Exists(candidate) OrElse System.IO.File.Exists(candidate)
                    candidate = candidateRoot & System.IO.Path.DirectorySeparatorChar & layout.CollisionDirectoryName(stamp, suffix)
                    If suffix = System.Int32.MaxValue Then Throw New System.IO.IOException("The host cannot allocate another job directory. Choose another output location and retry.")
                    suffix += 1
                End While
                If Not FitsJobStorage(candidate, manifest.Items, layout) Then Continue For
                manifest.RootDirectory = candidate
                If Not System.String.Equals(candidateRoot, destinationRoot, System.StringComparison.OrdinalIgnoreCase) Then
                    manifest.StorageNotice = "The source output location leaves too little room for the batch working files. This job uses Desktop output; Open results opens its actual folder. Original source files and their names remain unchanged."
                End If
                Exit For
            Next
            If System.String.IsNullOrWhiteSpace(manifest.RootDirectory) Then Throw New System.IO.PathTooLongException("Neither the selected output location nor Desktop leaves enough room for the batch working files. Select a source folder closer to the drive root and turn off Desktop output, then start a new batch.")
            For Each item As Item In manifest.Items
                PrepareItemStorage(manifest, item)
            Next
            System.IO.Directory.CreateDirectory(manifest.RootDirectory)
            System.IO.Directory.CreateDirectory(System.IO.Path.Combine(manifest.RootDirectory, "Results"))
            Save(manifest)
            Return manifest
        End Function

        Public Shared Function LoadManifest(path As System.String) As Manifest
            Dim result = ReadJson(Of Manifest)(path)
            If result Is Nothing OrElse result.Options Is Nothing OrElse result.Items Is Nothing OrElse result.Items.Count = 0 Then Throw New System.IO.InvalidDataException("Invalid batch checkpoint.")
            Dim layout As StorageLayout = LayoutFor(result)
            If Not System.String.Equals(System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(path)), System.IO.Path.GetFullPath(result.RootDirectory), System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("The batch checkpoint has moved. Restore its original job directory before retrying.")
            If result.Items.Any(Function(item) item Is Nothing OrElse item.Attempts Is Nothing OrElse item.Outputs Is Nothing) OrElse
               Not result.Items.Select(Function(item) item.Number).SequenceEqual(System.Linq.Enumerable.Range(1, result.Items.Count)) Then Throw New System.IO.InvalidDataException("Invalid batch input order.")
            Dim gap As System.Boolean = False
            Dim sources As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each item In result.Items
                Dim prefix As System.String = item.Number.ToString("0000", System.Globalization.CultureInfo.InvariantCulture)
                If layout.IsCompact Then
                    If item.StorageDirectoryName <> prefix Then Throw New System.IO.InvalidDataException("Invalid compact batch working-directory name.")
                ElseIf Not System.String.IsNullOrWhiteSpace(item.StorageDirectoryName) Then
                    If item.StorageDirectoryName <> System.IO.Path.GetFileName(item.StorageDirectoryName) OrElse
                       (item.StorageDirectoryName <> prefix AndAlso Not item.StorageDirectoryName.StartsWith(prefix & " - ", System.StringComparison.Ordinal)) Then
                        Throw New System.IO.InvalidDataException("Invalid batch working-directory name.")
                    End If
                End If
                Dim expectedSource = ContainedPath(result.Options.SourceDirectory, item.RelativeName)
                If Not System.String.Equals(expectedSource, System.IO.Path.GetFullPath(item.SourcePath), System.StringComparison.OrdinalIgnoreCase) OrElse Not sources.Add(expectedSource) Then Throw New System.IO.InvalidDataException("Invalid/duplicate source in the batch checkpoint.")
                If Not (New System.String() {"pending", "running", "failed", "cancelled", "succeeded"}).Contains(item.Status, System.StringComparer.Ordinal) OrElse
                   item.AttemptCount <> item.Attempts.Count OrElse item.Attempts.Any(Function(value) value Is Nothing) OrElse
                   Not item.Attempts.Select(Function(value) value.Number).SequenceEqual(System.Linq.Enumerable.Range(1, item.AttemptCount)) Then Throw New System.IO.InvalidDataException("Invalid item/attempt state in the batch checkpoint.")
                If item.Status <> "succeeded" Then
                    gap = True
                    If item.Status = "running" Then
                        item.Status = "failed"
                        item.ErrorMessage = "The process stopped before the host committed a verified checkpoint."
                        For Each attempt In item.Attempts.Where(Function(value) value.Status = "running")
                            attempt.Status = "interrupted"
                            attempt.ErrorMessage = item.ErrorMessage & " Resume detected the interruption."
                            attempt.FinishedUtc = System.DateTime.UtcNow
                        Next
                    End If
                Else
                    If gap AndAlso result.Options.CarryForward Then Throw New System.IO.InvalidDataException("Cumulative checkpoint has a gap; no automatic replay is safe.")
                    VerifySnapshot(result, item)
                End If
            Next
            Return result
        End Function

        Public Shared Sub VerifySnapshot(manifest As Manifest, item As Item)
            Dim snapshot = ContainedPath(manifest.RootDirectory, item.Snapshot)
            If item.Outputs Is Nothing OrElse item.Outputs.Count = 0 Then Throw New System.IO.InvalidDataException("A successful checkpoint must have outputs.")
            Global.SharedLibrary.Agents.OutputNameBinder.ValidateBindings(item.Outputs.Select(Function(output As OutputFile) As Global.SharedLibrary.Agents.OutputNameBinding
                If output Is Nothing Then Return Nothing
                Return New Global.SharedLibrary.Agents.OutputNameBinding With {.Name = output.Name, .LogicalDeliverableId = output.LogicalDeliverableId, .OutputSlotId = output.OutputSlotId}
            End Function))
            Dim names As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each output In item.Outputs
                If output Is Nothing OrElse output.Name <> System.IO.Path.GetFileName(output.Name) OrElse Not names.Add(output.Name) Then Throw New System.IO.InvalidDataException("Invalid output manifest.")
                Dim path = ContainedPath(snapshot, output.Name)
                If Not System.IO.File.Exists(path) OrElse Hash(path) <> output.Sha256 Then Throw New System.IO.InvalidDataException("A committed batch result is missing or has changed: " & output.Name)
            Next
        End Sub

        Public Shared Sub Save(manifest As Manifest)
            LayoutFor(manifest)
            ' Commit the authoritative state first. Reporting must never turn a committed
            ' item into a retryable model run, even when a viewer locks a Markdown file.
            WriteJson(ContainedPath(manifest.RootDirectory, "Job.json"), manifest)
            RefreshReports(manifest)
        End Sub

        Private Shared Function MarkdownText(value As System.String) As System.String
            Return If(value, "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").
                Replace("\", "\\").Replace("[", "\[").Replace("]", "\]").Replace("|", "\|").Replace("*", "\*").Replace("_", "\_").
                Replace(System.Environment.NewLine, " ").Replace(Microsoft.VisualBasic.vbLf, " ").Replace(Microsoft.VisualBasic.vbCr, " ")
        End Function

        Private Shared Function ReportLink(relative As System.String) As System.String
            Return System.String.Join("/", relative.Replace("\", "/").Split("/"c).
                Select(Function(segment) If(segment = "..", "..", System.Uri.EscapeDataString(segment))))
        End Function

        Private Shared Sub TryReportFile(path As System.String, text As System.String,
                                         warnings As System.Collections.Generic.List(Of System.String))
            Try
                If System.IO.File.Exists(path) AndAlso System.String.Equals(System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8), text, System.StringComparison.Ordinal) Then Return
                AtomicText(path, text)
            Catch ex As System.Exception
                Dim warning = "The host could not save " & path & ": " & ex.Message
                warnings.Add(warning)
                Try
                    System.Diagnostics.Trace.TraceWarning(warning)
                Catch
                    ' Report warnings must not re-enter the execution retry path.
                End Try
            End Try
        End Sub

        Public Shared Sub RefreshReports(manifest As Manifest)
            If manifest Is Nothing OrElse System.String.IsNullOrWhiteSpace(manifest.RootDirectory) Then Return
            Dim warnings As New System.Collections.Generic.List(Of System.String)()
            Try
                Dim layout As StorageLayout = LayoutFor(manifest)
                Dim report As New System.Text.StringBuilder("Item,Source,ItemStatus,AttemptStatus,Attempt,StartedUtc,FinishedUtc,Message,ToolingLog" & System.Environment.NewLine)
                For Each item In manifest.Items
                    If item.Attempts.Count = 0 Then report.AppendLine(System.String.Join(",", New System.String() {item.Number.ToString(System.Globalization.CultureInfo.InvariantCulture), item.RelativeName, item.Status, "not_started", "0", "", "", item.ErrorMessage, ""}.Select(Function(value) Csv(value))))
                    For Each attempt In item.Attempts
                        report.AppendLine(System.String.Join(",", New System.String() {item.Number.ToString(System.Globalization.CultureInfo.InvariantCulture), item.RelativeName, item.Status, attempt.Status, attempt.Number.ToString(System.Globalization.CultureInfo.InvariantCulture), attempt.StartedUtc.ToString("o"), If(attempt.FinishedUtc = System.DateTime.MinValue, "", attempt.FinishedUtc.ToString("o")), attempt.ErrorMessage, attempt.LogPath}.Select(Function(value) Csv(value))))
                        Dim attemptDirectory = System.IO.Path.GetDirectoryName(ContainedPath(manifest.RootDirectory, attempt.LogPath))
                        Dim status = BuildAttemptReport(manifest, item, attempt)
                        TryReportFile(ContainedPath(attemptDirectory, "Status.md"), status, warnings)
                        Dim visible = ContainedPath(manifest.RootDirectory, System.IO.Path.Combine("Results", StatusDirectoryName, ItemDirectory(item), layout.AttemptDirectoryName(attempt.Number)))
                        TryReportFile(ContainedPath(visible, "Status.md"), status, warnings)
                        If Not System.String.IsNullOrWhiteSpace(attempt.UserResponse) Then
                            Dim answer = "# Agent response" & System.Environment.NewLine & System.Environment.NewLine &
                                "Host status: **" & MarkdownText(attempt.Status) & "**. Read Status.md for completion and output checks; the answer alone does not prove success." &
                                System.Environment.NewLine & System.Environment.NewLine & attempt.UserResponse
                            TryReportFile(ContainedPath(attemptDirectory, "Response.md"), answer, warnings)
                            TryReportFile(ContainedPath(visible, "Response.md"), answer, warnings)
                        End If
                    Next
                Next
                TryReportFile(ContainedPath(manifest.RootDirectory, "Run-report.csv"), report.ToString(), warnings)
                TryReportFile(ContainedPath(manifest.RootDirectory, "Run-report.md"), BuildRunReport(manifest, ""), warnings)
                TryReportFile(ContainedPath(manifest.RootDirectory, System.IO.Path.Combine("Results", StatusDirectoryName, "Run-report.md")), BuildRunReport(manifest, "../../"), warnings)
            Catch ex As System.Exception
                warnings.Add("The host could not refresh the batch reports: " & ex.Message)
                Try
                    System.Diagnostics.Trace.TraceWarning(warnings.Last())
                Catch
                    ' Keep reporting non-throwing after the authoritative commit.
                End Try
            End Try
            manifest.ReportWarning = System.String.Join(System.Environment.NewLine, warnings)
        End Sub

        Private Shared Function BuildAttemptReport(manifest As Manifest, item As Item, attempt As Attempt) As System.String
            Dim text As New System.Text.StringBuilder()
            text.AppendLine("# Batch attempt status").AppendLine()
            text.AppendLine("Input: " & MarkdownText(item.RelativeName)).AppendLine()
            text.AppendLine("Attempt: " & attempt.Number.ToString(System.Globalization.CultureInfo.InvariantCulture))
            text.AppendLine("Status: **" & MarkdownText(attempt.Status) & "**")
            text.AppendLine("Last stage: " & MarkdownText(attempt.Phase))
            If attempt.ToolingStatus <> "" Then text.AppendLine("Agent workflow status: " & MarkdownText(attempt.ToolingStatus))
            If attempt.IssueCode <> "" Then text.AppendLine("Last technical issue: " & MarkdownText(attempt.IssueCode))
            text.AppendLine("Started (UTC): " & attempt.StartedUtc.ToString("o"))
            If attempt.FinishedUtc <> System.DateTime.MinValue Then text.AppendLine("Finished (UTC): " & attempt.FinishedUtc.ToString("o"))
            text.AppendLine()
            If attempt.Status = "succeeded" Then
                text.AppendLine("The host accepted this attempt and saved its checkpoint. Check Run-report.md for publication and overall coverage.")
            Else
                text.AppendLine("This attempt does not prove completion. The host keeps its working files separately; it does not promote them to a successful result.")
                If attempt.Status = "running" Then text.AppendLine("If no batch job is running, the process stopped before its closing update. Choose Retry / resume last job to check the checkpoint and continue.")
                text.AppendLine("Choose Open report for the overall status. Choose Retry / resume last job to continue unfinished inputs from the last successful checkpoint.")
                Select Case attempt.Phase
                    Case "input_preparation"
                        text.AppendLine("Check that the original input still exists and matches the recorded file. Start a new batch only when you intend to use changed source files.")
                    Case "output_validation"
                        text.AppendLine("Open the answer and working-files links in Run-report.md to inspect the current work. The host has not accepted this attempt's output; do not mistake it for the final result.")
                    Case "publication"
                        text.AppendLine("Close result files that an editor keeps open, then choose Retry / resume last job. The host republishes the successful checkpoint without another agent run.")
                End Select
                If attempt.FailureType <> "" Then text.AppendLine("Failure category: " & MarkdownText(attempt.FailureType))
            End If
            If Not System.String.IsNullOrWhiteSpace(attempt.UserResponse) Then text.AppendLine().AppendLine("Read Response.md for the complete agent answer. The host keeps that answer even when a later validation fails.")
            text.AppendLine().AppendLine("Use the Working files and Technical log links in Run-report.md to inspect this attempt; no file search is necessary.")
            Return text.ToString()
        End Function

        Private Shared Function BuildRunReport(manifest As Manifest, prefix As System.String) As System.String
            Dim layout As StorageLayout = LayoutFor(manifest)
            Dim succeeded = manifest.Items.Where(Function(item) item.Status = "succeeded").ToList()
            Dim failed = manifest.Items.Where(Function(item) item.Status = "failed" OrElse item.Status = "cancelled").ToList()
            Dim text As New System.Text.StringBuilder()
            text.AppendLine("# Batch Agent Job — run report").AppendLine()
            text.AppendLine("Status: **" & MarkdownText(manifest.RunStatus) & "**")
            text.AppendLine("Report time (UTC): " & System.DateTime.UtcNow.ToString("o"))
            text.AppendLine("Successful inputs: **" & succeeded.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " / " & manifest.Items.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & "**")
            text.AppendLine("Failed or cancelled inputs: " & failed.Count.ToString(System.Globalization.CultureInfo.InvariantCulture))
            text.AppendLine("Result publication: " & MarkdownText(manifest.PublicationStatus)).AppendLine()
            If Not System.String.IsNullOrWhiteSpace(manifest.StorageNotice) Then text.AppendLine(MarkdownText(manifest.StorageNotice)).AppendLine()
            If manifest.EnvironmentVersion = 1 Then
                Dim selectionSource As System.String = If(manifest.Options.UseLocalChatSelection, "Local Chat selection", "Batch selection")
                text.AppendLine("Agent configuration: " & selectionSource & "; context from Local Chat " & manifest.SourceChatId.ToString(System.Globalization.CultureInfo.InvariantCulture) & "; " & manifest.ToolNames.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " selected/session tools.")
                Dim resources = manifest.ResourceFiles.Where(Function(value) value IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(value.ToolName)).Select(Function(value) value.ToolName).ToList()
                If resources.Count > 0 Then text.AppendLine("Skills and agents: " & MarkdownText(System.String.Join(", ", resources)))
                text.AppendLine()
            End If
            If succeeded.Count < manifest.Items.Count Then text.AppendLine("**The results do not cover the entire input list.** Read the input overview below before using them.").AppendLine()
            If manifest.Options.CarryForward Then
                If succeeded.Count = 0 Then
                    text.AppendLine("The host has no successful cumulative checkpoint yet.")
                Else
                    Dim last = succeeded.Last()
                    text.AppendLine("The latest cumulative checkpoint includes inputs 1 through " & last.Number.ToString(System.Globalization.CultureInfo.InvariantCulture) & "; the last input is " & MarkdownText(last.RelativeName) & ".")
                End If
            End If
            text.AppendLine().AppendLine("## What you can do next").AppendLine()
            If manifest.RunStatus = "complete" AndAlso manifest.PublicationStatus = "current" Then
                text.AppendLine("Open Results for the output files. Open Batch status for the answers and status of each attempt.")
            Else
                text.AppendLine("Choose **Open report** in Batch Agent Job to read this report. Choose **Retry / resume last job** to continue unfinished inputs. The host skips successful inputs and rebuilds Results from their verified checkpoints.")
                If manifest.PublicationStatus <> "current" Then text.AppendLine("Close any result files that an editor keeps open, then choose **Retry / resume last job**. The host republishes successful checkpoints without rerunning their agents.")
                text.AppendLine("Review the failed attempt's answer and working files through the links below. Do not treat working files as checked results. A retry of an unfinished input can repeat its tool actions.")
                If manifest.RunStatus = "running" Then text.AppendLine("This report records a running job. If no job is running, choose Retry / resume last job; the host will identify the interrupted attempt.")
            End If
            text.AppendLine().AppendLine("## Input overview").AppendLine()
            For Each item In manifest.Items
                text.AppendLine("### " & item.Number.ToString(System.Globalization.CultureInfo.InvariantCulture) & ". " & MarkdownText(item.RelativeName))
                text.AppendLine("Status: **" & MarkdownText(item.Status) & "**").AppendLine()
                If item.Status = "succeeded" Then
                    For Each output In item.Outputs
                        Dim checkpoint = System.IO.Path.Combine(item.Snapshot, output.Name)
                        text.AppendLine("Checkpoint file: [" & MarkdownText(output.Name) & "](" & ReportLink(prefix & checkpoint) & ")")
                    Next
                End If
                For Each attempt In item.Attempts
                    Dim attemptRelative = System.IO.Path.GetDirectoryName(attempt.LogPath)
                    If System.String.IsNullOrWhiteSpace(attemptRelative) Then Continue For
                    text.Append("Attempt ").Append(attempt.Number.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append(" — ").Append(MarkdownText(attempt.Status)).Append(": ")
                    text.Append("[Status](").Append(ReportLink(prefix & System.IO.Path.Combine(attemptRelative, "Status.md"))).Append(")")
                    If Not System.String.IsNullOrWhiteSpace(attempt.UserResponse) OrElse System.IO.File.Exists(ContainedPath(manifest.RootDirectory, System.IO.Path.Combine(attemptRelative, "Response.md"))) Then
                        text.Append(" · [Answer](").Append(ReportLink(prefix & System.IO.Path.Combine(attemptRelative, "Response.md"))).Append(")")
                    End If
                    text.Append(" · [Working files — not verified](").Append(ReportLink(prefix & System.IO.Path.Combine(attemptRelative, layout.WorkDirectoryName))).Append(")")
                    text.Append(" · [Technical log](").Append(ReportLink(prefix & attempt.LogPath)).AppendLine(")")
                Next
                text.AppendLine()
            Next
            Return text.ToString()
        End Function

        Private Shared Function Csv(value As System.String) As System.String
            Dim text = If(value, "")
            If text.Length > 0 AndAlso "=+-@".IndexOf(text(0)) >= 0 Then text = "'" & text
            Return """" & text.Replace("""", """""") & """"
        End Function

        Public Shared Sub Publish(manifest As Manifest, Optional justCommitted As Item = Nothing)
            LayoutFor(manifest)
            manifest.PublicationStatus = "pending"
            If justCommitted IsNot Nothing AndAlso Not manifest.Options.CarryForward Then
                ' Independent items have disjoint, numbered names. Publish only this
                ' group during the run, rather than copying every earlier result again.
                ' Startup/end publication below still rebuilds the complete projection.
                VerifySnapshot(manifest, justCommitted)
                Dim results = ContainedPath(manifest.RootDirectory, "Results")
                System.IO.Directory.CreateDirectory(results)
                For Each output In justCommitted.Outputs
                    Dim name = justCommitted.Number.ToString("0000", System.Globalization.CultureInfo.InvariantCulture) & " - " & output.Name
                    Dim target = ContainedPath(results, name)
                    Dim staged = ContainedPath(results, name & ".publishing")
                    Global.SharedLibrary.Agents.FilePathBudget.RequireGeneratedPath(staged)
                    Try
                        System.IO.File.Copy(ContainedPath(ContainedPath(manifest.RootDirectory, justCommitted.Snapshot), output.Name), staged, True)
                        If Hash(staged) <> output.Sha256 Then Throw New System.IO.IOException("Result changed while publishing: " & output.Name)
                        If System.IO.File.Exists(target) Then
                            System.IO.File.Replace(staged, target, Nothing)
                        Else
                            System.IO.File.Move(staged, target)
                        End If
                    Finally
                        If System.IO.File.Exists(staged) Then System.IO.File.Delete(staged)
                    End Try
                Next
                manifest.PublicationStatus = "current"
                Save(manifest)
                Return
            End If
            Dim successful = manifest.Items.Where(Function(item) item.Status = "succeeded").ToList()
            If manifest.Options.CarryForward AndAlso successful.Count > 0 Then successful = New System.Collections.Generic.List(Of Item) From {successful.Last()}
            Dim pending = ContainedPath(manifest.RootDirectory, "Results.pending")
            Dim final = ContainedPath(manifest.RootDirectory, "Results")
            Dim previous = ContainedPath(manifest.RootDirectory, "Results.previous")
            If System.IO.Directory.Exists(pending) Then System.IO.Directory.Delete(pending, True)
            System.IO.Directory.CreateDirectory(pending)
            For Each item In successful
                VerifySnapshot(manifest, item)
                For Each output In item.Outputs
                    Dim name = If(manifest.Options.CarryForward, output.Name, item.Number.ToString("0000", System.Globalization.CultureInfo.InvariantCulture) & " - " & output.Name)
                    Dim target As System.String = ContainedPath(pending, name)
                    Global.SharedLibrary.Agents.FilePathBudget.RequireGeneratedPath(target)
                    System.IO.File.Copy(ContainedPath(ContainedPath(manifest.RootDirectory, item.Snapshot), output.Name), target, False)
                Next
            Next
            ' The manifest and immutable snapshots are authoritative. Publication can be
            ' rebuilt after a crash; never reuse a partially written Results directory.
            If System.IO.Directory.Exists(previous) Then System.IO.Directory.Delete(previous, True)
            If System.IO.Directory.Exists(final) Then System.IO.Directory.Move(final, previous)
            Try
                System.IO.Directory.Move(pending, final)
            Catch ex As System.Exception
                If Not System.IO.Directory.Exists(final) AndAlso System.IO.Directory.Exists(previous) Then System.IO.Directory.Move(previous, final)
                Throw
            End Try
            If System.IO.Directory.Exists(previous) Then System.IO.Directory.Delete(previous, True)
            manifest.PublicationStatus = "current"
            Save(manifest)
        End Sub
    End Class
End Namespace
