' Part of "Red Ink" (Red Ink Semantic Archive Worker)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Strict On
Option Explicit On
Option Infer On

Namespace SemanticArchiveWorker
    Friend NotInheritable Class WorkerOptions
        Friend Property ConfigurationSource As System.String = Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_SEMANTICARCHIVE_WORKER_CONFIGURATION_SOURCE
        Friend Property ArchiveSelectors As New System.Collections.Generic.List(Of System.String)()
        Friend Property AllArchives As System.Boolean
        Friend Property DocumentIds As New System.Collections.Generic.List(Of System.String)()
        Friend Property AllDocuments As System.Boolean
        Friend Property Operation As System.String = Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_SEMANTICARCHIVE_WORKER_OPERATION
        Friend Property OperationId As System.String = System.Guid.NewGuid().ToString("N")
        Friend Property LoopContinuously As System.Boolean
        Friend Property MaximumFiles As System.Int32 = Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_SEMANTICARCHIVE_WORKER_MAXIMUM_FILES
        Friend Property DiscoveryEntries As System.Int32 = Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_SEMANTICARCHIVE_WORKER_DISCOVERY_ENTRIES
        Friend Property DiscoverySeconds As System.Int32 = Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_SEMANTICARCHIVE_WORKER_DISCOVERY_SECONDS
        Friend Property IntervalSeconds As System.Int32 = Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_SEMANTICARCHIVE_WORKER_INTERVAL_SECONDS
        Friend Property MaximumCycles As System.Int32 = Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_SEMANTICARCHIVE_WORKER_MAXIMUM_CYCLES
        Friend Property MaximumSeconds As System.Int32 = Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_SEMANTICARCHIVE_WORKER_MAXIMUM_SECONDS
        Friend Property LogPath As System.String = Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_SEMANTICARCHIVE_WORKER_LOG_PATH
        Friend Property ShowHelp As System.Boolean

        Friend Shared Function Parse(arguments As System.String()) As WorkerOptions
            Dim result As New WorkerOptions()
            Dim modeChosen As System.Boolean = False
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim values As System.String() = If(arguments, New System.String() {})
            Dim index As System.Int32 = 0

            ' Friendly shorthand: redink-sa-worker.exe repair "Archive A" "Archive B"
            ' drains one or more selected archives with the standard Red Ink configuration
            ' and all documents. Positional archive selectors end at the first -- option.
            If values.Length > 0 AndAlso Not values(0).StartsWith("-", System.StringComparison.Ordinal) Then
                result.Operation = NormalizeOperation(values(0))
                index = 1
                While index < values.Length AndAlso Not values(index).StartsWith("--", System.StringComparison.Ordinal)
                    Dim selector As System.String = values(index).Trim()
                    If selector.Length = 0 Then Throw New System.ArgumentException("The shorthand syntax contains an empty archive name or stable ID.")
                    If result.ArchiveSelectors.Exists(Function(existing As System.String) System.String.Equals(existing, selector, System.StringComparison.OrdinalIgnoreCase)) Then Throw New System.ArgumentException("Duplicate archive selector.")
                    result.ArchiveSelectors.Add(selector)
                    index += 1
                End While
                If result.ArchiveSelectors.Count = 0 Then Throw New System.ArgumentException("The shorthand syntax requires at least one archive name or stable ID after the operation.")
                result.AllDocuments = True
                result.LoopContinuously = True
                modeChosen = True
            End If

            While index < values.Length
                Dim optionName As System.String = values(index)
                If optionName = "--help" OrElse optionName = "-h" Then
                    result.ShowHelp = True
                    Return result
                End If
                If optionName <> "--archive" AndAlso optionName <> "--document" AndAlso Not seen.Add(optionName) Then Throw New System.ArgumentException("A command-line option was supplied more than once.")
                Select Case optionName
                    Case "--once", "--loop"
                        If modeChosen Then Throw New System.ArgumentException("Choose at most one of --once or --loop; shorthand already implies --loop.")
                        modeChosen = True
                        result.LoopContinuously = optionName = "--loop"
                    Case "--all-archives"
                        result.AllArchives = True
                    Case "--all-documents"
                        result.AllDocuments = True
                    Case "--ini"
                        result.ConfigurationSource = NextValue(values, index)
                    Case "--archive"
                        Dim selector As System.String = NextValue(values, index).Trim()
                        If selector.Length = 0 Then Throw New System.ArgumentException("--archive requires an archive name or stable ID.")
                        If result.ArchiveSelectors.Exists(Function(existing As System.String) System.String.Equals(existing, selector, System.StringComparison.OrdinalIgnoreCase)) Then Throw New System.ArgumentException("Duplicate archive selector.")
                        result.ArchiveSelectors.Add(selector)
                    Case "--document"
                        Dim documentId As System.String = Global.SharedLibrary.SharedLibrary.SemanticArchiveIdentity.ValidateId(NextValue(values, index), "document")
                        If result.DocumentIds.Contains(documentId) Then Throw New System.ArgumentException("Duplicate document ID.")
                        result.DocumentIds.Add(documentId)
                    Case "--operation"
                        result.Operation = NormalizeOperation(NextValue(values, index))
                    Case "--operation-id"
                        Dim operationGuid As System.Guid
                        If Not System.Guid.TryParse(NextValue(values, index), operationGuid) Then Throw New System.ArgumentException("--operation-id must be a GUID.")
                        result.OperationId = operationGuid.ToString("N")
                    Case "--batch-files"
                        result.MaximumFiles = BoundedInteger(NextValue(values, index), 1, 4096, "--batch-files")
                    Case "--discovery-entries"
                        result.DiscoveryEntries = BoundedInteger(NextValue(values, index), 1, 10000, "--discovery-entries")
                    Case "--discovery-seconds"
                        result.DiscoverySeconds = BoundedInteger(NextValue(values, index), 1, 120, "--discovery-seconds")
                    Case "--interval-seconds"
                        result.IntervalSeconds = BoundedInteger(NextValue(values, index), 1, 86400, "--interval-seconds")
                    Case "--max-cycles"
                        result.MaximumCycles = BoundedInteger(NextValue(values, index), 1, 1000000, "--max-cycles")
                    Case "--max-seconds"
                        result.MaximumSeconds = BoundedInteger(NextValue(values, index), 1, 604800, "--max-seconds")
                    Case "--log"
                        result.LogPath = Global.SharedLibrary.SharedLibrary.SemanticArchivePathGuard.RequireWindowsCompatiblePath(NextValue(values, index))
                    Case Else
                        Throw New System.ArgumentException("Unknown command-line option. Use --help for the supported options.")
                End Select
                index += 1
            End While

            ' The worker's normal job is to drain one requested maintenance operation.
            If Not modeChosen Then result.LoopContinuously = True
            If Not result.AllDocuments AndAlso result.DocumentIds.Count = 0 Then result.AllDocuments = True
            If result.AllArchives = (result.ArchiveSelectors.Count > 0) Then Throw New System.ArgumentException("Choose --all-archives or one or more --archive names/IDs.")
            If result.AllDocuments = (result.DocumentIds.Count > 0) Then Throw New System.ArgumentException("Choose --all-documents or one or more --document IDs.")
            If result.DocumentIds.Count > 0 AndAlso (result.AllArchives OrElse result.ArchiveSelectors.Count <> 1) Then Throw New System.ArgumentException("Selected document IDs require exactly one --archive.")
            If result.DocumentIds.Count > 4096 OrElse result.ArchiveSelectors.Count > 4096 Then Throw New System.ArgumentException("An explicit scope may contain at most 4096 IDs/names.")
            If Not result.LoopContinuously AndAlso (seen.Contains("--interval-seconds") OrElse seen.Contains("--max-cycles")) Then Throw New System.ArgumentException("Interval and cycle options require --loop.")
            Return result
        End Function

        Private Shared Function NormalizeOperation(value As System.String) As System.String
            Dim operation As System.String = If(value, "").Trim().ToLowerInvariant()
            Select Case operation
                Case "refresh", "retry", "repair", "reindex", "extract", "permissions"
                    Return operation
                Case Else
                    Throw New System.ArgumentException("Operation must be refresh, retry, repair, reindex, extract or permissions.")
            End Select
        End Function

        Private Shared Function NextValue(values As System.String(), ByRef index As System.Int32) As System.String
            index += 1
            If index >= values.Length OrElse System.String.IsNullOrWhiteSpace(values(index)) OrElse values(index).StartsWith("--", System.StringComparison.Ordinal) Then Throw New System.ArgumentException("A command-line option is missing its value.")
            Return values(index)
        End Function

        Private Shared Function BoundedInteger(text As System.String, minimum As System.Int32, maximum As System.Int32, optionName As System.String) As System.Int32
            Dim result As System.Int32
            If Not System.Int32.TryParse(text, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, result) OrElse result < minimum OrElse result > maximum Then Throw New System.ArgumentException(optionName & " is outside its documented integer range.")
            Return result
        End Function
    End Class
End Namespace
