' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. For license to use see https://redink.ai.
' Reuse complete column definitions independently of documents, answers and review identities.
' The host retains library selections and authorizes destinations; this service handles JSON and files.
Option Explicit On
Option Strict On
Option Infer On

Namespace Global.SharedLibrary.TabularReview
    Public NotInheritable Class ReviewTemplateColumn
        Public Property Title As System.String = ""
        Public Property Instruction As System.String = ""
        Public Property Kind As System.String = "analysis"
        Public Property Options As New System.Collections.Generic.List(Of System.String)()
        Public Property Width As System.Nullable(Of System.Int32)
    End Class

    Public NotInheritable Class ReviewColumnTemplate
        Public Property Name As System.String = ""
        Public Property Description As System.String = ""
        Public Property ReportDefinitions As New System.Collections.Generic.List(Of ReviewReportDefinition)()
        Public Property Columns As New System.Collections.Generic.List(Of ReviewTemplateColumn)()
        Public Property DocumentColumnWidth As System.Nullable(Of System.Int32)
        Public Property FlagColors As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
    End Class

    ' File paths, hashes and entry ordinals belong to retained host state, never portable JSON.
    Public NotInheritable Class ReviewTemplateWriteTarget
        Public Property FilePath As System.String = ""
        Public Property FileHash As System.String = ""
        Public Property Exists As System.Boolean
    End Class

    Public NotInheritable Class ReviewTemplateFileSnapshot
        Public Property FilePath As System.String = ""
        Public Property FileHash As System.String = ""
        Public Property IsCatalog As System.Boolean
        Public Property Templates As New System.Collections.Generic.List(Of ReviewColumnTemplate)()
    End Class

    Public NotInheritable Class ReviewTemplateLibraryEntry
        Public Property DisplayName As System.String = ""
        ' personal | shared
        Public Property Scope As System.String = ""
        Public Property FilePath As System.String = ""
        Public Property FileHash As System.String = ""
        Public Property Index As System.Int32
        Public Property Template As ReviewColumnTemplate
    End Class

    Public NotInheritable Class ReviewTemplateLibraryNotice
        Public Property Scope As System.String = ""
        Public Property FileName As System.String = ""
        Public Property Message As System.String = ""
    End Class

    Public NotInheritable Class ReviewTemplateLibraryResult
        Public Property Entries As New System.Collections.Generic.List(Of ReviewTemplateLibraryEntry)()
        Public Property Notices As New System.Collections.Generic.List(Of ReviewTemplateLibraryNotice)()
    End Class

    Public NotInheritable Class ReviewTemplateService
        Private Const ColumnsFormat As System.String = "redink.tabular-review.columns"
        Private Const LibraryFormat As System.String = "redink.tabular-review.library"
        Private Const CurrentVersion As System.Int32 = 1
        Private Const DefaultLibraryFileName As System.String = "redink-tabular-review-templates.json"
        Private Const MaximumJsonBytes As System.Int32 = 32 * 1024 * 1024
        Private Const MaximumJsonDepth As System.Int32 = 16
        Private Shared ReadOnly StrictUtf8 As New System.Text.UTF8Encoding(False, True)

        Private Sub New()
        End Sub

        ' Normalize structural values only. Authored names, questions, descriptions,
        ' options and flag labels retain their exact text, including whitespace and case.
        Public Shared Function ValidateAndClone(template As ReviewColumnTemplate) As ReviewColumnTemplate
            If template Is Nothing Then Throw New System.IO.InvalidDataException("Choose a column template.")
            RequireText(template.Name, "template name", True)
            RequireText(template.Description, "template description", False)
            If template.Columns Is Nothing OrElse template.Columns.Count = 0 Then Throw New System.IO.InvalidDataException("A column template needs at least one column.")
            RequireWidth(template.DocumentColumnWidth, 160, 600, "document column")
            Dim result As New ReviewColumnTemplate With {
                .Name = template.Name, .Description = template.Description, .DocumentColumnWidth = template.DocumentColumnWidth
            }
            For Each column As ReviewTemplateColumn In template.Columns
                If column Is Nothing Then Throw New System.IO.InvalidDataException("A column template contains a missing column.")
                RequireText(column.Title, "column title", True)
                RequireText(column.Instruction, "column instruction", False)
                RequireText(column.Kind, "column type", True)
                Dim kind As System.String = column.Kind.Trim().ToLowerInvariant()
                Select Case kind
                    Case "analysis", "text", "priority", "flag"
                    Case Else
                        Throw New System.IO.InvalidDataException("Choose analysis, text, priority or flag as the column type.")
                End Select
                If kind = "analysis" AndAlso System.String.IsNullOrWhiteSpace(column.Instruction) Then Throw New System.IO.InvalidDataException("Every analysis column needs its complete question or instruction.")
                If column.Options Is Nothing Then Throw New System.IO.InvalidDataException("A column needs an options list, which may be empty.")
                RequireWidth(column.Width, 140, 1200, "review column")
                Dim copy As New ReviewTemplateColumn With {
                    .Title = column.Title, .Instruction = column.Instruction, .Kind = kind, .Width = column.Width
                }
                Dim optionValues As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                For Each optionText As System.String In column.Options
                    RequireText(optionText, "column option", True)
                    If Not optionValues.Add(optionText) Then Throw New System.IO.InvalidDataException("A column repeats an option. Keep each exact option only once.")
                    copy.Options.Add(optionText)
                Next
                result.Columns.Add(copy)
            Next
            If template.FlagColors Is Nothing Then Throw New System.IO.InvalidDataException("A column template needs a flag color object, which may be empty.")
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In template.FlagColors
                RequireText(pair.Key, "flag label", True)
                RequireText(pair.Value, "flag color", True)
                Dim color As System.String = pair.Value.Trim()
                If System.Text.RegularExpressions.Regex.IsMatch(color, "\A#[0-9a-fA-F]{3}\z") Then
                    color = "#" & New System.String(color(1), 2) & New System.String(color(2), 2) & New System.String(color(3), 2)
                End If
                If Not System.Text.RegularExpressions.Regex.IsMatch(color, "\A#[0-9a-fA-F]{6}\z") Then Throw New System.IO.InvalidDataException("Use a hexadecimal flag color such as #336699.")
                result.FlagColors.Add(pair.Key, color.ToUpperInvariant())
            Next
            Dim columnKeys As New System.Collections.Generic.List(Of System.String)()
            For index As System.Int32 = 0 To result.Columns.Count - 1
                columnKeys.Add(index.ToString(System.Globalization.CultureInfo.InvariantCulture))
            Next
            ReviewContract.ValidateReportDefinitions(template.ReportDefinitions, columnKeys)
            result.ReportDefinitions = Newtonsoft.Json.JsonConvert.DeserializeObject(Of System.Collections.Generic.List(Of ReviewReportDefinition))(ReviewContract.Serialize(template.ReportDefinitions), ReviewContract.JsonSettings)
            Return result
        End Function

        Public Shared Function ParseTemplateJson(json As System.String) As ReviewColumnTemplate
            Dim snapshot As ReviewTemplateFileSnapshot = ParseFileJson(json)
            If snapshot.IsCatalog Then Throw New System.IO.InvalidDataException("This JSON contains a template library. Choose one of its named templates from the library.")
            Return snapshot.Templates(0)
        End Function

        Public Shared Function SerializeTemplate(template As ReviewColumnTemplate) As System.String
            Return SerializeFile(New System.Collections.Generic.List(Of ReviewColumnTemplate) From {ValidateAndClone(template)}, False)
        End Function

        Public Shared Function ReadFile(filePath As System.String, cancellationToken As System.Threading.CancellationToken) As ReviewTemplateFileSnapshot
            cancellationToken.ThrowIfCancellationRequested()
            Dim fullPath As System.String = ResolvePath(filePath)
            Dim bytes As System.Byte() = ReadBytes(fullPath, cancellationToken)
            Dim snapshot As ReviewTemplateFileSnapshot = ParseFileJson(DecodeJson(bytes))
            cancellationToken.ThrowIfCancellationRequested()
            snapshot.FilePath = fullPath
            snapshot.FileHash = HashBytes(bytes)
            Return snapshot
        End Function

        Public Shared Function LoadLibraries(personalPath As System.String, sharedPath As System.String,
                                             cancellationToken As System.Threading.CancellationToken) As ReviewTemplateLibraryResult
            cancellationToken.ThrowIfCancellationRequested()
            Dim result As New ReviewTemplateLibraryResult()
            Dim labels As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            LoadScope(personalPath, "personal", result, labels, cancellationToken)
            LoadScope(sharedPath, "shared", result, labels, cancellationToken)
            cancellationToken.ThrowIfCancellationRequested()
            Return result
        End Function

        ' Existing files remain file targets regardless of their extension. For a
        ' missing configured path, .json means a file; any other suffix means a folder.
        Public Shared Function ResolveWriteTarget(configuredPath As System.String,
                                                  cancellationToken As System.Threading.CancellationToken) As ReviewTemplateWriteTarget
            cancellationToken.ThrowIfCancellationRequested()
            Dim fullPath As System.String = ResolvePath(configuredPath)
            Dim attributes As System.Nullable(Of System.IO.FileAttributes) = GetAttributesOrMissing(fullPath)
            Dim isDirectory As System.Boolean = If(attributes.HasValue,
                (attributes.GetValueOrDefault() And System.IO.FileAttributes.Directory) <> 0,
                Not System.String.Equals(System.IO.Path.GetExtension(fullPath), ".json", System.StringComparison.OrdinalIgnoreCase))
            If isDirectory Then fullPath = System.IO.Path.Combine(fullPath, DefaultLibraryFileName)
            Return InspectDestination(fullPath, cancellationToken)
        End Function

        ' Use this direct-file snapshot for an explicitly chosen JSON export destination.
        ' Empty FileHash means that the file did not exist; it never means unconditional overwrite.
        Public Shared Function InspectDestination(filePath As System.String,
                                                  cancellationToken As System.Threading.CancellationToken) As ReviewTemplateWriteTarget
            cancellationToken.ThrowIfCancellationRequested()
            Dim fullPath As System.String = ResolvePath(filePath)
            Dim bytes As System.Byte() = ReadOptionalBytes(fullPath, cancellationToken)
            Return New ReviewTemplateWriteTarget With {
                .FilePath = fullPath, .Exists = bytes IsNot Nothing,
                .FileHash = If(bytes Is Nothing, "", HashBytes(bytes))
            }
        End Function

        ' The host must supply its retained selection, not a browser-supplied path.
        Public Shared Function ReadEntry(entry As ReviewTemplateLibraryEntry,
                                         cancellationToken As System.Threading.CancellationToken) As ReviewColumnTemplate
            If entry Is Nothing Then Throw New System.IO.InvalidDataException("Choose a template from the current library list.")
            RequireExpectedHash(entry.FileHash)
            If entry.FileHash.Length = 0 Then Throw New System.IO.InvalidDataException("Refresh the library list and choose the template again.")
            Dim snapshot As ReviewTemplateFileSnapshot = ReadFile(entry.FilePath, cancellationToken)
            If Not System.String.Equals(snapshot.FileHash, entry.FileHash, System.StringComparison.OrdinalIgnoreCase) OrElse
               entry.Index < 0 OrElse entry.Index >= snapshot.Templates.Count Then
                Throw New System.IO.IOException("The template library changed. Refresh the library list and choose the template again.")
            End If
            Return ValidateAndClone(snapshot.Templates(entry.Index))
        End Function

        Public Shared Function SaveTemplate(filePath As System.String, template As ReviewColumnTemplate,
                                            expectedFileHash As System.String, cancellationToken As System.Threading.CancellationToken,
                                            Optional replaceName As System.String = Nothing,
                                            Optional validateDestination As System.Action(Of System.String) = Nothing) As ReviewTemplateFileSnapshot
            Return Publish(filePath, ValidateAndClone(template), expectedFileHash, cancellationToken, True, replaceName, validateDestination)
        End Function

        Public Shared Function ExportTemplate(filePath As System.String, template As ReviewColumnTemplate,
                                              expectedFileHash As System.String, cancellationToken As System.Threading.CancellationToken,
                                              Optional validateDestination As System.Action(Of System.String) = Nothing) As ReviewTemplateFileSnapshot
            Return Publish(filePath, ValidateAndClone(template), expectedFileHash, cancellationToken, False, Nothing, validateDestination)
        End Function

        Private Shared Function Publish(filePath As System.String, template As ReviewColumnTemplate,
                                        expectedFileHash As System.String, cancellationToken As System.Threading.CancellationToken,
                                        saveToLibrary As System.Boolean, replaceName As System.String,
                                        validateDestination As System.Action(Of System.String)) As ReviewTemplateFileSnapshot
            cancellationToken.ThrowIfCancellationRequested()
            RequireExpectedHash(expectedFileHash)
            Dim fullPath As System.String = ResolvePath(filePath)
            If validateDestination IsNot Nothing Then validateDestination(fullPath)
            If Not saveToLibrary AndAlso Not System.String.Equals(System.IO.Path.GetExtension(fullPath), ".json", System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.IO.InvalidDataException("Choose a .json filename for the column settings export.")
            End If
            Dim directory As System.String = System.IO.Path.GetDirectoryName(fullPath)
            Dim lockPath As System.String = System.IO.Path.Combine(directory, ".redink-columns.lock")
            Global.SharedLibrary.Agents.FilePathBudget.RequireGeneratedPath(lockPath)
            cancellationToken.ThrowIfCancellationRequested()
            System.IO.Directory.CreateDirectory(directory)
            ' Keep this zero-byte lock file after closing its handle. Deleting it can
            ' let another process lock a new file while an earlier handle still exists.
            Using writeLock As System.IO.FileStream = OpenWriteLock(lockPath)
                cancellationToken.ThrowIfCancellationRequested()
                If validateDestination IsNot Nothing Then validateDestination(fullPath)
                Dim existingBytes As System.Byte() = ReadOptionalBytes(fullPath, cancellationToken)
                AssertExpectedHash(existingBytes, expectedFileHash)
                Dim json As System.String
                If saveToLibrary Then
                    json = PrepareLibrarySave(existingBytes, template, replaceName)
                Else
                    json = SerializeTemplate(template)
                End If
                Dim bytes As System.Byte() = StrictUtf8.GetBytes(json)
                Dim writtenHash As System.String = HashBytes(bytes)
                Dim temporary As System.String = Global.SharedLibrary.Agents.FilePathBudget.UniqueCopyPath(directory, ".ri-columns-" & System.Guid.NewGuid().ToString("N") & ".tmp")
                Dim committed As System.Boolean = False
                Try
                    Using stream As New System.IO.FileStream(temporary, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None)
                        Dim offset As System.Int32 = 0
                        While offset < bytes.Length
                            cancellationToken.ThrowIfCancellationRequested()
                            Dim count As System.Int32 = System.Math.Min(65536, bytes.Length - offset)
                            stream.Write(bytes, offset, count)
                            offset += count
                        End While
                        stream.Flush(True)
                    End Using
                    Dim verified As ReviewTemplateFileSnapshot = ReadFile(temporary, cancellationToken)
                    If Not System.String.Equals(verified.FileHash, writtenHash, System.StringComparison.Ordinal) OrElse
                       Not System.String.Equals(SerializeFile(verified.Templates, verified.IsCatalog), json, System.StringComparison.Ordinal) Then
                        Throw New System.IO.InvalidDataException("The column settings failed their complete roundtrip check. The destination file has not changed.")
                    End If
                    AssertExpectedHash(ReadOptionalBytes(fullPath, cancellationToken), expectedFileHash)
                    cancellationToken.ThrowIfCancellationRequested()
                    If validateDestination IsNot Nothing Then validateDestination(fullPath)
                    cancellationToken.ThrowIfCancellationRequested()
                    If existingBytes IsNot Nothing Then
                        ' Keep the old file if this filesystem cannot replace it atomically.
                        ' File.Replace retains destination metadata; never delete and recreate.
                        System.IO.File.Replace(temporary, fullPath, Nothing)
                    Else
                        ' Move refuses a destination another process created meanwhile.
                        System.IO.File.Move(temporary, fullPath)
                    End If
                    committed = True
                    ' Publication has completed. A late cancellation must not report that
                    ' the successful write was cancelled and encourage a duplicate retry.
                    Dim published As ReviewTemplateFileSnapshot = ReadFile(fullPath, System.Threading.CancellationToken.None)
                    If Not System.String.Equals(published.FileHash, writtenHash, System.StringComparison.Ordinal) Then
                        Throw New System.IO.IOException("The destination changed after the column settings save.")
                    End If
                    Return published
                Catch failure As System.Exception When committed
                    Throw New System.IO.IOException("The column settings save completed, but verification could not confirm the current file. Refresh or reopen the file before saving again.", failure)
                Finally
                    Try
                        If System.IO.File.Exists(temporary) Then System.IO.File.Delete(temporary)
                    Catch cleanupFailure As System.Exception
                        System.Diagnostics.Trace.TraceWarning("[TabularReview:template_cleanup_failed] " & cleanupFailure.Message)
                    End Try
                End Try
            End Using
        End Function

        Private Shared Function OpenWriteLock(lockPath As System.String) As System.IO.FileStream
            Try
                Return New System.IO.FileStream(lockPath, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
            Catch failure As System.Exception When TypeOf failure Is System.IO.IOException OrElse
                TypeOf failure Is System.UnauthorizedAccessException OrElse TypeOf failure Is System.Security.SecurityException
                System.Diagnostics.Trace.TraceWarning("[TabularReview:template_lock_failed] " & failure.GetType().FullName & " | " & failure.Message)
                Throw New System.IO.IOException("Cannot start the template save. Check that the folder is available and writable, then retry after any other save finishes.", failure)
            End Try
        End Function

        Private Shared Function PrepareLibrarySave(existingBytes As System.Byte(), template As ReviewColumnTemplate,
                                                   replaceName As System.String) As System.String
            Dim snapshot As ReviewTemplateFileSnapshot = If(existingBytes Is Nothing, New ReviewTemplateFileSnapshot With {.IsCatalog = True}, ParseFileJson(DecodeJson(existingBytes)))
            If replaceName Is Nothing Then
                For Each existing As ReviewColumnTemplate In snapshot.Templates
                    If System.String.Equals(existing.Name, template.Name, System.StringComparison.OrdinalIgnoreCase) Then
                        Throw New System.IO.IOException("A template with this name already exists in this file. Choose it for replacement or enter a different name.")
                    End If
                Next
                snapshot.Templates.Add(template)
            Else
                Dim selected As System.Int32 = snapshot.Templates.FindIndex(Function(existing As ReviewColumnTemplate) System.String.Equals(existing.Name, replaceName, System.StringComparison.Ordinal))
                If selected < 0 Then Throw New System.IO.IOException("The template to replace no longer exists under that name. Refresh the library list and choose it again.")
                snapshot.Templates(selected) = template
            End If
            ' Adding a second entry to a standalone file creates a catalog and retains
            ' the entire first definition. Replacing its sole entry keeps its format.
            Return SerializeFile(snapshot.Templates, snapshot.IsCatalog OrElse snapshot.Templates.Count > 1)
        End Function

        Private Shared Sub LoadScope(configuredPath As System.String, scope As System.String,
                                     result As ReviewTemplateLibraryResult,
                                     labels As System.Collections.Generic.HashSet(Of System.String),
                                     cancellationToken As System.Threading.CancellationToken)
            cancellationToken.ThrowIfCancellationRequested()
            If System.String.IsNullOrWhiteSpace(configuredPath) Then Return
            Dim fullPath As System.String = ""
            Dim files As New System.Collections.Generic.List(Of System.String)()
            Try
                fullPath = ResolvePath(configuredPath)
                Dim attributes As System.Nullable(Of System.IO.FileAttributes) = GetAttributesOrMissing(fullPath)
                If Not attributes.HasValue Then
                    result.Notices.Add(New ReviewTemplateLibraryNotice With {
                        .Scope = scope, .FileName = DisplayFileName(fullPath),
                        .Message = "No library exists at this configured path. Save a column template to create it."
                    })
                    Return
                End If
                If (attributes.Value And System.IO.FileAttributes.Directory) <> 0 Then
                    For Each file As System.String In System.IO.Directory.EnumerateFiles(fullPath, "*.json", System.IO.SearchOption.TopDirectoryOnly)
                        cancellationToken.ThrowIfCancellationRequested()
                        files.Add(file)
                    Next
                    files.Sort(New System.Comparison(Of System.String)(AddressOf CompareFilePaths))
                Else
                    files.Add(fullPath)
                End If
            Catch failure As System.Exception When IsLibraryReadFailure(failure)
                AddReadNotice(result, scope, fullPath, failure)
                Return
            End Try
            For Each file As System.String In files
                cancellationToken.ThrowIfCancellationRequested()
                Try
                    Dim snapshot As ReviewTemplateFileSnapshot = ReadFile(file, cancellationToken)
                    For index As System.Int32 = 0 To snapshot.Templates.Count - 1
                        cancellationToken.ThrowIfCancellationRequested()
                        Dim template As ReviewColumnTemplate = snapshot.Templates(index)
                        Dim baseLabel As System.String = template.Name & " (" & scope & " - " & DisplayFileName(file) & ")"
                        Dim label As System.String = baseLabel
                        Dim suffix As System.Int32 = 2
                        While Not labels.Add(label)
                            label = baseLabel & " [" & suffix.ToString(System.Globalization.CultureInfo.InvariantCulture) & "]"
                            suffix += 1
                        End While
                        result.Entries.Add(New ReviewTemplateLibraryEntry With {
                            .DisplayName = label, .Scope = scope, .FilePath = snapshot.FilePath,
                            .FileHash = snapshot.FileHash, .Index = index, .Template = template
                        })
                    Next
                Catch failure As System.Exception When IsLibraryReadFailure(failure)
                    AddReadNotice(result, scope, file, failure)
                End Try
            Next
        End Sub

        Private Shared Function IsLibraryReadFailure(failure As System.Exception) As System.Boolean
            Return TypeOf failure Is System.IO.InvalidDataException OrElse TypeOf failure Is System.IO.IOException OrElse
                TypeOf failure Is System.UnauthorizedAccessException OrElse
                TypeOf failure Is System.Security.SecurityException OrElse TypeOf failure Is System.ArgumentException OrElse
                TypeOf failure Is System.NotSupportedException
        End Function

        Private Shared Sub AddReadNotice(result As ReviewTemplateLibraryResult, scope As System.String,
                                         filePath As System.String, failure As System.Exception)
            Dim message As System.String = "Could not read this library. Check its availability and your read access, then refresh the library list."
            If TypeOf failure Is System.IO.InvalidDataException Then message = failure.Message
            If TypeOf failure Is System.IO.PathTooLongException Then message = Global.SharedLibrary.Agents.FilePathBudget.PathTooLongMessage
            result.Notices.Add(New ReviewTemplateLibraryNotice With {.Scope = scope, .FileName = DisplayFileName(filePath), .Message = message})
            System.Diagnostics.Trace.TraceWarning("[TabularReview:template_library_read_failed] scope=" & scope & " | " & failure.GetType().FullName & " | " & failure.Message)
        End Sub

        Private Shared Function CompareFilePaths(left As System.String, right As System.String) As System.Int32
            Dim result As System.Int32 = System.StringComparer.OrdinalIgnoreCase.Compare(left, right)
            Return If(result = 0, System.StringComparer.Ordinal.Compare(left, right), result)
        End Function

        Private Shared Function DisplayFileName(filePath As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(filePath) Then Return "Configured library"
            Return System.IO.Path.GetFileName(filePath.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar))
        End Function

        Private Shared Function ResolvePath(value As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(value) Then Throw New System.IO.InvalidDataException("Configure a template library path or choose a JSON file.")
            Dim expanded As System.String = Global.SharedLibrary.SharedLibrary.SharedMethods.ExpandEnvironmentVariables(value)
            If System.String.IsNullOrWhiteSpace(expanded) OrElse System.Text.RegularExpressions.Regex.IsMatch(expanded, "%[^%]+%") Then
                Throw New System.IO.InvalidDataException("The template path contains an unresolved placeholder. Correct the path in Settings or choose an accessible JSON file.")
            End If
            Return Global.SharedLibrary.Agents.FilePathBudget.GetFullPath(expanded)
        End Function

        Private Shared Function GetAttributesOrMissing(filePath As System.String) As System.Nullable(Of System.IO.FileAttributes)
            Try
                Return System.IO.File.GetAttributes(filePath)
            Catch missingFile As System.IO.FileNotFoundException
                Return Nothing
            Catch missingDirectory As System.IO.DirectoryNotFoundException
                Return Nothing
            End Try
        End Function

        Private Shared Function ReadOptionalBytes(filePath As System.String, cancellationToken As System.Threading.CancellationToken) As System.Byte()
            cancellationToken.ThrowIfCancellationRequested()
            Dim attributes As System.Nullable(Of System.IO.FileAttributes) = GetAttributesOrMissing(filePath)
            If Not attributes.HasValue Then Return Nothing
            If (attributes.Value And System.IO.FileAttributes.Directory) <> 0 Then Throw New System.IO.InvalidDataException("Choose a JSON filename, not a folder, for this operation.")
            Return ReadBytes(filePath, cancellationToken)
        End Function

        Private Shared Function ReadBytes(filePath As System.String, cancellationToken As System.Threading.CancellationToken) As System.Byte()
            cancellationToken.ThrowIfCancellationRequested()
            Using stream As New System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                If stream.Length > MaximumJsonBytes Then Throw New System.IO.InvalidDataException("This column settings file exceeds 32 MiB. Split the library into smaller JSON files; no definitions have been shortened.")
                Using content As New System.IO.MemoryStream()
                    Dim buffer(65535) As System.Byte
                    Do
                        cancellationToken.ThrowIfCancellationRequested()
                        Dim count As System.Int32 = stream.Read(buffer, 0, buffer.Length)
                        If count = 0 Then Exit Do
                        If content.Length + count > MaximumJsonBytes Then Throw New System.IO.InvalidDataException("This column settings file exceeds 32 MiB. Split the library into smaller JSON files; no definitions have been shortened.")
                        content.Write(buffer, 0, count)
                    Loop
                    cancellationToken.ThrowIfCancellationRequested()
                    Return content.ToArray()
                End Using
            End Using
        End Function

        Private Shared Function DecodeJson(bytes As System.Byte()) As System.String
            Dim start As System.Int32 = If(bytes.Length >= 3 AndAlso bytes(0) = &HEF AndAlso bytes(1) = &HBB AndAlso bytes(2) = &HBF, 3, 0)
            Try
                Return StrictUtf8.GetString(bytes, start, bytes.Length - start)
            Catch invalidEncoding As System.Text.DecoderFallbackException
                Throw New System.IO.InvalidDataException("Save the column settings JSON as valid UTF-8 and import it again. The file contains an invalid text encoding.", invalidEncoding)
            End Try
        End Function

        Private Shared Function HashBytes(bytes As System.Byte()) As System.String
            Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Return System.BitConverter.ToString(algorithm.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant()
            End Using
        End Function

        Private Shared Sub RequireExpectedHash(value As System.String)
            If value Is Nothing OrElse (value.Length > 0 AndAlso Not System.Text.RegularExpressions.Regex.IsMatch(value, "\A[0-9a-fA-F]{64}\z")) Then
                Throw New System.IO.InvalidDataException("Refresh the destination file before saving column settings.")
            End If
        End Sub

        Private Shared Sub AssertExpectedHash(bytes As System.Byte(), expected As System.String)
            Dim actual As System.String = If(bytes Is Nothing, "", HashBytes(bytes))
            If Not System.String.Equals(actual, expected, System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.IO.IOException("The destination file changed after you opened it. Refresh the library or choose another filename to keep both versions.")
            End If
        End Sub

        Private Shared Sub RequireText(value As System.String, field As System.String, requireContent As System.Boolean)
            If value Is Nothing OrElse (requireContent AndAlso System.String.IsNullOrWhiteSpace(value)) Then Throw New System.IO.InvalidDataException("Provide a valid " & field & ".")
            Try
                StrictUtf8.GetByteCount(value)
            Catch invalidText As System.Text.EncoderFallbackException
                Throw New System.IO.InvalidDataException("The " & field & " contains an incomplete Unicode character. Correct it before saving.", invalidText)
            End Try
        End Sub

        Private Shared Sub RequireWidth(value As System.Nullable(Of System.Int32), minimum As System.Int32,
                                        maximum As System.Int32, field As System.String)
            If value.HasValue AndAlso (value.Value < minimum OrElse value.Value > maximum) Then
                Throw New System.IO.InvalidDataException("Set the " & field & " width between " & minimum.ToString(System.Globalization.CultureInfo.InvariantCulture) & " and " & maximum.ToString(System.Globalization.CultureInfo.InvariantCulture) & " pixels.")
            End If
        End Sub

        Private Shared Function ParseFileJson(json As System.String) As ReviewTemplateFileSnapshot
            RequireText(json, "column settings JSON", True)
            If StrictUtf8.GetByteCount(json) > MaximumJsonBytes Then Throw New System.IO.InvalidDataException("This column settings JSON exceeds 32 MiB. Split the library into smaller files; no definitions have been shortened.")
            If json(0) = System.Convert.ToChar(&HFEFF) Then json = json.Substring(1)
            ValidateEscapedUnicode(json)
            Dim root As Newtonsoft.Json.Linq.JObject
            Try
                Using textReader As New System.IO.StringReader(json)
                    Using reader As New StrictTemplateJsonReader(textReader)
                        reader.MaxDepth = MaximumJsonDepth
                        reader.DateParseHandling = Newtonsoft.Json.DateParseHandling.None
                        reader.FloatParseHandling = Newtonsoft.Json.FloatParseHandling.Decimal
                        root = TryCast(Newtonsoft.Json.Linq.JToken.Load(reader, New Newtonsoft.Json.Linq.JsonLoadSettings With {
                            .DuplicatePropertyNameHandling = Newtonsoft.Json.Linq.DuplicatePropertyNameHandling.Error
                        }), Newtonsoft.Json.Linq.JObject)
                        If reader.Read() Then Throw New System.IO.InvalidDataException("Keep exactly one JSON object in the column settings file.")
                    End Using
                End Using
            Catch invalidJson As Newtonsoft.Json.JsonException
                Throw New System.IO.InvalidDataException("The column settings JSON is invalid. Check its structure and import it again.", invalidJson)
            End Try
            If root Is Nothing Then Throw New System.IO.InvalidDataException("Column settings must contain a JSON object with a format and version.")
            Dim header As System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken) = ReadObjectFields(root,
                New System.String() {"format", "version", "name", "description", "columns", "documentColumnWidth", "flagColors", "reportDefinitions", "templates"}, "column settings")
            Dim format As System.String = StringValue(header, "format", "").Trim().ToLowerInvariant()
            If IntegerValue(header, "version", False).GetValueOrDefault() <> CurrentVersion Then Throw New System.IO.InvalidDataException("This column settings version is unsupported. Use a compatible Red Ink version.")
            Dim snapshot As New ReviewTemplateFileSnapshot()
            If format = ColumnsFormat Then
                Dim fields As System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken) = ReadObjectFields(root,
                    New System.String() {"format", "version", "name", "description", "columns", "documentColumnWidth", "flagColors", "reportDefinitions"}, "column template")
                snapshot.Templates.Add(ParseTemplate(fields))
            ElseIf format = LibraryFormat Then
                Dim fields As System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken) = ReadObjectFields(root,
                    New System.String() {"format", "version", "templates"}, "template library")
                Dim templates As Newtonsoft.Json.Linq.JArray = TryCast(RequiredValue(fields, "templates"), Newtonsoft.Json.Linq.JArray)
                If templates Is Nothing Then Throw New System.IO.InvalidDataException("A template library needs a templates array.")
                snapshot.IsCatalog = True
                For Each value As Newtonsoft.Json.Linq.JToken In templates
                    Dim item As Newtonsoft.Json.Linq.JObject = TryCast(value, Newtonsoft.Json.Linq.JObject)
                    If item Is Nothing Then Throw New System.IO.InvalidDataException("Every library entry must be a column template object.")
                    snapshot.Templates.Add(ParseTemplate(ReadObjectFields(item,
                        New System.String() {"name", "description", "columns", "documentColumnWidth", "flagColors", "reportDefinitions"}, "library template")))
                Next
                RequireUniqueNames(snapshot.Templates)
            Else
                Throw New System.IO.InvalidDataException("Choose a Red Ink column settings JSON file or template library. This file has an unsupported format.")
            End If
            Return snapshot
        End Function

        Private Shared Function ParseTemplate(fields As System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken)) As ReviewColumnTemplate
            Dim template As New ReviewColumnTemplate With {
                .Name = StringValue(fields, "name", ""), .Description = StringValue(fields, "description", ""),
                .DocumentColumnWidth = IntegerValue(fields, "documentColumnWidth", True)
            }
            Dim columns As Newtonsoft.Json.Linq.JArray = TryCast(RequiredValue(fields, "columns"), Newtonsoft.Json.Linq.JArray)
            If columns Is Nothing Then Throw New System.IO.InvalidDataException("A column template needs an ordered columns array.")
            For Each value As Newtonsoft.Json.Linq.JToken In columns
                Dim item As Newtonsoft.Json.Linq.JObject = TryCast(value, Newtonsoft.Json.Linq.JObject)
                If item Is Nothing Then Throw New System.IO.InvalidDataException("Each column definition must be a JSON object.")
                Dim columnFields As System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken) = ReadObjectFields(item,
                    New System.String() {"title", "instruction", "kind", "options", "width"}, "column definition")
                Dim column As New ReviewTemplateColumn With {
                    .Title = StringValue(columnFields, "title", ""), .Instruction = StringValue(columnFields, "instruction", ""),
                    .Kind = StringValue(columnFields, "kind", "analysis"), .Width = IntegerValue(columnFields, "width", True)
                }
                Dim optionsToken As Newtonsoft.Json.Linq.JToken = Nothing
                If columnFields.TryGetValue("options", optionsToken) Then
                    Dim options As Newtonsoft.Json.Linq.JArray = TryCast(optionsToken, Newtonsoft.Json.Linq.JArray)
                    If options Is Nothing Then Throw New System.IO.InvalidDataException("Column options must be an array of text values.")
                    For Each optionToken As Newtonsoft.Json.Linq.JToken In options
                        If optionToken.Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("Every column option must contain text.")
                        column.Options.Add(DirectCast(DirectCast(optionToken, Newtonsoft.Json.Linq.JValue).Value, System.String))
                    Next
                End If
                template.Columns.Add(column)
            Next
            Dim colorsToken As Newtonsoft.Json.Linq.JToken = Nothing
            If fields.TryGetValue("flagColors", colorsToken) Then
                Dim colors As Newtonsoft.Json.Linq.JObject = TryCast(colorsToken, Newtonsoft.Json.Linq.JObject)
                If colors Is Nothing Then Throw New System.IO.InvalidDataException("Flag colors must be an object mapping exact flag labels to hexadecimal colors.")
                For Each color As Newtonsoft.Json.Linq.JProperty In colors.Properties()
                    If color.Value.Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("Each flag color must contain a hexadecimal text value.")
                    template.FlagColors.Add(color.Name, DirectCast(DirectCast(color.Value, Newtonsoft.Json.Linq.JValue).Value, System.String))
                Next
            End If
            Dim reportsToken As Newtonsoft.Json.Linq.JToken = Nothing
            If fields.TryGetValue("reportDefinitions", reportsToken) Then
                If reportsToken.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then Throw New System.IO.InvalidDataException("Report definitions must be an array.")
                For Each reportToken As Newtonsoft.Json.Linq.JToken In DirectCast(reportsToken, Newtonsoft.Json.Linq.JArray)
                    Dim reportObject As Newtonsoft.Json.Linq.JObject = TryCast(reportToken, Newtonsoft.Json.Linq.JObject)
                    If reportObject Is Nothing Then Throw New System.IO.InvalidDataException("Each report must contain an object.")
                    Dim reportFields As System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken) = ReadObjectFields(reportObject,
                        New System.String() {"id", "name", "title", "documentScope", "kind", "orientation", "design", "paperSize", "fields"}, "report definition")
                    Dim report As New ReviewReportDefinition With {.Id = StringValue(reportFields, "id", ""), .Name = StringValue(reportFields, "name", ""),
                        .Title = StringValue(reportFields, "title", ""), .DocumentScope = StringValue(reportFields, "documentScope", "filtered"), .Kind = StringValue(reportFields, "kind", ""), .PaperSize = StringValue(reportFields, "paperSize", "A4"), .Design = StringValue(reportFields, "design", "redink"), .Orientation = StringValue(reportFields, "orientation", "")}
                    Dim keys As Newtonsoft.Json.Linq.JArray = TryCast(RequiredValue(reportFields, "fields"), Newtonsoft.Json.Linq.JArray)
                    If keys Is Nothing Then Throw New System.IO.InvalidDataException("Report fields must be an ordered array.")
                    For Each key As Newtonsoft.Json.Linq.JToken In keys
                        If key.Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("Report field keys must contain text.")
                        report.Fields.Add(key.ToString())
                    Next
                    template.ReportDefinitions.Add(report)
                Next
            End If
            Return ValidateAndClone(template)
        End Function

        Private Shared Function ReadObjectFields(value As Newtonsoft.Json.Linq.JObject, allowed As System.String(),
                                       context As System.String) As System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken)
            Dim names As New System.Collections.Generic.HashSet(Of System.String)(allowed, System.StringComparer.OrdinalIgnoreCase)
            Dim result As New System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken)(System.StringComparer.OrdinalIgnoreCase)
            For Each propertyValue As Newtonsoft.Json.Linq.JProperty In value.Properties()
                If Not names.Contains(propertyValue.Name) Then Throw New System.IO.InvalidDataException("The " & context & " contains an unsupported field: " & propertyValue.Name & ". Use a compatible format; no unknown fields will be discarded.")
                If result.ContainsKey(propertyValue.Name) Then Throw New System.IO.InvalidDataException("The " & context & " repeats a field with different capitalization. Keep each field only once.")
                result.Add(propertyValue.Name, propertyValue.Value)
            Next
            Return result
        End Function

        Private Shared Function RequiredValue(fields As System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken),
                                               name As System.String) As Newtonsoft.Json.Linq.JToken
            Dim value As Newtonsoft.Json.Linq.JToken = Nothing
            If Not fields.TryGetValue(name, value) Then Throw New System.IO.InvalidDataException("The column settings are missing the " & name & " field.")
            Return value
        End Function

        Private Shared Function StringValue(fields As System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken),
                                             name As System.String, defaultValue As System.String) As System.String
            Dim value As Newtonsoft.Json.Linq.JToken = Nothing
            If Not fields.TryGetValue(name, value) Then Return defaultValue
            If value.Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("The " & name & " field must contain text.")
            Return DirectCast(DirectCast(value, Newtonsoft.Json.Linq.JValue).Value, System.String)
        End Function

        Private Shared Function IntegerValue(fields As System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken),
                                              name As System.String, optionalValue As System.Boolean) As System.Nullable(Of System.Int32)
            Dim value As Newtonsoft.Json.Linq.JToken = Nothing
            If Not fields.TryGetValue(name, value) Then
                If optionalValue Then Return Nothing
                Throw New System.IO.InvalidDataException("The column settings are missing the " & name & " field.")
            End If
            If optionalValue AndAlso value.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return Nothing
            Dim result As System.Int32
            If value.Type <> Newtonsoft.Json.Linq.JTokenType.Integer OrElse
               Not System.Int32.TryParse(value.ToString(Newtonsoft.Json.Formatting.None), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, result) Then
                Throw New System.IO.InvalidDataException("The " & name & " field must contain a whole number.")
            End If
            Return result
        End Function

        Private Shared Sub RequireUniqueNames(templates As System.Collections.Generic.List(Of ReviewColumnTemplate))
            Dim names As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each template As ReviewColumnTemplate In templates
                If Not names.Add(template.Name) Then Throw New System.IO.InvalidDataException("A template library repeats a name. Give each template in that file a different name.")
            Next
        End Sub

        Private Shared Function SerializeFile(templates As System.Collections.Generic.List(Of ReviewColumnTemplate),
                                               isCatalog As System.Boolean) As System.String
            If templates Is Nothing OrElse (Not isCatalog AndAlso templates.Count <> 1) Then Throw New System.IO.InvalidDataException("Choose one complete column template for export.")
            Dim copies As New System.Collections.Generic.List(Of ReviewColumnTemplate)()
            For Each template As ReviewColumnTemplate In templates
                copies.Add(ValidateAndClone(template))
            Next
            RequireUniqueNames(copies)
            Dim root As New Newtonsoft.Json.Linq.JObject()
            root.Add("format", New Newtonsoft.Json.Linq.JValue(If(isCatalog, LibraryFormat, ColumnsFormat)))
            root.Add("version", New Newtonsoft.Json.Linq.JValue(CurrentVersion))
            If isCatalog Then
                Dim entries As New Newtonsoft.Json.Linq.JArray()
                For Each template As ReviewColumnTemplate In copies
                    entries.Add(TemplateObject(template))
                Next
                root.Add("templates", entries)
            Else
                For Each propertyValue As Newtonsoft.Json.Linq.JProperty In TemplateObject(copies(0)).Properties()
                    root.Add(propertyValue.Name, propertyValue.Value.DeepClone())
                Next
            End If
            Dim json As System.String = root.ToString(Newtonsoft.Json.Formatting.Indented)
            If StrictUtf8.GetByteCount(json) > MaximumJsonBytes Then Throw New System.IO.InvalidDataException("This column settings JSON exceeds 32 MiB. Split the library into smaller files; no definitions have been shortened.")
            Return json
        End Function

        Private Shared Function TemplateObject(template As ReviewColumnTemplate) As Newtonsoft.Json.Linq.JObject
            Dim root As New Newtonsoft.Json.Linq.JObject()
            root.Add("name", New Newtonsoft.Json.Linq.JValue(template.Name))
            root.Add("description", New Newtonsoft.Json.Linq.JValue(template.Description))
            Dim columns As New Newtonsoft.Json.Linq.JArray()
            For Each column As ReviewTemplateColumn In template.Columns
                Dim options As New Newtonsoft.Json.Linq.JArray()
                For Each optionText As System.String In column.Options
                    options.Add(New Newtonsoft.Json.Linq.JValue(optionText))
                Next
                Dim value As New Newtonsoft.Json.Linq.JObject()
                value.Add("title", New Newtonsoft.Json.Linq.JValue(column.Title))
                value.Add("instruction", New Newtonsoft.Json.Linq.JValue(column.Instruction))
                value.Add("kind", New Newtonsoft.Json.Linq.JValue(column.Kind))
                value.Add("options", options)
                If column.Width.HasValue Then value.Add("width", New Newtonsoft.Json.Linq.JValue(column.Width.Value))
                columns.Add(value)
            Next
            root.Add("columns", columns)
            If template.DocumentColumnWidth.HasValue Then root.Add("documentColumnWidth", New Newtonsoft.Json.Linq.JValue(template.DocumentColumnWidth.Value))
            Dim colors As New Newtonsoft.Json.Linq.JObject()
            Dim labels As New System.Collections.Generic.List(Of System.String)(template.FlagColors.Keys)
            labels.Sort(System.StringComparer.Ordinal)
            For Each label As System.String In labels
                colors.Add(label, New Newtonsoft.Json.Linq.JValue(template.FlagColors(label)))
            Next
            root.Add("flagColors", colors)
            root.Add("reportDefinitions", Newtonsoft.Json.Linq.JArray.Parse(ReviewContract.Serialize(template.ReportDefinitions)))
            Return root
        End Function

        Private NotInheritable Class StrictTemplateJsonReader
            Inherits Newtonsoft.Json.JsonTextReader

            Public Sub New(reader As System.IO.TextReader)
                MyBase.New(reader)
            End Sub

            Public Overrides Function Read() As System.Boolean
                Dim result As System.Boolean = MyBase.Read()
                If result AndAlso TokenType = Newtonsoft.Json.JsonToken.Comment Then Throw New System.IO.InvalidDataException("Use JSON fields instead of comments in column settings files.")
                Return result
            End Function
        End Class

        ' JsonTextReader can replace an unmatched escaped surrogate with U+FFFD.
        ' Reject that input before parsing so an import never silently changes text.
        Private Shared Sub ValidateEscapedUnicode(json As System.String)
            Dim quote As System.Char = System.Convert.ToChar(0)
            Dim index As System.Int32 = 0
            While index < json.Length
                Dim current As System.Char = json(index)
                If quote = System.Convert.ToChar(0) Then
                    If current = """"c OrElse current = "'"c Then quote = current
                ElseIf current = quote Then
                    quote = System.Convert.ToChar(0)
                ElseIf current = "\"c Then
                    If index + 1 < json.Length AndAlso json(index + 1) = "u"c Then
                        Dim unit As System.Int32 = EscapedCodeUnit(json, index)
                        If unit >= &HD800 AndAlso unit <= &HDBFF Then
                            Dim second As System.Int32 = EscapedCodeUnit(json, index + 6)
                            If second < &HDC00 OrElse second > &HDFFF Then Throw New System.IO.InvalidDataException("The column settings JSON contains an incomplete Unicode escape pair.")
                            index += 11
                        Else
                            If unit >= &HDC00 AndAlso unit <= &HDFFF Then Throw New System.IO.InvalidDataException("The column settings JSON contains an unmatched Unicode escape.")
                            index += 5
                        End If
                    Else
                        index += 1
                    End If
                End If
                index += 1
            End While
        End Sub

        Private Shared Function EscapedCodeUnit(json As System.String, start As System.Int32) As System.Int32
            Dim result As System.Int32
            If start < 0 OrElse start + 5 >= json.Length OrElse json(start) <> "\"c OrElse json(start + 1) <> "u"c OrElse
               Not System.Int32.TryParse(json.Substring(start + 2, 4), System.Globalization.NumberStyles.AllowHexSpecifier, System.Globalization.CultureInfo.InvariantCulture, result) Then
                Throw New System.IO.InvalidDataException("The column settings JSON contains an invalid Unicode escape.")
            End If
            Return result
        End Function
    End Class
End Namespace
