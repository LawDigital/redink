' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
Option Strict On
Option Explicit On
Option Infer On

Namespace SharedLibrary
    ''' <summary>Bounded diagnostic data, not document content. Source identity is resolved afresh before display.</summary>
    Public NotInheritable Class SemanticArchiveSourceDiagnostic
        Public Property Code As System.String = ""
        Public Property DocumentId As System.String = ""
        Public Property Detail As System.String = ""
        Public Property Occurrences As System.Int32 = 1
    End Class

    Public NotInheritable Partial Class SemanticArchiveBuilder
        Friend Shared Function IsTechnicalDiagnostic(message As System.String) As System.Boolean
            Return message IsNot Nothing AndAlso (message.StartsWith("routing_metadata_reduced:", System.StringComparison.Ordinal) OrElse
                message.StartsWith("local_split:", System.StringComparison.Ordinal) OrElse
                message.StartsWith("cooperative_contributed:", System.StringComparison.Ordinal) OrElse
                message.StartsWith("cooperative_reuse:", System.StringComparison.Ordinal) OrElse
                message.StartsWith("selected_source_identity_changed:", System.StringComparison.Ordinal) OrElse
                message.StartsWith("retired_records_pruned:", System.StringComparison.Ordinal) OrElse
                (message.StartsWith("permission_reconciliation:", System.StringComparison.Ordinal) AndAlso
                 message.IndexOf("process-local rights", System.StringComparison.OrdinalIgnoreCase) >= 0))
        End Function

        Friend Shared Function ParseSourceDiagnostic(message As System.String) As SemanticArchiveSourceDiagnostic
            If System.String.IsNullOrWhiteSpace(message) Then Return Nothing
            Dim delimiter As System.Int32 = message.IndexOf(":"c)
            If delimiter < 1 Then Return Nothing
            Dim code As System.String = message.Substring(0, delimiter)
            If IsTechnicalDiagnostic(message) Then Return New SemanticArchiveSourceDiagnostic With {.Code = code, .Detail = DiagnosticLine(message.Substring(delimiter + 1))}
            Select Case code
                Case "coverage_excluded", "failed", "pending_host", "shared_claim_deferred", "cooperative_local_only",
                     "cooperative_candidate_unavailable", "cooperative_contribution_pending", "cooperative_contributed", "cooperative_reuse",
                     "permission_reconciliation", "selected_source_identity_changed", "extraction_configuration_changed", "semantic_configuration_changed"
                Case Else : Return Nothing
            End Select
            Dim remainder As System.String = message.Substring(delimiter + 1).TrimStart()
            Dim separator As System.Int32 = remainder.IndexOf(";"c)
            If separator < 1 Then Return Nothing
            Dim documentId As System.String = remainder.Substring(0, separator).Trim()
            If Not IsDiagnosticDocumentId(documentId) Then Return Nothing
            Return New SemanticArchiveSourceDiagnostic With {.Code = code, .DocumentId = documentId, .Detail = DiagnosticLine(remainder.Substring(separator + 1))}
        End Function

        Private Shared Function IsDiagnosticDocumentId(value As System.String) As System.Boolean
            If value Is Nothing OrElse value.Length <> 68 OrElse Not value.StartsWith("doc_", System.StringComparison.Ordinal) Then Return False
            For Each character As System.Char In value.Substring(4)
                If Not ((character >= "0"c AndAlso character <= "9"c) OrElse (character >= "a"c AndAlso character <= "f"c)) Then Return False
            Next
            Return True
        End Function

        Private Shared Function DiagnosticLine(value As System.String) As System.String
            If value Is Nothing Then Return ""
            Dim maximum As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_DIAGNOSTIC_EVENT_CHARACTERS
            Dim text As New System.Text.StringBuilder()
            For Each character As System.Char In value
                If text.Length >= maximum Then Exit For
                text.Append(If(System.Char.IsControl(character), " "c, character))
            Next
            If text.Length > 0 AndAlso System.Char.IsHighSurrogate(text(text.Length - 1)) Then text.Length -= 1
            Dim normalized As System.String = text.ToString().Trim()
            If value.Length > maximum Then normalized &= " [detail truncated]"
            Return normalized
        End Function

        Private Shared Function RawDiagnostic(entry As SemanticArchiveSourceDiagnostic) As System.String
            Return entry.Code & ": " & If(entry.DocumentId.Length = 0, "", entry.DocumentId & "; ") & entry.Detail
        End Function

        Friend Shared Sub CaptureDiagnostic(record As SemanticArchiveOperationDiagnostic, entry As SemanticArchiveSourceDiagnostic)
            If record Is Nothing Then Throw New System.ArgumentNullException(NameOf(record))
            If entry Is Nothing OrElse entry.Code Is Nothing OrElse entry.DocumentId Is Nothing OrElse entry.Detail Is Nothing Then Return
            Dim checked As SemanticArchiveSourceDiagnostic = ParseSourceDiagnostic(RawDiagnostic(entry))
            If checked Is Nothing Then Return
            For Each old As SemanticArchiveSourceDiagnostic In record.SourceDiagnostics
                If old.Code = checked.Code AndAlso old.DocumentId = checked.DocumentId AndAlso old.Detail = checked.Detail Then
                    If IsTechnicalDiagnostic(RawDiagnostic(checked)) AndAlso old.Occurrences < System.Int32.MaxValue Then old.Occurrences += 1
                    Return
                End If
            Next
            If record.SourceDiagnostics.Count >= SharedMethods.DEFAULT_SEMANTICARCHIVE_DIAGNOSTIC_MAXIMUM_EVENTS Then
                Dim technicalIndex As System.Int32 = record.SourceDiagnostics.FindLastIndex(Function(old) IsTechnicalDiagnostic(RawDiagnostic(old)))
                If record.OmittedDiagnostics < System.Int32.MaxValue Then record.OmittedDiagnostics += 1
                If IsTechnicalDiagnostic(RawDiagnostic(checked)) OrElse technicalIndex < 0 Then Return
                record.SourceDiagnostics.RemoveAt(technicalIndex)
            End If
            record.SourceDiagnostics.Add(checked)
        End Sub

        Private Shared Sub ReportSourceDiagnostic(progress As System.IProgress(Of SemanticArchiveBuildProgress),
                         result As SemanticArchiveBuildResult, archive As SemanticArchiveDefinition, item As SemanticArchiveWorkItem,
                         code As System.String, detail As System.String)
            Dim entry As New SemanticArchiveSourceDiagnostic With {.Code = code, .DocumentId = item.DocumentId, .Detail = DiagnosticLine(detail)}
            result.Diagnostics.Add(RawDiagnostic(entry))
            If progress IsNot Nothing Then progress.Report(New SemanticArchiveBuildProgress With {
                .ArchiveId = result.ArchiveId, .Stage = code, .Message = FormatSourceEntry(entry, archive, item, SemanticArchiveAccessContext.CreateForCurrentUser()),
                .SourceDiagnostic = entry, .CompletedFiles = result.ProcessedFiles, .PendingFiles = result.PendingFiles, .DeferredFiles = result.DeferredFiles})
        End Sub

        Private Shared Function FormatUserSourceEntry(entry As SemanticArchiveSourceDiagnostic, archive As SemanticArchiveDefinition,
                         item As SemanticArchiveWorkItem, access As SemanticArchiveAccessContext) As System.String
            If item Is Nothing OrElse Not CanDiscloseQueuedSource(item, archive, access) Then
                Return "Attention required for " & entry.DocumentId & ". Current source access is unavailable; details are hidden."
            End If
            Dim name As System.String = DiagnosticLine(System.IO.Path.GetFileName(item.SourcePath))
            Select Case entry.Code
                Case "coverage_excluded"
                    If entry.Detail.IndexOf("status: empty", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return name & ": no readable text was extracted. Enable OCR or re-extract this document."
                    If entry.Detail.IndexOf("status: incomplete", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return name & ": extracted text is incomplete and is excluded from search by the current archive policy."
                    If entry.Detail.IndexOf("status: unknown", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return name & ": extraction coverage could not be verified and is excluded from search by the current archive policy."
                    Return name & ": processed, but currently excluded from search."
                Case "failed" : Return name & ": processing failed. Enable Show technical details for the recorded reason."
                Case "pending_host" : Return name & ": waiting for foreground Office processing."
                Case "cooperative_local_only" : Return name & ": processed privately; shared reuse is currently unavailable."
                Case "cooperative_contribution_pending", "shared_claim_deferred" : Return name & ": shared publication is pending."
                Case "permission_reconciliation" : Return name & ": generated-file permissions require attention."
                Case Else : Return name & ": attention required. Enable Show technical details for details."
            End Select
        End Function

        Private Shared Function FormatSourceEntry(entry As SemanticArchiveSourceDiagnostic, archive As SemanticArchiveDefinition,
                         item As SemanticArchiveWorkItem, access As SemanticArchiveAccessContext) As System.String
            If Not CanDiscloseQueuedSource(item, archive, access) Then
                Return entry.Code & ": " & entry.DocumentId & "; current source access is unavailable. Filename, path and diagnostic detail are withheld."
            End If
            Return entry.Code & ": " & DiagnosticLine(System.IO.Path.GetFileName(item.SourcePath)) &
                " | source: " & DiagnosticLine(item.SourcePath) & " | " & DiagnosticLine(entry.Detail) & " [" & entry.DocumentId & "]"
        End Function

        ''' <summary>Administration-only presentation for the current Windows user. No scan, writer, model or dialog is started.</summary>
        Public Shared Function PresentDiagnostics(store As SemanticArchiveStore, archiveId As System.String,
                         messages As System.Collections.Generic.IEnumerable(Of System.String),
                         Optional includeTechnical As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_SHOW_TECHNICAL_DIAGNOSTICS) As System.Collections.Generic.List(Of System.String)
            If store Is Nothing Then Throw New System.ArgumentNullException(NameOf(store))
            Dim values As New System.Collections.Generic.List(Of System.String)()
            If messages Is Nothing Then Return values
            Dim archive As SemanticArchiveDefinition = store.GetArchive(archiveId)
            Dim access As SemanticArchiveAccessContext = SemanticArchiveAccessContext.CreateForCurrentUser()
            Dim generation As SemanticArchiveGenerationManifest = Nothing
            Dim generationAttempted As System.Boolean = False
            Dim sources As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveWorkItem)(System.StringComparer.Ordinal)
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim hidden As System.Int32 = 0
            For Each message As System.String In messages
                If System.String.IsNullOrWhiteSpace(message) OrElse Not seen.Add(message) Then Continue For
                If IsTechnicalDiagnostic(message) Then
                    If Not includeTechnical Then
                        hidden += 1
                        Continue For
                    End If
                    values.Add("[technical] " & DiagnosticLine(message))
                    Continue For
                End If
                Dim entry As SemanticArchiveSourceDiagnostic = ParseSourceDiagnostic(message)
                If entry Is Nothing Then
                    values.Add(BoundProcessingDiagnostic(message))
                    Continue For
                End If
                Dim item As SemanticArchiveWorkItem = Nothing
                If Not sources.TryGetValue(entry.DocumentId, item) Then
                    Try
                        Dim directory As System.String = store.GetWorkDirectory(archiveId)
                        Dim path As System.String = System.IO.Path.Combine(directory, "items", entry.DocumentId & ".json")
                        If System.IO.File.Exists(path) Then
                            SemanticArchivePathGuard.ValidateContainedPath(directory, path, True)
                            SemanticArchiveStore.RequirePrivateArtifact(path)
                            item = SemanticArchiveQueueIndex.ReadHeader(path)
                            If item.DocumentId <> entry.DocumentId Then item = Nothing
                        End If
                        If item Is Nothing Then
                            If Not generationAttempted Then
                                generationAttempted = True
                                generation = store.PinGenerationForAdministration(archiveId)
                            End If
                            Dim document As SemanticArchiveDocumentRecord = If(generation Is Nothing, Nothing, store.LoadDocument(generation, entry.DocumentId))
                            If document IsNot Nothing Then item = New SemanticArchiveWorkItem With {
                                .DocumentId = document.DocumentId, .SourcePath = document.SourcePath,
                                .CanonicalSourceKey = document.CanonicalSourceKey, .BindingIds = document.BindingIds}
                        End If
                    Catch failure As System.Exception
                        ' An unavailable/changed record never grants access to saved details.
                        item = Nothing
                    End Try
                    sources(entry.DocumentId) = item
                End If
                values.Add(If(includeTechnical, FormatSourceEntry(entry, archive, item, access), FormatUserSourceEntry(entry, archive, item, access)))
            Next
            If hidden > 0 Then values.Add(hidden.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                " distinct technical details omitted. Enable Show technical details to include them.")
            Return values
        End Function

        Private Shared Function CaptureExtractionSourceMap(export As Global.SharedLibrary.Agents.TextExportResult) As System.String
            Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                .processed_ranges = export.ProcessedRanges, .page_count = export.PageCount,
                .extraction_coverage_basis = export.ExtractionCoverageBasis,
                .extraction_warnings = Global.SharedLibrary.Agents.TextExtractionDiagnostics.CopyWarnings(export.ExtractionWarnings),
                .ocr_used = export.OcrUsed, .ocr_attempted = export.OcrAttempted, .ocr_skipped = export.OcrSkipped,
                .ocr_duration_milliseconds = export.OcrDurationMilliseconds,
                .options_fingerprint = export.OptionsFingerprint, .configuration_fingerprint = export.ConfigurationFingerprint,
                .processing_signature = export.ProcessingSignature})
        End Function

        Friend Shared Function DescribeExtraction(sourceMapJson As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(sourceMapJson) Then Return "Extraction observations were not recorded. Re-extract this source for current reader diagnostics."
            If sourceMapJson.Length > SharedMethods.DEFAULT_SEMANTICARCHIVE_DIAGNOSTIC_MAXIMUM_BYTES Then Return "Extraction observations exceed the diagnostic limit."
            Try
                Dim map As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(sourceMapJson)
                Dim parts As New System.Collections.Generic.List(Of System.String)()
                For Each field As System.String In New System.String() {"extraction_coverage_basis", "page_count", "ocr_used", "ocr_attempted", "ocr_skipped"}
                    Dim token As Newtonsoft.Json.Linq.JToken = map(field)
                    If token IsNot Nothing AndAlso TypeOf token Is Newtonsoft.Json.Linq.JValue AndAlso token.Type <> Newtonsoft.Json.Linq.JTokenType.Null Then
                        parts.Add(field & "=" & DiagnosticLine(token.ToString()))
                    End If
                Next
                Dim warnings As Newtonsoft.Json.Linq.JArray = TryCast(map("extraction_warnings"), Newtonsoft.Json.Linq.JArray)
                If warnings IsNot Nothing Then
                    Dim count As System.Int32 = 0
                    For Each warning As Newtonsoft.Json.Linq.JToken In warnings
                        If count >= SharedMethods.DEFAULT_TEXTEXPORT_MAXIMUM_COVERAGE_WARNINGS Then Exit For
                        If warning.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                            parts.Add(DiagnosticLine(warning.ToObject(Of System.String)()))
                            count += 1
                        End If
                    Next
                Else
                    parts.Add("Detailed reader observations were not recorded in this older extraction; re-extract this source to obtain them.")
                End If
                Return BoundProcessingDiagnostic(System.String.Join("; ", parts))
            Catch failure As Newtonsoft.Json.JsonException
                Return "Stored extraction observations could not be read; completeness was not changed."
            End Try
        End Function
    End Class
End Namespace
