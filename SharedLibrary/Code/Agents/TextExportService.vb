' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Strict On
Option Explicit On
Option Infer On

Namespace Agents

    Public Enum TextExportStatus
        Unknown = 0
        Readable = 1
        Empty = 2
        Unsupported = 3
        Failed = 4
        Incomplete = 5
    End Enum

    Public NotInheritable Class TextExportOptions
        Public Property OutputFormat As System.String = "text"
        Public Property OcrPdf As System.Boolean
        Public Property OcrBatchPages As System.Int32 = 1
        Public Property Overwrite As System.Boolean
        ' Execution policy only: bypass session reuse for an explicit extraction retry.
        ' This does not change document identity, the content signature or overwrite permission.
        Public Property ForceFreshExtraction As System.Boolean
        ' Host capability, supplied for an explicit foreground job. The adapter dispatches
        ' only its synchronous host reader; enumeration, snapshots and model calls stay off UI.
        Public Property HostReaderDispatcher As System.Func(Of System.Func(Of System.String), System.Threading.CancellationToken, System.Threading.Tasks.Task(Of System.String))
    End Class

    ''' <summary>
    ''' Typed extraction/publication provenance. OutputSha256 covers the complete published file,
    ''' including its UTF-8 BOM; it is deliberately not an indexed text's content-payload hash.
    ''' A Nothing completeness/page/OCR value is unknown. HasReadableText does not assert completeness.
    ''' </summary>
    Public NotInheritable Class TextExportResult
        Public Property Status As TextExportStatus = TextExportStatus.Unknown
        Public Property SourcePath As System.String = System.String.Empty
        Public Property OutputPath As System.String = System.String.Empty
        Public Property SourceSha256 As System.String = System.String.Empty
        Public Property OutputSha256 As System.String = System.String.Empty
        Public Property SourceByteCount As System.Nullable(Of System.Int64)
        Public Property OutputByteCount As System.Nullable(Of System.Int64)
        Public Property CharacterCount As System.Nullable(Of System.Int64)
        Public Property EncodingName As System.String = "utf-8-bom"
        Public Property ExtractorId As System.String = System.String.Empty
        Public Property ExtractorVersion As System.String = System.String.Empty
        Public Property OptionsVersion As System.String = "text-export-options-v1"
        Public Property OptionsFingerprint As System.String = System.String.Empty
        Public Property ConfigurationFingerprint As System.String = System.String.Empty
        Public Property ProcessingSignature As System.String = System.String.Empty
        Public Property ExtractionComplete As System.Nullable(Of System.Boolean)
        Public Property ExtractionCoverageBasis As System.String = "unverified"
        Public Property ExtractionWarnings As New System.Collections.Generic.List(Of System.String)()
        Public Property PageCount As System.Nullable(Of System.Int32)
        Public Property OcrUsed As System.Nullable(Of System.Boolean)
        Public Property OcrAttempted As System.Nullable(Of System.Boolean)
        Public Property OcrSkipped As System.Nullable(Of System.Boolean)
        Public Property OcrDurationMilliseconds As System.Nullable(Of System.Int64)
        Public Property ProcessedRanges As New System.Collections.Generic.List(Of TextExtractionProcessedRange)()
        Public Property SourceAssociationVerified As System.Boolean
        Public Property ErrorCode As System.String = System.String.Empty
        Public Property Message As System.String = System.String.Empty
        Public Property HasReadableText As System.Boolean
        Friend Property LegacyItem As Newtonsoft.Json.Linq.JObject
    End Class

    ''' <summary>
    ''' Reusable single-file boundary. Both ordinary text tools and background consumers enter
    ''' the same source-snapshot, extraction-resource, sandbox-reader and atomic-publication path.
    ''' Legacy JSON is adapted here; callers never interpret extractor error strings as content.
    ''' </summary>
    Public NotInheritable Class TextExportService
        Private Sub New()
        End Sub

        Public Shared Function GetSupportedExtensions() As System.Collections.Generic.IReadOnlyCollection(Of System.String)
            Return TextTools.GetTextExportSupportedExtensions()
        End Function

        ''' <summary>Includes the sibling source snapshot and atomic UTF-8 temporary filename.</summary>
        Friend Shared Function GetRequiredOutputChildPathLength(sourcePath As System.String, outputFileName As System.String) As System.Int32
            Dim sourceSnapshot As System.String = ".ri-source-" & System.Guid.Empty.ToString("N") & System.IO.Path.GetExtension(sourcePath)
            Dim textTemporary As System.String = ".ri-text-" & System.Guid.Empty.ToString("N") & ".tmp"
            Return 1 + System.Math.Max(outputFileName.Length, System.Math.Max(sourceSnapshot.Length, textTemporary.Length))
        End Function

        Public Shared Function GetProcessingSignature(
            context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext,
            sourcePath As System.String,
            options As TextExportOptions
        ) As System.String
            Return TextTools.GetTextExportProcessingSignature(sourcePath, context, If(options, New TextExportOptions()))
        End Function

        Public Shared Async Function ExportFileAsync(
            context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext,
            sourcePath As System.String,
            outputPath As System.String,
            options As TextExportOptions,
            Optional cancellationToken As System.Threading.CancellationToken = Nothing
        ) As System.Threading.Tasks.Task(Of TextExportResult)
            Dim effectiveOptions As TextExportOptions = If(options, New TextExportOptions())
            Dim isolatedContext As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext = context
            If context IsNot Nothing Then isolatedContext = Global.SharedLibrary.SharedLibrary.SharedMethods.CreateIsolatedModelCallContext(context)
            Dim item As Newtonsoft.Json.Linq.JObject = Await TextTools.ExportSingleTextFileCoreAsync(
                sourcePath, outputPath, effectiveOptions.Overwrite, effectiveOptions.OcrPdf,
                isolatedContext, cancellationToken, effectiveOptions).ConfigureAwait(False)
            Dim result As TextExportResult = AdaptLegacyResult(item)
            result.ProcessingSignature = TextTools.BuildTextExportProcessingSignature(result.ExtractorId, result.ExtractorVersion,
                result.ConfigurationFingerprint, result.OptionsFingerprint)
            Return result
        End Function

        ''' <summary>Runs an existing synchronous non-Office STA reader on a bounded private thread.</summary>
        Friend Shared Function RunStaReaderAsync(reader As System.Func(Of System.String),
                                                cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.String)
            If reader Is Nothing Then Throw New System.ArgumentNullException(NameOf(reader))
            cancellationToken.ThrowIfCancellationRequested()
            Dim completion As New System.Threading.Tasks.TaskCompletionSource(Of System.String)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
            Dim worker As New System.Threading.Thread(
                Sub()
                    Try
                        cancellationToken.ThrowIfCancellationRequested()
                        Dim value As System.String = reader()
                        cancellationToken.ThrowIfCancellationRequested()
                        completion.TrySetResult(value)
                    Catch ex As System.OperationCanceledException
                        completion.TrySetCanceled()
                    Catch ex As System.Exception
                        completion.TrySetException(ex)
                    End Try
                End Sub)
            worker.IsBackground = True
            worker.Name = "Red Ink text reader"
            worker.SetApartmentState(System.Threading.ApartmentState.STA)
            worker.Start()
            Return completion.Task
        End Function

        Friend Shared Function AdaptLegacyResult(item As Newtonsoft.Json.Linq.JObject) As TextExportResult
            If item Is Nothing Then
                Return New TextExportResult With {.ErrorCode = "missing_export_result", .Message = "The exporter returned no result."}
            End If
            Dim result As New TextExportResult With {
                .LegacyItem = item,
                .SourcePath = ValueOrEmpty(item, "source_path"), .OutputPath = ValueOrEmpty(item, "output_path"),
                .SourceSha256 = ValueOrEmpty(item, "source_sha256"), .OutputSha256 = ValueOrEmpty(item, "snapshot_sha256"),
                .SourceByteCount = NullableValue(Of System.Int64)(item, "source_byte_count"),
                .OutputByteCount = NullableValue(Of System.Int64)(item, "byte_count"),
                .CharacterCount = NullableValue(Of System.Int64)(item, "char_count"),
                .ExtractorId = ValueOrEmpty(item, "adapter_id"), .ExtractorVersion = ValueOrEmpty(item, "adapter_version"),
                .OptionsFingerprint = ValueOrEmpty(item, "options_fingerprint"),
                .ConfigurationFingerprint = ValueOrEmpty(item, "configuration_fingerprint"),
                .ExtractionComplete = NullableValue(Of System.Boolean)(item, "extraction_complete"),
                .ExtractionCoverageBasis = ValueOrEmpty(item, "extraction_coverage_basis"),
                .PageCount = NullableValue(Of System.Int32)(item, "page_count"),
                .OcrUsed = NullableValue(Of System.Boolean)(item, "ocr_used"),
                .OcrAttempted = NullableValue(Of System.Boolean)(item, "ocr_attempted"),
                .OcrSkipped = NullableValue(Of System.Boolean)(item, "ocr_skipped_due_to_heuristics"),
                .OcrDurationMilliseconds = NullableValue(Of System.Int64)(item, "ocr_duration_ms"),
                .SourceAssociationVerified = NullableValue(Of System.Boolean)(item, "source_association_verified").GetValueOrDefault(),
                .ErrorCode = ValueOrEmpty(item, "error"), .Message = ValueOrEmpty(item, "message")
            }
            Dim ranges As Newtonsoft.Json.Linq.JToken = item("processed_ranges")
            If ranges IsNot Nothing AndAlso ranges.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                For Each range As Newtonsoft.Json.Linq.JToken In ranges
                    result.ProcessedRanges.Add(New TextExtractionProcessedRange With {
                        .StartPage = range.Value(Of System.Int32)("start_page"),
                        .EndPage = range.Value(Of System.Int32)("end_page"),
                        .Association = range.Value(Of System.String)("association")
                    })
                Next
            End If
            Dim warnings As Newtonsoft.Json.Linq.JArray = TryCast(item("extraction_warnings"), Newtonsoft.Json.Linq.JArray)
            If warnings IsNot Nothing Then
                For Each warning As Newtonsoft.Json.Linq.JToken In warnings
                    If warning.Type = Newtonsoft.Json.Linq.JTokenType.String Then TextExtractionDiagnostics.AddWarning(result.ExtractionWarnings, warning.ToObject(Of System.String)())
                Next
            End If
            Dim legacyStatus As System.String = ValueOrEmpty(item, "status")
            If legacyStatus = "skipped_unsupported" OrElse result.ErrorCode.StartsWith("unsupported_", System.StringComparison.Ordinal) OrElse
               result.ErrorCode = "legacy_doc_disabled" Then
                result.Status = TextExportStatus.Unsupported
            ElseIf result.ErrorCode = "empty_pdf_extraction" OrElse result.ErrorCode = "empty_extraction" Then
                result.Status = TextExportStatus.Empty
            ElseIf legacyStatus = "failed" OrElse result.ErrorCode.Length > 0 Then
                result.Status = TextExportStatus.Failed
            ElseIf Not result.SourceAssociationVerified Then
                result.Status = TextExportStatus.Unknown
                result.EncodingName = "unknown"
                result.Message = ValueOrEmpty(item, "warning")
            ElseIf result.CharacterCount.GetValueOrDefault() = 0 Then
                result.Status = TextExportStatus.Empty
            Else
                result.HasReadableText = True
                result.Status = If(result.ExtractionComplete.HasValue AndAlso Not result.ExtractionComplete.Value,
                                   TextExportStatus.Incomplete, TextExportStatus.Readable)
            End If
            Return result
        End Function

        Private Shared Function ValueOrEmpty(item As Newtonsoft.Json.Linq.JObject, name As System.String) As System.String
            Dim value As Newtonsoft.Json.Linq.JToken = item(name)
            If value Is Nothing OrElse value.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return System.String.Empty
            Return value.ToObject(Of System.String)()
        End Function

        Private Shared Function NullableValue(Of TValue As Structure)(item As Newtonsoft.Json.Linq.JObject, name As System.String) As System.Nullable(Of TValue)
            Dim value As Newtonsoft.Json.Linq.JToken = item(name)
            If value Is Nothing OrElse value.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return Nothing
            Return value.ToObject(Of TValue)()
        End Function
    End Class

    ''' <summary>Bounded reader observations; never source text or model response content.</summary>
    Public NotInheritable Class TextExtractionDiagnostics
        Private Sub New()
        End Sub

        Public Shared Sub AddWarning(target As System.Collections.Generic.List(Of System.String), value As System.String)
            If target Is Nothing OrElse System.String.IsNullOrWhiteSpace(value) Then Return
            Dim maximum As System.Int32 = Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_TEXTEXPORT_MAXIMUM_COVERAGE_WARNINGS
            Dim limit As System.Int32 = Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_TEXTEXPORT_COVERAGE_WARNING_CHARACTERS
            Dim normalized As System.String = value.Replace(System.Convert.ToChar(13), " "c).Replace(System.Convert.ToChar(10), " "c).Trim()
            If normalized.Length > limit Then
                Dim length As System.Int32 = limit - 3
                If System.Char.IsHighSurrogate(normalized(length - 1)) Then length -= 1
                normalized = normalized.Substring(0, length) & "..."
            End If
            If target.Contains(normalized) Then Return
            If target.Count < maximum - 1 Then
                target.Add(normalized)
            ElseIf target.Count < maximum Then
                target.Add("Additional reader observations omitted by the diagnostic limit.")
            End If
        End Sub

        ''' <summary>Prioritized bounded exception observations for the existing private reader diagnostics.
        ''' Does not collect request bodies, attachments, model replies or FusionLog.</summary>
        Public Shared Sub AddExceptionDetails(target As System.Collections.Generic.List(Of System.String),
                                              code As System.String, failure As System.Exception,
                                              stage As System.String)
            If target Is Nothing OrElse failure Is Nothing Then Return
            Dim details As New System.Collections.Generic.List(Of System.String)()
            Try
                Dim chain As New System.Collections.Generic.List(Of System.Exception)()
                Dim current As System.Exception = failure
                While current IsNot Nothing AndAlso chain.Count < Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_TEXTEXPORT_EXCEPTION_DIAGNOSTIC_DEPTH
                    chain.Add(current)
                    current = current.InnerException
                End While
                AddWarning(details, code & ": stage=" & stage & "; exception=" & failure.GetType().FullName &
                    "; hresult=" & failure.HResult.ToString("X8", System.Globalization.CultureInfo.InvariantCulture))
                ' Missing dependency identity precedes messages and stacks so a bounded
                ' downstream display can still identify the file inside a wrapper exception.
                For index As System.Int32 = 0 To chain.Count - 1
                    Dim missing As System.IO.FileNotFoundException = TryCast(chain(index), System.IO.FileNotFoundException)
                    Dim loadFailure As System.IO.FileLoadException = TryCast(chain(index), System.IO.FileLoadException)
                    Dim prefix As System.String = code & ": inner_depth=" & index.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    AddExceptionObservation(details, prefix, "file",
                        Function() If(missing IsNot Nothing, missing.FileName, If(loadFailure IsNot Nothing, loadFailure.FileName, Nothing)))
                Next
                For index As System.Int32 = 0 To chain.Count - 1
                    Dim item As System.Exception = chain(index)
                    Dim prefix As System.String = code & ": inner_depth=" & index.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    AddWarning(details, prefix & "; type=" & item.GetType().FullName)
                    AddExceptionObservation(details, prefix, "message", Function() item.Message)
                    AddExceptionObservation(details, prefix, "target", Function() ExceptionTargetName(item))
                Next
                For index As System.Int32 = 0 To chain.Count - 1
                    Dim item As System.Exception = chain(index)
                    Dim prefix As System.String = code & ": inner_depth=" & index.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    ' Formatting TargetSite/StackTrace can itself resolve missing runtime
                    ' dependencies. Isolate each observation after essential file/message data.
                    AddExceptionObservation(details, prefix, "frame", Function() ExceptionStackLines(item))
                Next
                If current IsNot Nothing Then AddWarning(details, code & ": additional inner exceptions omitted by the diagnostic depth limit.")
            Catch diagnosticFailure As System.Exception
                AddObservationFailure(details, code, "exception_chain", diagnosticFailure)
            Finally
                ' Partial diagnostic success must survive later metadata-resolution failure.
                details.AddRange(target)
                Dim bounded As System.Collections.Generic.List(Of System.String) = CopyWarnings(details)
                target.Clear()
                target.AddRange(bounded)
            End Try
        End Sub

        Private Shared Sub AddExceptionObservation(details As System.Collections.Generic.List(Of System.String),
                                                    prefix As System.String, field As System.String,
                                                    readValue As System.Func(Of System.String))
            Try
                Dim value As System.String = readValue.Invoke()
                If Not System.String.IsNullOrWhiteSpace(value) Then AddWarning(details, prefix & "; " & field & "=" & value)
            Catch diagnosticFailure As System.Exception
                AddObservationFailure(details, prefix, field, diagnosticFailure)
            End Try
        End Sub

        Private Shared Sub AddObservationFailure(details As System.Collections.Generic.List(Of System.String),
                                                 prefix As System.String, field As System.String,
                                                 failure As System.Exception)
            AddWarning(details, prefix & "; observation_unavailable=" & field & "; exception=" & failure.GetType().FullName)
            ' Read only basic failure fields here; never reflect over the secondary
            ' exception or format its stack, which could reproduce the same load failure.
            Try
                Dim missing As System.IO.FileNotFoundException = TryCast(failure, System.IO.FileNotFoundException)
                Dim loadFailure As System.IO.FileLoadException = TryCast(failure, System.IO.FileLoadException)
                Dim fileName As System.String = If(missing IsNot Nothing, missing.FileName, If(loadFailure IsNot Nothing, loadFailure.FileName, Nothing))
                If Not System.String.IsNullOrWhiteSpace(fileName) Then AddWarning(details, prefix & "; observation=" & field & "; diagnostic_file=" & fileName)
            Catch basicFailure As System.Exception
                AddWarning(details, prefix & "; diagnostic_file_unavailable=" & basicFailure.GetType().FullName)
            End Try
            Try
                AddWarning(details, prefix & "; observation=" & field & "; diagnostic_message=" & failure.Message)
            Catch basicFailure As System.Exception
                AddWarning(details, prefix & "; diagnostic_message_unavailable=" & basicFailure.GetType().FullName)
            End Try
        End Sub

        Private Shared Function ExceptionTargetName(failure As System.Exception) As System.String
            Dim site As System.Reflection.MethodBase = failure.TargetSite
            If site Is Nothing Then Return Nothing
            Dim owner As System.Type = site.DeclaringType
            Return If(owner Is Nothing, System.String.Empty, owner.FullName & ".") & site.Name
        End Function

        Private Shared Function ExceptionStackLines(failure As System.Exception) As System.String
            Dim lines As New System.Collections.Generic.List(Of System.String)()
            Using reader As New System.IO.StringReader(If(failure.StackTrace, System.String.Empty))
                For frame As System.Int32 = 1 To Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_TEXTEXPORT_EXCEPTION_STACK_FRAMES
                    Dim line As System.String = reader.ReadLine()
                    If line Is Nothing Then Exit For
                    lines.Add(line.Trim())
                Next
            End Using
            Return System.String.Join(" | ", lines)
        End Function

        Public Shared Function CopyWarnings(values As System.Collections.Generic.IEnumerable(Of System.String)) As System.Collections.Generic.List(Of System.String)
            Dim result As New System.Collections.Generic.List(Of System.String)()
            If values IsNot Nothing Then
                For Each value As System.String In values
                    AddWarning(result, value)
                    If result.Count >= Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_TEXTEXPORT_MAXIMUM_COVERAGE_WARNINGS Then Exit For
                Next
            End If
            Return result
        End Function
    End Class
End Namespace
