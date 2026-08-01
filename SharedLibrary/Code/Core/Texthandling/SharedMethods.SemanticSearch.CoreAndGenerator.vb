' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.SemanticSearch.CoreAndGenerator.vb
' Purpose: Provides generic semantic-index data models, JSON utilities, serialized
'          special-task LLM invocation, natural UTF-8 segmentation, and creation
'          of self-indexed text files with content-relative byte offsets. The concept
'          is referred to as "flat semantic search" or FSS.
'
' Architecture:
'  - File Format: Writes a generic versioned index marker, a JSON index document,
'    a content marker, and the exact UTF-8 content bytes used for all offsets.
'  - Segmentation: Selects natural structural boundaries within configurable UTF-8
'    byte limits without splitting surrogate pairs or modifying content afterward.
'  - Metadata: Uses a strictly separated LLM task to describe each segment while
'    application code assigns IDs, order, links, hashes, and byte ranges.
'  - Model Switching: Serializes special-task calls, preserves the current model
'    configuration, applies an available special-task model, and restores state.
'  - Durability: Writes to a temporary file, validates complete byte coverage, and
'    moves or replaces the destination only after successful generation.
' =============================================================================

Option Strict On
Option Explicit On
Option Infer On

Imports System.Linq
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    Partial Public Class SharedMethods


        Public Const SemanticSearchIndexStartMarker As String = "<<<SEMANTIC-SEARCH-INDEX-V1>>>"
        Public Const SemanticSearchContentStartMarker As String = "<<<SEMANTIC-SEARCH-CONTENT-V1>>>"
        Public Const SemanticSearchCurrentFormatVersion As Integer = 1
        Public Const SemanticSearchDefaultGeneratorVersion As String = "1.0.0"

        Private Shared ReadOnly SemanticSearchUtf8NoBom As New System.Text.UTF8Encoding(False, True)
        Private Shared ReadOnly SemanticSearchSpecialTaskSemaphore As New System.Threading.SemaphoreSlim(1, 1)

        ' Generic structured-document wrapper markers (combine-standard separator):
        '   <documentN name="..."> ... </documentN>
        ' The generator prefers these positions as strong structural break candidates
        ' and records which wrapped documents a segment covers.
        Private Shared ReadOnly SemanticSearchDocumentWrapperOpenRegex As New System.Text.RegularExpressions.Regex(
            "<document(\d+)(?:\s+name=""([^""]*)"")?\s*>",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Compiled)
        Private Shared ReadOnly SemanticSearchDocumentWrapperCloseRegex As New System.Text.RegularExpressions.Regex(
            "</document(\d+)\s*>",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Compiled)
        Private Shared ReadOnly SemanticSearchDocumentWrapperEventRegex As New System.Text.RegularExpressions.Regex(
            "(<document(\d+)(?:\s+name=""([^""]*)"")?\s*>)|(</document(\d+)\s*>)",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Compiled)

        Public Class SemanticSearchIndexDocument
            Public Property FormatVersion As Integer
            Public Property Encoding As String = "utf-8"
            Public Property OffsetUnit As String = "byte"
            Public Property OffsetBase As String = "content"
            Public Property ContentSha256 As String = ""
            Public Property CreatedUtc As String = ""
            Public Property GeneratorVersion As String = SemanticSearchDefaultGeneratorVersion
            Public Property SegmentCount As Integer
            Public Property Entries As New System.Collections.Generic.List(Of SemanticSearchIndexEntry)()
        End Class

        Public Class SemanticSearchIndexEntry
            Public Property Id As String = ""
            Public Property Order As Integer
            Public Property Title As String = ""
            Public Property Summary As String = ""
            Public Property Topics As New System.Collections.Generic.List(Of String)()
            Public Property UserIntents As New System.Collections.Generic.List(Of String)()
            Public Property ExactTerms As New System.Collections.Generic.List(Of String)()
            Public Property Actions As New System.Collections.Generic.List(Of String)()
            Public Property Constraints As New System.Collections.Generic.List(Of String)()
            Public Property CrossReferences As New System.Collections.Generic.List(Of String)()
            Public Property SourceDocuments As New System.Collections.Generic.List(Of String)()
            Public Property PreviousId As String = Nothing
            Public Property NextId As String = Nothing
            Public Property RelatedIds As New System.Collections.Generic.List(Of String)()
            Public Property StartByte As Long
            Public Property LengthBytes As Long
        End Class

        Public Class SemanticSearchSegmentMetadataResult
            Public Property Title As String = ""
            Public Property Summary As String = ""
            Public Property Topics As New System.Collections.Generic.List(Of String)()
            Public Property UserIntents As New System.Collections.Generic.List(Of String)()
            Public Property ExactTerms As New System.Collections.Generic.List(Of String)()
            Public Property Actions As New System.Collections.Generic.List(Of String)()
            Public Property Constraints As New System.Collections.Generic.List(Of String)()
            Public Property CrossReferences As New System.Collections.Generic.List(Of String)()
        End Class

        Public Class SemanticSearchIndexGeneratorOptions
            Public Property TargetBytes As Integer = 32 * 1024
            Public Property MinimumBytes As Integer = 16 * 1024
            Public Property MaximumBytes As Integer = 48 * 1024
            Public Property SourceEncoding As System.Text.Encoding = Nothing
            Public Property GeneratorVersion As String = SemanticSearchDefaultGeneratorVersion
            Public Property SpecialTaskName As String = "SemanticSearchIndex"
            Public Property MaximumMetadataAttempts As Integer = 2
            Public Property OverwriteOutput As Boolean = False
        End Class

        Public Class SemanticSearchIndexGenerationProgress
            Public Property SegmentNumber As Integer
            Public Property SegmentCount As Integer
            Public Property SegmentId As String = ""
            Public Property Message As String = ""
        End Class

        Public Class SemanticSearchIndexGenerationResult
            Public Property OutputPath As String = ""
            Public Property ContentByteLength As Long
            Public Property SegmentCount As Integer
            Public Property ContentSha256 As String = ""
            Public Property IndexDocument As SemanticSearchIndexDocument = Nothing
        End Class

        Private Class SemanticSearchRawSegment
            Public Property StartCharacter As Integer
            Public Property CharacterLength As Integer
            Public Property StartByte As Long
            Public Property LengthBytes As Long
            Public Property Text As String = ""
            Public Property SourceDocuments As New System.Collections.Generic.List(Of String)()
        End Class

        Private Class SemanticSearchBreakCandidate
            Public Property Position As Integer
            Public Property StructuralScore As Integer
            Public Property DistanceFromTarget As Integer
        End Class

        ''' <summary>
        ''' Creates one UTF-8 file containing a JSON semantic index followed by the unchanged
        ''' UTF-8 content bytes on which all stored byte offsets are based.
        ''' </summary>
        Public Shared Async Function CreateSemanticSearchIndexedTextFileAsync(
            inputPath As String,
            outputPath As String,
            context As ISharedContext,
            Optional options As SemanticSearchIndexGeneratorOptions = Nothing,
            Optional progress As System.IProgress(Of SemanticSearchIndexGenerationProgress) = Nothing,
            Optional cancellationToken As System.Threading.CancellationToken = Nothing
        ) As System.Threading.Tasks.Task(Of SemanticSearchIndexGenerationResult)

            If String.IsNullOrWhiteSpace(inputPath) Then
                Throw New System.ArgumentException("An input path is required.", NameOf(inputPath))
            End If
            If String.IsNullOrWhiteSpace(outputPath) Then
                Throw New System.ArgumentException("An output path is required.", NameOf(outputPath))
            End If
            If context Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(context))
            End If
            If Not System.IO.File.Exists(inputPath) Then
                Throw New System.IO.FileNotFoundException("The source text file was not found.", inputPath)
            End If

            Dim effectiveOptions As SemanticSearchIndexGeneratorOptions = If(options, New SemanticSearchIndexGeneratorOptions())
            ValidateSemanticSearchGeneratorOptions(effectiveOptions)

            Dim fullInputPath As String = System.IO.Path.GetFullPath(inputPath)
            Dim fullOutputPath As String = System.IO.Path.GetFullPath(outputPath)

            If String.Equals(fullInputPath, fullOutputPath, System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.ArgumentException("Input and output paths must be different.", NameOf(outputPath))
            End If
            If System.IO.File.Exists(fullOutputPath) AndAlso Not effectiveOptions.OverwriteOutput Then
                Throw New System.IO.IOException("The output file already exists. Set OverwriteOutput=True to replace it.")
            End If

            cancellationToken.ThrowIfCancellationRequested()

            Dim originalText As String = ReadSemanticSearchSourceText(fullInputPath, effectiveOptions.SourceEncoding)
            Dim contentBytes As Byte() = SemanticSearchUtf8NoBom.GetBytes(originalText)
            Dim rawSegments As System.Collections.Generic.List(Of SemanticSearchRawSegment) = CreateSemanticSearchRawSegments(originalText, effectiveOptions)

            Dim indexDocument As New SemanticSearchIndexDocument() With {
                .FormatVersion = SemanticSearchCurrentFormatVersion,
                .Encoding = "utf-8",
                .OffsetUnit = "byte",
                .OffsetBase = "content",
                .ContentSha256 = ComputeSemanticSearchSha256Hex(contentBytes),
                .CreatedUtc = System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture),
                .GeneratorVersion = effectiveOptions.GeneratorVersion,
                .SegmentCount = rawSegments.Count
            }

            For segmentIndex As Integer = 0 To rawSegments.Count - 1
                cancellationToken.ThrowIfCancellationRequested()

                Dim rawSegment As SemanticSearchRawSegment = rawSegments(segmentIndex)
                Dim segmentId As String = "S" & (segmentIndex + 1).ToString("0000", System.Globalization.CultureInfo.InvariantCulture)

                If progress IsNot Nothing Then
                    progress.Report(New SemanticSearchIndexGenerationProgress() With {
                        .SegmentNumber = segmentIndex + 1,
                        .SegmentCount = rawSegments.Count,
                        .SegmentId = segmentId,
                        .Message = "Generating semantic metadata"
                    })
                End If

                Dim metadata As SemanticSearchSegmentMetadataResult = Await GenerateSemanticSearchSegmentMetadataAsync(
                    context,
                    effectiveOptions.SpecialTaskName,
                    segmentId,
                    rawSegment.Text,
                    effectiveOptions.MaximumMetadataAttempts,
                    cancellationToken).ConfigureAwait(False)

                indexDocument.Entries.Add(New SemanticSearchIndexEntry() With {
                    .Id = segmentId,
                    .Order = segmentIndex + 1,
                    .Title = CleanSemanticSearchSingleLine(metadata.Title),
                    .Summary = CleanSemanticSearchText(metadata.Summary),
                    .Topics = NormalizeSemanticSearchStringList(metadata.Topics),
                    .UserIntents = NormalizeSemanticSearchStringList(metadata.UserIntents),
                    .ExactTerms = NormalizeSemanticSearchStringList(metadata.ExactTerms),
                    .Actions = NormalizeSemanticSearchStringList(metadata.Actions),
                    .Constraints = NormalizeSemanticSearchStringList(metadata.Constraints),
                    .CrossReferences = NormalizeSemanticSearchStringList(metadata.CrossReferences),
                    .SourceDocuments = New System.Collections.Generic.List(Of String)(rawSegment.SourceDocuments),
                    .RelatedIds = New System.Collections.Generic.List(Of String)(),
                    .StartByte = rawSegment.StartByte,
                    .LengthBytes = rawSegment.LengthBytes
                })
            Next

            For segmentIndex As Integer = 0 To indexDocument.Entries.Count - 1
                indexDocument.Entries(segmentIndex).PreviousId = If(segmentIndex > 0, indexDocument.Entries(segmentIndex - 1).Id, Nothing)
                indexDocument.Entries(segmentIndex).NextId = If(segmentIndex < indexDocument.Entries.Count - 1, indexDocument.Entries(segmentIndex + 1).Id, Nothing)
            Next

            ValidateGeneratedSemanticSearchIndex(indexDocument, contentBytes.LongLength)

            Dim json As String = SerializeSemanticSearchJson(indexDocument)
            Dim headerText As String = SemanticSearchIndexStartMarker & vbLf & json & vbLf & SemanticSearchContentStartMarker & vbLf
            Dim headerBytes As Byte() = SemanticSearchUtf8NoBom.GetBytes(headerText)

            Dim outputDirectory As String = System.IO.Path.GetDirectoryName(fullOutputPath)
            If Not String.IsNullOrWhiteSpace(outputDirectory) Then
                System.IO.Directory.CreateDirectory(outputDirectory)
            End If

            Dim temporaryPath As String = fullOutputPath & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
            Try
                Using outputStream As New System.IO.FileStream(
                    temporaryPath,
                    System.IO.FileMode.CreateNew,
                    System.IO.FileAccess.Write,
                    System.IO.FileShare.None,
                    81920,
                    True)

                    Await outputStream.WriteAsync(headerBytes, 0, headerBytes.Length, cancellationToken).ConfigureAwait(False)
                    Await outputStream.WriteAsync(contentBytes, 0, contentBytes.Length, cancellationToken).ConfigureAwait(False)
                    Await outputStream.FlushAsync(cancellationToken).ConfigureAwait(False)
                End Using

                ValidateWrittenSemanticSearchFile(
                    temporaryPath,
                    headerBytes,
                    contentBytes.LongLength,
                    indexDocument.ContentSha256)

                MoveSemanticSearchTemporaryFileIntoPlace(temporaryPath, fullOutputPath, effectiveOptions.OverwriteOutput)
            Catch ex As System.Exception
                Try
                    If System.IO.File.Exists(temporaryPath) Then
                        System.IO.File.Delete(temporaryPath)
                    End If
                Catch cleanupException As System.Exception
                    System.Diagnostics.Debug.WriteLine(cleanupException.Message)
                End Try
                Throw
            End Try

            Return New SemanticSearchIndexGenerationResult() With {
                .OutputPath = fullOutputPath,
                .ContentByteLength = contentBytes.LongLength,
                .SegmentCount = indexDocument.Entries.Count,
                .ContentSha256 = indexDocument.ContentSha256,
                .IndexDocument = indexDocument
            }
        End Function

        ''' <summary>
        ''' Convenience wrapper that creates a self-indexed text file directly from an in-memory
        ''' string, without the caller having to manage a temporary input file. The text is written
        ''' to a short-named temporary UTF-8 file, indexed into <paramref name="outputPath"/>, and the
        ''' temporary file is always removed. Callers that combine multiple documents should use the
        ''' canonical wrapper separator (&lt;documentN name="..."&gt; ... &lt;/documentN&gt;) so the
        ''' generator can align segments to document boundaries.
        ''' </summary>
        Public Shared Async Function CreateSemanticSearchIndexFromTextAsync(
            text As String,
            outputPath As String,
            context As ISharedContext,
            Optional options As SemanticSearchIndexGeneratorOptions = Nothing,
            Optional progress As System.IProgress(Of SemanticSearchIndexGenerationProgress) = Nothing,
            Optional cancellationToken As System.Threading.CancellationToken = Nothing
        ) As System.Threading.Tasks.Task(Of SemanticSearchIndexGenerationResult)

            If text Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(text))
            End If
            If String.IsNullOrWhiteSpace(outputPath) Then
                Throw New System.ArgumentException("An output path is required.", NameOf(outputPath))
            End If
            If context Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(context))
            End If

            Dim temporaryInputPath As String = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "ss" & System.Guid.NewGuid().ToString("N").Substring(0, 8) & ".txt")

            Try
                Await WriteAllTextAsyncCompat(temporaryInputPath, text, cancellationToken).ConfigureAwait(False)

                Return Await CreateSemanticSearchIndexedTextFileAsync(
                    temporaryInputPath,
                    outputPath,
                    context,
                    options,
                    progress,
                    cancellationToken).ConfigureAwait(False)
            Finally
                Try
                    If System.IO.File.Exists(temporaryInputPath) Then
                        System.IO.File.Delete(temporaryInputPath)
                    End If
                Catch cleanupException As System.Exception
                    System.Diagnostics.Debug.WriteLine(cleanupException.Message)
                End Try
            End Try
        End Function

        ''' <summary>
        ''' Writes UTF-8 text (no BOM) to a file asynchronously. All stored byte offsets rely on
        ''' the exact bytes read back by the generator, so the source encoding used here matches the
        ''' generator's default text reading.
        ''' </summary>
        Private Shared Async Function WriteAllTextAsyncCompat(
            path As String,
            text As String,
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task

            Dim bytes As Byte() = New System.Text.UTF8Encoding(False).GetBytes(If(text, ""))

            Using outputStream As New System.IO.FileStream(
                path,
                System.IO.FileMode.Create,
                System.IO.FileAccess.Write,
                System.IO.FileShare.None,
                81920,
                True)

                Await outputStream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(False)
                Await outputStream.FlushAsync(cancellationToken).ConfigureAwait(False)
            End Using
        End Function

        ''' <summary>
        ''' Executes a serialized LLM call and applies an available special-task model in the
        ''' same manner as the existing application code. The prior configuration is restored.
        ''' </summary>
        Public Shared Async Function CallSemanticSearchSpecialTaskLlmAsync(
            context As ISharedContext,
            specialTaskName As String,
            systemPrompt As String,
            userPrompt As String,
            Optional cancellationToken As System.Threading.CancellationToken = Nothing
        ) As System.Threading.Tasks.Task(Of String)

            If context Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(context))
            End If
            If String.IsNullOrWhiteSpace(specialTaskName) Then
                Throw New System.ArgumentException("A special-task name is required.", NameOf(specialTaskName))
            End If

            cancellationToken.ThrowIfCancellationRequested()
            Await SemanticSearchSpecialTaskSemaphore.WaitAsync(cancellationToken).ConfigureAwait(False)

            Dim restoreConfiguration As System.Action = Nothing
            Dim useSecondApi As Boolean = False
            Dim timeout As Long = context.INI_Timeout

            Try
                If Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                    Dim previousConfiguration = SharedMethods.GetCurrentConfig(context)
                    If previousConfiguration IsNot Nothing Then
                        restoreConfiguration = Sub() SharedMethods.RestoreDefaults(context, previousConfiguration)
                    End If

                    If SharedMethods.GetSpecialTaskModel(context, context.INI_AlternateModelPath, specialTaskName) Then
                        useSecondApi = True
                        timeout = If(context.INI_Timeout_2 > 0, context.INI_Timeout_2, context.INI_Timeout)
                    End If
                End If

                cancellationToken.ThrowIfCancellationRequested()

                Return Await SharedMethods.LLM(
                    context,
                    systemPrompt,
                    userPrompt,
                    "",
                    "",
                    timeout,
                    useSecondApi,
                    True).ConfigureAwait(False)
            Finally
                Try
                    If restoreConfiguration IsNot Nothing Then
                        restoreConfiguration()
                    End If
                Finally
                    SemanticSearchSpecialTaskSemaphore.Release()
                End Try
            End Try
        End Function

        Private Shared Async Function CallSemanticSearchStructuredLlmAsync(Of TResult As Class)(
            context As ISharedContext,
            specialTaskName As String,
            systemPrompt As String,
            userPrompt As String,
            maximumAttempts As Integer,
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of TResult)

            If maximumAttempts < 1 OrElse maximumAttempts > 5 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(maximumAttempts))
            End If

            Dim effectiveUserPrompt As String = If(userPrompt, "")
            Dim lastException As System.Exception = Nothing
            Dim lastRawResponse As String = ""

            For attemptNumber As Integer = 1 To maximumAttempts
                cancellationToken.ThrowIfCancellationRequested()

                Try
                    Dim response As String = Await CallSemanticSearchSpecialTaskLlmAsync(
                        context,
                        specialTaskName,
                        systemPrompt,
                        effectiveUserPrompt,
                        cancellationToken).ConfigureAwait(False)

                    lastRawResponse = If(response, "")

                    Dim result As TResult = DeserializeSemanticSearchJson(Of TResult)(ExtractSemanticSearchJsonObject(response))
                    If result Is Nothing Then
                        Throw New System.FormatException("The LLM returned no usable JSON object.")
                    End If

                    Return result
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    lastException = ex
                    System.Diagnostics.Debug.WriteLine(
                        "Semantic search structured task '" & specialTaskName & "' attempt " &
                        attemptNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) & " failed: " & ex.ToString())
                    If attemptNumber < maximumAttempts Then
                        effectiveUserPrompt &= vbCrLf & vbCrLf &
                            "The previous response could not be parsed. Return exactly one valid JSON object with the requested properties and no surrounding prose."
                    End If
                End Try
            Next

            Throw New System.InvalidOperationException(
                "The structured LLM task failed after " &
                maximumAttempts.ToString(System.Globalization.CultureInfo.InvariantCulture) & " attempts." &
                BuildSemanticSearchStructuredFailureDetail(specialTaskName, lastException, lastRawResponse),
                lastException)
        End Function

        ''' <summary>
        ''' Builds a human-readable failure detail suffix for a structured LLM task, including the
        ''' underlying error (or its type when the message is empty), the inner error, and a short,
        ''' single-line snippet of the last raw model response so the concrete cause is visible
        ''' without attaching a debugger.
        ''' </summary>
        Private Shared Function BuildSemanticSearchStructuredFailureDetail(
            specialTaskName As String,
            lastException As System.Exception,
            lastRawResponse As String
        ) As String

            Dim detail As New System.Text.StringBuilder()

            If Not String.IsNullOrWhiteSpace(specialTaskName) Then
                detail.Append(" Task: " & specialTaskName & ".")
            End If

            If lastException IsNot Nothing Then
                If Not String.IsNullOrWhiteSpace(lastException.Message) Then
                    detail.Append(" Last error: " & lastException.Message)
                Else
                    detail.Append(" Last error: " & lastException.GetType().FullName)
                End If

                If lastException.InnerException IsNot Nothing AndAlso
                   Not String.IsNullOrWhiteSpace(lastException.InnerException.Message) Then
                    detail.Append(" (inner: " & lastException.InnerException.Message & ")")
                End If
            End If

            If Not String.IsNullOrWhiteSpace(lastRawResponse) Then
                Dim snippet As String = lastRawResponse.Replace(vbCr, " "c).Replace(vbLf, " "c).Trim()
                Const maximumSnippetLength As Integer = 300
                If snippet.Length > maximumSnippetLength Then
                    snippet = snippet.Substring(0, maximumSnippetLength) & "…"
                End If
                detail.Append(" Raw response: " & snippet)
            Else
                detail.Append(" Raw response: <empty>")
            End If

            Return detail.ToString()
        End Function

        Private Shared Async Function GenerateSemanticSearchSegmentMetadataAsync(
            context As ISharedContext,
            specialTaskName As String,
            segmentId As String,
            segmentValue As String,
            maximumAttempts As Integer,
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of SemanticSearchSegmentMetadataResult)

            Dim systemPrompt As String =
                "Create semantic directory metadata for one segment of source material. " &
                "Do not answer a user question. Do not create byte positions, IDs, order values or segment links. " &
                "Use only facts present in the supplied segment. Return only one JSON object containing " &
                "Title, Summary, Topics, UserIntents, ExactTerms, Actions, Constraints and CrossReferences. " &
                "Every list property must be a JSON array of strings."

            Dim userPrompt As String =
                "Segment ID for orientation only: " & segmentId & vbCrLf & vbCrLf &
                "Capture topics appearing anywhere in the text, likely user intentions, exact technical and UI terms, " &
                "actions, prerequisites, restrictions, warnings and textual cross-references. " &
                "Do not add unsupported facts." & vbCrLf & vbCrLf &
                "<SEGMENT>" & vbCrLf & segmentValue & vbCrLf & "</SEGMENT>"

            Dim lastException As System.Exception = Nothing

            For attemptNumber As Integer = 1 To maximumAttempts
                cancellationToken.ThrowIfCancellationRequested()
                Try
                    Dim response As String = Await CallSemanticSearchSpecialTaskLlmAsync(
                        context,
                        specialTaskName,
                        systemPrompt,
                        userPrompt,
                        cancellationToken).ConfigureAwait(False)

                    Dim metadata As SemanticSearchSegmentMetadataResult = DeserializeSemanticSearchJson(Of SemanticSearchSegmentMetadataResult)(ExtractSemanticSearchJsonObject(response))
                    If metadata Is Nothing Then
                        Throw New System.InvalidOperationException("The LLM returned no usable metadata object.")
                    End If
                    NormalizeSemanticSearchMetadata(metadata)
                    If String.IsNullOrWhiteSpace(metadata.Title) OrElse
                       String.IsNullOrWhiteSpace(metadata.Summary) Then
                        Throw New System.FormatException("The metadata object must contain a non-empty Title and Summary.")
                    End If
                    Return metadata
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    lastException = ex
                    If attemptNumber < maximumAttempts Then
                        userPrompt &= vbCrLf & vbCrLf &
                            "The previous response could not be parsed or validated. Return exactly one valid JSON object with the requested properties and no surrounding prose."
                    End If
                End Try
            Next

            Throw New System.InvalidOperationException(
                "Semantic metadata generation failed after " & maximumAttempts.ToString(System.Globalization.CultureInfo.InvariantCulture) & " attempts.",
                lastException)
        End Function

        Private Shared Function ReadSemanticSearchSourceText(path As String, sourceEncoding As System.Text.Encoding) As String
            Dim effectiveEncoding As System.Text.Encoding = If(sourceEncoding, SemanticSearchUtf8NoBom)
            Using reader As New System.IO.StreamReader(path, effectiveEncoding, True)
                Return reader.ReadToEnd()
            End Using
        End Function

        Private Shared Function CreateSemanticSearchRawSegments(text As String, options As SemanticSearchIndexGeneratorOptions) As System.Collections.Generic.List(Of SemanticSearchRawSegment)
            Dim result As New System.Collections.Generic.List(Of SemanticSearchRawSegment)()
            If String.IsNullOrEmpty(text) Then
                Return result
            End If

            Dim characterStart As Integer = 0
            Dim byteStart As Long = 0
            Dim currentOpenDocument As String = Nothing

            While characterStart < text.Length
                Dim remainingCharacters As Integer = text.Length - characterStart
                Dim maximumEnd As Integer = FindSemanticSearchLargestUtf8CharacterEnd(text, characterStart, remainingCharacters, options.MaximumBytes)
                If maximumEnd <= characterStart Then
                    maximumEnd = GetSemanticSearchSafeCharacterEnd(text, characterStart + 1)
                End If

                Dim minimumEnd As Integer = FindSemanticSearchLargestUtf8CharacterEnd(text, characterStart, maximumEnd - characterStart, options.MinimumBytes)
                Dim targetEnd As Integer = FindSemanticSearchLargestUtf8CharacterEnd(text, characterStart, maximumEnd - characterStart, options.TargetBytes)
                Dim selectedEnd As Integer = ChooseSemanticSearchNaturalBreak(text, characterStart, minimumEnd, targetEnd, maximumEnd)
                selectedEnd = GetSemanticSearchSafeCharacterEnd(text, selectedEnd)

                If selectedEnd <= characterStart Then
                    selectedEnd = maximumEnd
                End If

                Dim segmentValue As String = text.Substring(characterStart, selectedEnd - characterStart)
                Dim segmentBytes As Byte() = SemanticSearchUtf8NoBom.GetBytes(segmentValue)
                Dim segmentDocuments As System.Collections.Generic.List(Of String) =
                    CollectSemanticSearchSegmentDocuments(segmentValue, currentOpenDocument)

                result.Add(New SemanticSearchRawSegment() With {
                    .StartCharacter = characterStart,
                    .CharacterLength = selectedEnd - characterStart,
                    .StartByte = byteStart,
                    .LengthBytes = segmentBytes.LongLength,
                    .Text = segmentValue,
                    .SourceDocuments = segmentDocuments
                })

                characterStart = selectedEnd
                byteStart += segmentBytes.LongLength
            End While

            Return result
        End Function

        ''' <summary>
        ''' Returns the ordered, distinct wrapped-document labels a segment covers, and updates
        ''' the currently open document as wrapper open/close tags are encountered. Enables
        ''' meaningful citations when many small documents share one segment.
        ''' </summary>
        Private Shared Function CollectSemanticSearchSegmentDocuments(
            segmentText As String,
            ByRef currentOpenDocument As String
        ) As System.Collections.Generic.List(Of String)

            Dim documents As New System.Collections.Generic.List(Of String)()
            If Not String.IsNullOrEmpty(currentOpenDocument) Then
                documents.Add(currentOpenDocument)
            End If

            For Each wrapperMatch As System.Text.RegularExpressions.Match In SemanticSearchDocumentWrapperEventRegex.Matches(segmentText)
                If wrapperMatch.Groups(1).Success Then
                    Dim number As String = wrapperMatch.Groups(2).Value
                    Dim name As String = If(wrapperMatch.Groups(3).Success, wrapperMatch.Groups(3).Value, "")
                    ' The wrapper number is an internal segmentation aid only and must never be
                    ' surfaced to the model or the user. Prefer the document name; fall back to a
                    ' neutral placeholder only when no name was provided by the combine step.
                    Dim label As String = If(String.IsNullOrWhiteSpace(name),
                                             "Unnamed document",
                                             name.Trim())
                    currentOpenDocument = label
                    If Not documents.Contains(label) Then
                        documents.Add(label)
                    End If
                ElseIf wrapperMatch.Groups(4).Success Then
                    currentOpenDocument = Nothing
                End If
            Next

            Return documents
        End Function

        Private Shared Function ChooseSemanticSearchNaturalBreak(
            text As String,
            startIndex As Integer,
            minimumEnd As Integer,
            targetEnd As Integer,
            maximumEnd As Integer
        ) As Integer

            If maximumEnd >= text.Length Then
                Return text.Length
            End If

            Dim lowerBound As Integer = System.Math.Max(startIndex + 1, minimumEnd)
            Dim preferredPosition As Integer = System.Math.Max(lowerBound, targetEnd)
            Dim candidates As New System.Collections.Generic.List(Of SemanticSearchBreakCandidate)()

            Dim scanPosition As Integer = lowerBound
            While scanPosition < maximumEnd
                Dim lineFeedPosition As Integer = text.IndexOf(ControlChars.Lf, scanPosition, maximumEnd - scanPosition)
                If lineFeedPosition < 0 Then
                    Exit While
                End If

                Dim candidatePosition As Integer = lineFeedPosition + 1
                Dim previousLine As String = GetSemanticSearchLineBefore(text, lineFeedPosition)
                Dim nextLine As String = GetSemanticSearchLineAfter(text, candidatePosition)
                Dim structuralScore As Integer = 45
                Dim previousIsList As Boolean = SemanticSearchIsListLine(previousLine)
                Dim nextIsList As Boolean = SemanticSearchIsListLine(nextLine)
                Dim previousIsTable As Boolean = SemanticSearchIsTableLikeLine(previousLine)
                Dim nextIsTable As Boolean = SemanticSearchIsTableLikeLine(nextLine)

                If SemanticSearchLooksLikeChapterHeading(nextLine) Then
                    structuralScore = 170
                ElseIf SemanticSearchLooksLikeHeading(nextLine) Then
                    structuralScore = 160
                ElseIf String.IsNullOrWhiteSpace(previousLine) Then
                    structuralScore = 140
                ElseIf SemanticSearchLooksLikeHeading(previousLine) AndAlso Not String.IsNullOrWhiteSpace(nextLine) Then
                    ' Keep a heading together with its first following paragraph.
                    structuralScore = 0
                ElseIf previousIsList AndAlso nextIsList Then
                    ' Avoid splitting a continuing numbered or bulleted sequence.
                    structuralScore = 5
                ElseIf previousIsList AndAlso Not nextIsList Then
                    structuralScore = 120
                ElseIf previousIsTable AndAlso nextIsTable Then
                    ' Avoid splitting a continuing table-like block.
                    structuralScore = 5
                ElseIf previousIsTable AndAlso Not nextIsTable Then
                    structuralScore = 110
                ElseIf previousLine.TrimEnd().EndsWith(".", System.StringComparison.Ordinal) OrElse
                       previousLine.TrimEnd().EndsWith(":", System.StringComparison.Ordinal) Then
                    structuralScore = 75
                End If

                candidates.Add(New SemanticSearchBreakCandidate() With {
                    .Position = candidatePosition,
                    .StructuralScore = structuralScore,
                    .DistanceFromTarget = System.Math.Abs(candidatePosition - preferredPosition)
                })

                scanPosition = candidatePosition
            End While

            Dim paragraphPatterns As String() = {vbCrLf & vbCrLf, vbLf & vbLf}
            For Each paragraphPattern As String In paragraphPatterns
                Dim searchStart As Integer = lowerBound
                While searchStart < maximumEnd
                    Dim paragraphPosition As Integer = text.IndexOf(paragraphPattern, searchStart, maximumEnd - searchStart, System.StringComparison.Ordinal)
                    If paragraphPosition < 0 Then
                        Exit While
                    End If
                    Dim candidatePosition As Integer = paragraphPosition + paragraphPattern.Length
                    candidates.Add(New SemanticSearchBreakCandidate() With {
                        .Position = candidatePosition,
                        .StructuralScore = 140,
                        .DistanceFromTarget = System.Math.Abs(candidatePosition - preferredPosition)
                    })
                    searchStart = candidatePosition
                End While
            Next

            ' Prefer breaking at structured-document wrapper boundaries so small documents
            ' (e.g. batched emails) are packed to the byte target without being split, and
            ' large documents still break at their own internal headings/paragraphs first.
            Dim wrapperRegion As String = text.Substring(lowerBound, maximumEnd - lowerBound)

            For Each openMatch As System.Text.RegularExpressions.Match In SemanticSearchDocumentWrapperOpenRegex.Matches(wrapperRegion)
                Dim candidatePosition As Integer = lowerBound + openMatch.Index
                candidates.Add(New SemanticSearchBreakCandidate() With {
                    .Position = candidatePosition,
                    .StructuralScore = 175,
                    .DistanceFromTarget = System.Math.Abs(candidatePosition - preferredPosition)
                })
            Next

            For Each closeMatch As System.Text.RegularExpressions.Match In SemanticSearchDocumentWrapperCloseRegex.Matches(wrapperRegion)
                Dim candidatePosition As Integer = lowerBound + closeMatch.Index + closeMatch.Length
                candidates.Add(New SemanticSearchBreakCandidate() With {
                    .Position = candidatePosition,
                    .StructuralScore = 168,
                    .DistanceFromTarget = System.Math.Abs(candidatePosition - preferredPosition)
                })
            Next

            If candidates.Count = 0 Then
                Return maximumEnd
            End If

            Dim selected As SemanticSearchBreakCandidate = candidates.
                Where(Function(candidate As SemanticSearchBreakCandidate) candidate.Position >= lowerBound AndAlso candidate.Position <= maximumEnd).
                OrderByDescending(Function(candidate As SemanticSearchBreakCandidate) candidate.StructuralScore).
                ThenBy(Function(candidate As SemanticSearchBreakCandidate) candidate.DistanceFromTarget).
                FirstOrDefault()

            Return If(selected Is Nothing, maximumEnd, selected.Position)
        End Function

        Private Shared Function FindSemanticSearchLargestUtf8CharacterEnd(
            text As String,
            startIndex As Integer,
            availableCharacters As Integer,
            byteLimit As Integer
        ) As Integer

            Dim boundedEnd As Integer = System.Math.Min(text.Length, startIndex + availableCharacters)
            Dim currentIndex As Integer = startIndex
            Dim byteCount As Integer = 0

            While currentIndex < boundedEnd
                Dim characterCount As Integer = 1

                If System.Char.IsHighSurrogate(text(currentIndex)) Then
                    If currentIndex + 1 >= boundedEnd OrElse
                       Not System.Char.IsLowSurrogate(text(currentIndex + 1)) Then
                        Throw New System.Text.EncoderFallbackException("The source text contains an unmatched UTF-16 surrogate.")
                    End If
                    characterCount = 2
                ElseIf System.Char.IsLowSurrogate(text(currentIndex)) Then
                    Throw New System.Text.EncoderFallbackException("The source text contains an unmatched UTF-16 surrogate.")
                End If

                Dim characterValue As String = text.Substring(currentIndex, characterCount)
                Dim characterByteCount As Integer = SemanticSearchUtf8NoBom.GetByteCount(characterValue)
                If byteCount > byteLimit - characterByteCount Then
                    Exit While
                End If

                byteCount += characterByteCount
                currentIndex += characterCount
            End While

            Return currentIndex
        End Function

        Private Shared Function GetSemanticSearchSafeCharacterEnd(text As String, endIndex As Integer) As Integer
            Dim boundedEnd As Integer = System.Math.Max(0, System.Math.Min(endIndex, text.Length))
            If boundedEnd > 0 AndAlso boundedEnd < text.Length AndAlso
               System.Char.IsHighSurrogate(text(boundedEnd - 1)) AndAlso
               System.Char.IsLowSurrogate(text(boundedEnd)) Then
                boundedEnd -= 1
            End If
            Return boundedEnd
        End Function

        Private Shared Function GetSemanticSearchLineBefore(text As String, lineFeedPosition As Integer) As String
            Dim previousLineFeed As Integer = text.LastIndexOf(ControlChars.Lf, System.Math.Max(0, lineFeedPosition - 1))
            Dim startIndex As Integer = If(previousLineFeed < 0, 0, previousLineFeed + 1)
            Dim length As Integer = lineFeedPosition - startIndex
            Return text.Substring(startIndex, length).TrimEnd(ControlChars.Cr)
        End Function

        Private Shared Function GetSemanticSearchLineAfter(text As String, startIndex As Integer) As String
            If startIndex >= text.Length Then
                Return ""
            End If
            Dim nextLineFeed As Integer = text.IndexOf(ControlChars.Lf, startIndex)
            Dim endIndex As Integer = If(nextLineFeed < 0, text.Length, nextLineFeed)
            Return text.Substring(startIndex, endIndex - startIndex).TrimEnd(ControlChars.Cr)
        End Function

        Private Shared Function SemanticSearchLooksLikeChapterHeading(line As String) As Boolean
            Dim value As String = If(line, "").Trim()
            If value.Length = 0 OrElse value.Length > 140 Then
                Return False
            End If

            Return System.Text.RegularExpressions.Regex.IsMatch(
                value,
                "^(?:(?:chapter|section|part|book|appendix|kapitel|abschnitt|teil|anhang)\s+)?(?:\d+|[IVXLCDM]+)(?:[\.:\)])?\s+\S+",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        End Function

        Private Shared Function SemanticSearchLooksLikeHeading(line As String) As Boolean
            Dim value As String = If(line, "").Trim()
            If value.Length = 0 OrElse value.Length > 140 Then
                Return False
            End If
            If System.Text.RegularExpressions.Regex.IsMatch(value, "^\d+(?:\.\d+)*[\.)]?\s+\S+") Then
                Return True
            End If
            If System.Text.RegularExpressions.Regex.IsMatch(value, "^[A-ZÄÖÜ0-9][A-ZÄÖÜ0-9\s\-–—:/]{3,}$") Then
                Return True
            End If
            If value.EndsWith(":", System.StringComparison.Ordinal) AndAlso value.Count(Function(character As Char) character = " "c) <= 12 Then
                Return True
            End If
            Return Not value.EndsWith(".", System.StringComparison.Ordinal) AndAlso
                   Not value.EndsWith(";", System.StringComparison.Ordinal) AndAlso
                   value.Count(Function(character As Char) character = " "c) <= 10
        End Function

        Private Shared Function SemanticSearchIsListLine(line As String) As Boolean
            Dim value As String = If(line, "").TrimStart()
            Return System.Text.RegularExpressions.Regex.IsMatch(value, "^(?:\d+[\.)]|[A-Za-z][\.)]|[-*•])\s+")
        End Function

        Private Shared Function SemanticSearchIsTableLikeLine(line As String) As Boolean
            Dim value As String = If(line, "")
            Return value.IndexOf(ControlChars.Tab) >= 0 OrElse
                   System.Text.RegularExpressions.Regex.IsMatch(value, "\S\s{2,}\S")
        End Function

        Private Shared Sub ValidateSemanticSearchGeneratorOptions(options As SemanticSearchIndexGeneratorOptions)
            If options.MinimumBytes <= 0 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.MinimumBytes))
            End If
            If options.MaximumBytes < 4 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.MaximumBytes), "MaximumBytes must allow at least one four-byte UTF-8 character.")
            End If
            If options.TargetBytes < options.MinimumBytes Then
                Throw New System.ArgumentException("TargetBytes must be at least MinimumBytes.")
            End If
            If options.MaximumBytes < options.TargetBytes Then
                Throw New System.ArgumentException("MaximumBytes must be at least TargetBytes.")
            End If
            If options.MaximumMetadataAttempts < 1 OrElse options.MaximumMetadataAttempts > 5 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.MaximumMetadataAttempts))
            End If
            If String.IsNullOrWhiteSpace(options.SpecialTaskName) Then
                options.SpecialTaskName = "SemanticSearchIndex"
            End If
            If String.IsNullOrWhiteSpace(options.GeneratorVersion) Then
                options.GeneratorVersion = SemanticSearchDefaultGeneratorVersion
            End If
        End Sub

        Private Shared Sub ValidateGeneratedSemanticSearchIndex(indexDocument As SemanticSearchIndexDocument, contentLength As Long)
            If indexDocument Is Nothing Then
                Throw New System.InvalidOperationException("The generated index document is missing.")
            End If
            If indexDocument.SegmentCount <> indexDocument.Entries.Count Then
                Throw New System.InvalidOperationException("The generated segment count is inconsistent.")
            End If

            Dim expectedStart As Long = 0
            For entryIndex As Integer = 0 To indexDocument.Entries.Count - 1
                Dim entry As SemanticSearchIndexEntry = indexDocument.Entries(entryIndex)
                If entry.Order <> entryIndex + 1 Then
                    Throw New System.InvalidOperationException("The generated segment order is inconsistent.")
                End If
                If entry.StartByte <> expectedStart Then
                    Throw New System.InvalidOperationException("The generated byte ranges contain a gap or overlap.")
                End If
                If entry.LengthBytes <= 0 Then
                    Throw New System.InvalidOperationException("A generated segment has an invalid byte length.")
                End If
                expectedStart += entry.LengthBytes
            Next

            If expectedStart <> contentLength Then
                Throw New System.InvalidOperationException("The generated byte ranges do not cover the complete content.")
            End If
        End Sub

        Private Shared Sub ValidateWrittenSemanticSearchFile(
            path As String,
            expectedHeader As Byte(),
            expectedContentLength As Long,
            expectedContentSha256 As String
        )
            Using stream As New System.IO.FileStream(
                path,
                System.IO.FileMode.Open,
                System.IO.FileAccess.Read,
                System.IO.FileShare.Read)

                If stream.Length <> expectedHeader.LongLength + expectedContentLength Then
                    Throw New System.IO.InvalidDataException("The generated file length is inconsistent.")
                End If

                Dim actualHeader(expectedHeader.Length - 1) As Byte
                ReadSemanticSearchExactly(stream, actualHeader)

                For byteIndex As Integer = 0 To expectedHeader.Length - 1
                    If actualHeader(byteIndex) <> expectedHeader(byteIndex) Then
                        Throw New System.IO.InvalidDataException("The generated file header is inconsistent.")
                    End If
                Next

                Using sha256 As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                    Dim hash As Byte() = sha256.ComputeHash(stream)
                    Dim builder As New System.Text.StringBuilder(hash.Length * 2)
                    For Each value As Byte In hash
                        builder.Append(value.ToString("x2", System.Globalization.CultureInfo.InvariantCulture))
                    Next

                    If Not String.Equals(
                        builder.ToString(),
                        expectedContentSha256,
                        System.StringComparison.OrdinalIgnoreCase) Then
                        Throw New System.IO.InvalidDataException("The generated content hash is inconsistent.")
                    End If
                End Using
            End Using
        End Sub

        Private Shared Sub ReadSemanticSearchExactly(stream As System.IO.Stream, data As Byte())
            Dim offset As Integer = 0
            While offset < data.Length
                Dim readCount As Integer = stream.Read(data, offset, data.Length - offset)
                If readCount = 0 Then
                    Throw New System.IO.EndOfStreamException("The generated file ended unexpectedly.")
                End If
                offset += readCount
            End While
        End Sub

        Private Shared Sub MoveSemanticSearchTemporaryFileIntoPlace(temporaryPath As String, outputPath As String, overwriteOutput As Boolean)
            If Not System.IO.File.Exists(outputPath) Then
                System.IO.File.Move(temporaryPath, outputPath)
                Return
            End If

            If Not overwriteOutput Then
                Throw New System.IO.IOException("The output file already exists.")
            End If

            Try
                System.IO.File.Replace(temporaryPath, outputPath, Nothing, True)
                Return
            Catch ex As System.PlatformNotSupportedException
                System.Diagnostics.Debug.WriteLine(ex.Message)
            Catch ex As System.IO.IOException
                System.Diagnostics.Debug.WriteLine(ex.Message)
            End Try

            Dim backupPath As String = outputPath & "." & System.Guid.NewGuid().ToString("N") & ".bak"
            System.IO.File.Move(outputPath, backupPath)

            Try
                System.IO.File.Move(temporaryPath, outputPath)
            Catch ex As System.Exception
                Try
                    If System.IO.File.Exists(outputPath) Then
                        System.IO.File.Delete(outputPath)
                    End If
                    If System.IO.File.Exists(backupPath) Then
                        System.IO.File.Move(backupPath, outputPath)
                    End If
                Catch restoreException As System.Exception
                    System.Diagnostics.Debug.WriteLine(restoreException.Message)
                End Try
                Throw
            End Try

            Try
                If System.IO.File.Exists(backupPath) Then
                    System.IO.File.Delete(backupPath)
                End If
            Catch cleanupException As System.Exception
                ' The new output is already valid and in place; a backup cleanup failure
                ' must not roll back the successful replacement.
                System.Diagnostics.Debug.WriteLine(cleanupException.Message)
            End Try
        End Sub

        Private Shared Sub NormalizeSemanticSearchMetadata(metadata As SemanticSearchSegmentMetadataResult)
            metadata.Title = CleanSemanticSearchSingleLine(metadata.Title)
            metadata.Summary = CleanSemanticSearchText(metadata.Summary)
            metadata.Topics = NormalizeSemanticSearchStringList(metadata.Topics)
            metadata.UserIntents = NormalizeSemanticSearchStringList(metadata.UserIntents)
            metadata.ExactTerms = NormalizeSemanticSearchStringList(metadata.ExactTerms)
            metadata.Actions = NormalizeSemanticSearchStringList(metadata.Actions)
            metadata.Constraints = NormalizeSemanticSearchStringList(metadata.Constraints)
            metadata.CrossReferences = NormalizeSemanticSearchStringList(metadata.CrossReferences)
        End Sub

        Private Shared Function ComputeSemanticSearchSha256Hex(data As Byte()) As String
            Using sha256 As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Dim hash As Byte() = sha256.ComputeHash(data)
                Dim builder As New System.Text.StringBuilder(hash.Length * 2)
                For Each value As Byte In hash
                    builder.Append(value.ToString("x2", System.Globalization.CultureInfo.InvariantCulture))
                Next
                Return builder.ToString()
            End Using
        End Function

        Private Shared Function NormalizeSemanticSearchStringList(values As System.Collections.Generic.IEnumerable(Of String)) As System.Collections.Generic.List(Of String)
            Dim result As New System.Collections.Generic.List(Of String)()
            If values Is Nothing Then
                Return result
            End If

            Dim seen As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
            For Each value As String In values
                Dim cleanedValue As String = CleanSemanticSearchText(value)
                If cleanedValue.Length > 0 AndAlso seen.Add(cleanedValue) Then
                    result.Add(cleanedValue)
                End If
            Next
            Return result
        End Function

        Private Shared Function CleanSemanticSearchSingleLine(value As String) As String
            Return CleanSemanticSearchText(value).Replace(vbCr, " ").Replace(vbLf, " ").Trim()
        End Function

        Private Shared Function CleanSemanticSearchText(value As String) As String
            Return If(value, "").Trim()
        End Function

        ''' <summary>
        ''' Shared deserialization settings that tolerate common LLM type deviations, such as a
        ''' numeric value returned as a string or a boolean returned as a string/number. This keeps
        ''' the structured retrieval pipeline agnostic and resilient without relying on any specific
        ''' prompt phrasing or language.
        ''' </summary>
        Private Shared ReadOnly SemanticSearchJsonSettings As New Newtonsoft.Json.JsonSerializerSettings() With {
            .Converters = New System.Collections.Generic.List(Of Newtonsoft.Json.JsonConverter)() From {
                New TolerantSemanticSearchDoubleConverter(),
                New TolerantSemanticSearchBooleanConverter()
            }
        }

        Private Shared Function SerializeSemanticSearchJson(Of T)(value As T) As String
            Return Newtonsoft.Json.JsonConvert.SerializeObject(value)
        End Function

        Private Shared Function DeserializeSemanticSearchJson(Of T As Class)(json As String) As T
            If String.IsNullOrWhiteSpace(json) Then
                Return Nothing
            End If

            Return Newtonsoft.Json.JsonConvert.DeserializeObject(Of T)(json, SemanticSearchJsonSettings)
        End Function

        ''' <summary>
        ''' Reads a JSON value as <see cref="System.Double"/> even when the model returns it as a
        ''' string or as a non-numeric explanation. Unparseable or missing values become 0.0 so the
        ''' downstream relevance clamps and thresholds can handle them uniformly.
        ''' </summary>
        Private Class TolerantSemanticSearchDoubleConverter
            Inherits Newtonsoft.Json.JsonConverter

            Public Overrides Function CanConvert(objectType As System.Type) As Boolean
                Return objectType Is GetType(Double) OrElse objectType Is GetType(Double?)
            End Function

            Public Overrides Function ReadJson(
                reader As Newtonsoft.Json.JsonReader,
                objectType As System.Type,
                existingValue As Object,
                serializer As Newtonsoft.Json.JsonSerializer
            ) As Object

                Dim isNullable As Boolean = objectType Is GetType(Double?)

                Select Case reader.TokenType
                    Case Newtonsoft.Json.JsonToken.Null
                        Return If(isNullable, CObj(Nothing), CObj(0.0R))
                    Case Newtonsoft.Json.JsonToken.Float, Newtonsoft.Json.JsonToken.Integer
                        Return System.Convert.ToDouble(reader.Value, System.Globalization.CultureInfo.InvariantCulture)
                    Case Newtonsoft.Json.JsonToken.String
                        Dim parsed As Double
                        If Double.TryParse(
                            System.Convert.ToString(reader.Value, System.Globalization.CultureInfo.InvariantCulture),
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            parsed) Then
                            Return parsed
                        End If
                        Return If(isNullable, CObj(Nothing), CObj(0.0R))
                    Case Newtonsoft.Json.JsonToken.Boolean
                        Return If(System.Convert.ToBoolean(reader.Value), 1.0R, 0.0R)
                    Case Else
                        reader.Skip()
                        Return If(isNullable, CObj(Nothing), CObj(0.0R))
                End Select
            End Function

            Public Overrides ReadOnly Property CanWrite As Boolean
                Get
                    Return False
                End Get
            End Property

            Public Overrides Sub WriteJson(
                writer As Newtonsoft.Json.JsonWriter,
                value As Object,
                serializer As Newtonsoft.Json.JsonSerializer
            )
                Throw New System.NotSupportedException()
            End Sub
        End Class

        ''' <summary>
        ''' Reads a JSON value as <see cref="System.Boolean"/> even when the model returns it as a
        ''' string (for example "true"/"yes"/"1") or as a number. Unrecognized values become False.
        ''' </summary>
        Private Class TolerantSemanticSearchBooleanConverter
            Inherits Newtonsoft.Json.JsonConverter

            Public Overrides Function CanConvert(objectType As System.Type) As Boolean
                Return objectType Is GetType(Boolean) OrElse objectType Is GetType(Boolean?)
            End Function

            Public Overrides Function ReadJson(
                reader As Newtonsoft.Json.JsonReader,
                objectType As System.Type,
                existingValue As Object,
                serializer As Newtonsoft.Json.JsonSerializer
            ) As Object

                Dim isNullable As Boolean = objectType Is GetType(Boolean?)

                Select Case reader.TokenType
                    Case Newtonsoft.Json.JsonToken.Null
                        Return If(isNullable, CObj(Nothing), CObj(False))
                    Case Newtonsoft.Json.JsonToken.Boolean
                        Return System.Convert.ToBoolean(reader.Value)
                    Case Newtonsoft.Json.JsonToken.Integer, Newtonsoft.Json.JsonToken.Float
                        Return System.Convert.ToDouble(reader.Value, System.Globalization.CultureInfo.InvariantCulture) <> 0.0R
                    Case Newtonsoft.Json.JsonToken.String
                        Dim raw As String = System.Convert.ToString(reader.Value, System.Globalization.CultureInfo.InvariantCulture)
                        Dim normalized As String = If(raw, "").Trim().ToLowerInvariant()
                        If normalized = "true" OrElse normalized = "yes" OrElse normalized = "1" Then
                            Return True
                        End If
                        If normalized = "false" OrElse normalized = "no" OrElse normalized = "0" Then
                            Return False
                        End If

                        Dim parsedBoolean As Boolean
                        If Boolean.TryParse(normalized, parsedBoolean) Then
                            Return parsedBoolean
                        End If
                        Return If(isNullable, CObj(Nothing), CObj(False))
                    Case Else
                        reader.Skip()
                        Return If(isNullable, CObj(Nothing), CObj(False))
                End Select
            End Function

            Public Overrides ReadOnly Property CanWrite As Boolean
                Get
                    Return False
                End Get
            End Property

            Public Overrides Sub WriteJson(
                writer As Newtonsoft.Json.JsonWriter,
                value As Object,
                serializer As Newtonsoft.Json.JsonSerializer
            )
                Throw New System.NotSupportedException()
            End Sub
        End Class

        Private Shared Function ExtractSemanticSearchJsonObject(value As String) As String
            Dim text As String = If(value, "").Trim()

            If text.StartsWith("```", System.StringComparison.Ordinal) Then
                Dim firstLineFeed As Integer = text.IndexOf(ControlChars.Lf)
                Dim finalFence As Integer = text.LastIndexOf("```", System.StringComparison.Ordinal)
                If firstLineFeed >= 0 AndAlso finalFence > firstLineFeed Then
                    text = text.Substring(firstLineFeed + 1, finalFence - firstLineFeed - 1).Trim()
                End If
            End If

            Dim objectStart As Integer = text.IndexOf("{"c)
            If objectStart < 0 Then
                Throw New System.FormatException("No JSON object was found in the LLM response.")
            End If

            Dim depth As Integer = 0
            Dim inString As Boolean = False
            Dim escaped As Boolean = False

            For characterIndex As Integer = objectStart To text.Length - 1
                Dim character As Char = text(characterIndex)

                If inString Then
                    If escaped Then
                        escaped = False
                    ElseIf character = "\"c Then
                        escaped = True
                    ElseIf character = """"c Then
                        inString = False
                    End If
                Else
                    If character = """"c Then
                        inString = True
                    ElseIf character = "{"c Then
                        depth += 1
                    ElseIf character = "}"c Then
                        depth -= 1
                        If depth = 0 Then
                            Return text.Substring(objectStart, characterIndex - objectStart + 1)
                        End If
                    End If
                End If
            Next

            Throw New System.FormatException("The JSON object in the LLM response is incomplete.")
        End Function


    End Class

End Namespace
