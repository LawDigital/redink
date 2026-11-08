' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. For license to use see https://redink.ai.
'
' Native-layout preview adapter. A private, byte-verified source snapshot is
' projected to PDF (where necessary) and highlighted only after an unambiguous
' text/geometry match. All fallbacks remain explicit. No source is modified.
Option Explicit On
Option Strict On
Option Infer On

Namespace Global.SharedLibrary.TabularReview
    Public NotInheritable Class ReviewPreviewService
        Private Sub New()
        End Sub

        Private Shared ReadOnly OfficePreviewGate As New System.Threading.SemaphoreSlim(1, 1)

        Private NotInheritable Class OfficePreviewAttempt
            Public ReadOnly Completion As New System.Threading.Tasks.TaskCompletionSource(Of System.String)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
            Public WorkDirectory As System.String
            Public SourcePath As System.String
            Public OutputPath As System.String
            Public Extension As System.String
            Public Abandoned As System.Int32
        End Class

        Private NotInheritable Class NativeWordBox
            Public GlyphIndex As System.Int32
            Public Start As System.Int32
            Public Length As System.Int32
            Public X As System.Double
            Public Y As System.Double
            Public Width As System.Double
            Public Height As System.Double
            Public Baseline As System.Double
        End Class

        Private NotInheritable Class NativeMatch
            Public PageNumber As System.Int32
            Public PageWidth As System.Double
            Public PageHeight As System.Double
            Public Boxes As New System.Collections.Generic.List(Of NativeWordBox)()
            Public Identity As System.String = System.String.Empty
            Public Projection As System.String = System.String.Empty
            Public LayoutNormalized As System.Boolean
        End Class

        Private NotInheritable Class NativeTextProjection
            Public Text As System.String = System.String.Empty
            Public Name As System.String = System.String.Empty
            Public Boxes As New System.Collections.Generic.List(Of NativeWordBox)()
        End Class

        Private NotInheritable Class NativeAnchor
            Public Text As System.String = System.String.Empty
            Public QuoteStart As System.Int32
            Public QuoteLength As System.Int32
            Public Matches As New System.Collections.Generic.Dictionary(Of System.String, NativeMatch)(System.StringComparer.Ordinal)
        End Class

        Private NotInheritable Class NativeLetterComparer
            Implements System.Collections.Generic.IEqualityComparer(Of UglyToad.PdfPig.Content.Letter)

            Public Overloads Function Equals(left As UglyToad.PdfPig.Content.Letter, right As UglyToad.PdfPig.Content.Letter) As System.Boolean Implements System.Collections.Generic.IEqualityComparer(Of UglyToad.PdfPig.Content.Letter).Equals
                Return System.Object.ReferenceEquals(left, right)
            End Function

            Public Overloads Function GetHashCode(value As UglyToad.PdfPig.Content.Letter) As System.Int32 Implements System.Collections.Generic.IEqualityComparer(Of UglyToad.PdfPig.Content.Letter).GetHashCode
                Return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value)
            End Function
        End Class

        Public Shared Async Function CreateAsync(document As ReviewDocument, evidence As ReviewEvidence,
                cacheDirectory As System.String, cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of ReviewPreviewResult)
            Dim result As New ReviewPreviewResult()
            Dim snapshot As System.String = System.String.Empty
            Dim publishedPreview As System.String = System.String.Empty
            Dim keepPreview As System.Boolean = False
            Try
                cancellationToken.ThrowIfCancellationRequested()
                If document Is Nothing Then Throw New System.ArgumentNullException(NameOf(document))
                If System.String.IsNullOrEmpty(document.SourcePath) OrElse System.String.IsNullOrEmpty(document.SourceHash) Then
                    Return Fallback("The original source has not been linked and extracted. The stored review text remains available.")
                End If
                If Not System.IO.File.Exists(document.SourcePath) Then Return Fallback("The source file is unavailable. The stored extraction and answers remain available; reconnect the source folder to restore the original preview.")
                Dim extension As System.String = System.IO.Path.GetExtension(document.SourcePath).ToLowerInvariant()
                If extension <> ".pdf" AndAlso Not IsOfficePreviewExtension(extension) Then
                    Return Fallback("This file type uses the extracted-text viewer.", "info")
                End If
                System.IO.Directory.CreateDirectory(cacheDirectory)
                ' Never trust a browser-supplied path or a persisted cache dependency.
                ' The host passes the registered document and its private session folder.
                snapshot = System.IO.Path.Combine(cacheDirectory, "source-" & System.Guid.NewGuid().ToString("N") & extension)
                Dim hash As System.String = Await System.Threading.Tasks.Task.Run(
                    Function() CopyAndHash(document.SourcePath, snapshot, cancellationToken), cancellationToken).ConfigureAwait(False)
                If Not System.String.Equals(hash, document.SourceHash, System.StringComparison.Ordinal) Then
                    Return Fallback("The file contents have changed since this document was extracted. No new-file preview is combined with the old evidence. Refresh the source before reviewing it; the stored extraction remains available.")
                End If

                Dim nativePdf As System.String = System.IO.Path.Combine(cacheDirectory, "native-" & hash & ".pdf")
                If Not System.IO.File.Exists(nativePdf) Then
                    If extension = ".pdf" Then
                        PublishCopy(snapshot, nativePdf)
                    Else
                        Dim conversionError As System.String = Await ExportOfficePdfAsync(snapshot, extension, nativePdf, cancellationToken).ConfigureAwait(False)
                        If Not System.String.IsNullOrEmpty(conversionError) Then Return Fallback(conversionError)
                    End If
                End If
                cancellationToken.ThrowIfCancellationRequested()
                Dim previewPath As System.String = System.IO.Path.Combine(cacheDirectory, "preview-" & System.Guid.NewGuid().ToString("N") & ".pdf")
                publishedPreview = previewPath
                result = Await System.Threading.Tasks.Task.Run(
                    Function() BuildPreview(nativePdf, previewPath, document, evidence, extension = ".pdf", cancellationToken), cancellationToken).ConfigureAwait(False)
                publishedPreview = result.FilePath
                ' Office conversion can take time. Recheck the live original after
                ' rendering as well; never label a changed source as still current.
                Dim finalHash As System.String = Await System.Threading.Tasks.Task.Run(
                    Function() ReviewSourceService.FileHash(document.SourcePath), cancellationToken).ConfigureAwait(False)
                If Not System.String.Equals(finalHash, document.SourceHash, System.StringComparison.Ordinal) Then Return Fallback("The original file changed while its preview was being prepared. Refresh the source before reviewing it; the stored extraction remains available.")
                cancellationToken.ThrowIfCancellationRequested()
                keepPreview = True
                Return result
            Catch ex As System.OperationCanceledException
                Throw
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine("[TabularReview] native preview failed: " & ex.ToString())
                Return Fallback("The native preview could not be produced: " & ex.Message & " The verified extracted-text viewer remains available.")
            Finally
                DeleteFile(snapshot)
                If Not keepPreview Then DeleteFile(publishedPreview)
            End Try
        End Function

        Private Shared Function IsOfficePreviewExtension(extension As System.String) As System.Boolean
            Return extension = ".docx" OrElse extension = ".docm" OrElse extension = ".doc" OrElse extension = ".rtf" OrElse
                extension = ".xlsx" OrElse extension = ".xlsm" OrElse extension = ".pptx" OrElse extension = ".pptm"
        End Function

        Private Shared Function Fallback(message As System.String, Optional messageKind As System.String = "warning") As ReviewPreviewResult
            System.Diagnostics.Trace.WriteLine("[TabularReview] preview fallback: " & message)
            Return New ReviewPreviewResult With {.Message = message, .MessageKind = messageKind}
        End Function

        Private Shared Function CopyAndHash(source As System.String, destination As System.String,
                                            cancellationToken As System.Threading.CancellationToken) As System.String
            Using input As New System.IO.FileStream(source, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                Using output As New System.IO.FileStream(destination, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None)
                    Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                        Dim buffer(81919) As System.Byte
                        Do
                            cancellationToken.ThrowIfCancellationRequested()
                            Dim count As System.Int32 = input.Read(buffer, 0, buffer.Length)
                            If count = 0 Then Exit Do
                            output.Write(buffer, 0, count)
                            algorithm.TransformBlock(buffer, 0, count, buffer, 0)
                        Loop
                        algorithm.TransformFinalBlock(New System.Byte() {}, 0, 0)
                        Return System.BitConverter.ToString(algorithm.Hash).Replace("-", System.String.Empty).ToLowerInvariant()
                    End Using
                End Using
            End Using
        End Function

        Private Shared Function BuildPreview(inputPdf As System.String, outputPdf As System.String,
                document As ReviewDocument, evidence As ReviewEvidence, isOriginalPdf As System.Boolean,
                cancellationToken As System.Threading.CancellationToken) As ReviewPreviewResult
            Dim result As New ReviewPreviewResult()
            Dim segment As ReviewSourceSegment = Nothing
            Dim verified As System.Boolean = IsVerifiedEvidence(document, evidence, segment)
            Dim pageHint As System.Int32 = If(isOriginalPdf AndAlso segment IsNot Nothing, segment.PageNumber, 0)
            If pageHint > 0 Then result.PageNumber = pageHint
            Dim match As NativeMatch = Nothing
            Dim matchCount As System.Int32 = 0
            Dim matchFailure As System.String = System.String.Empty
            If verified Then
                Try
                    match = FindNativeMatch(inputPdf, document, evidence, segment, pageHint, matchCount, matchFailure, cancellationToken)
                    If matchCount = 0 AndAlso matchFailure.Length = 0 Then
                        matchFailure = "The verified quote could not be mapped to the native PDF text layer. Use Extracted text to inspect its exact stored location."
                    ElseIf matchCount > 1 Then
                        matchFailure = "The quote and its available source context still identify more than one native location. Use Extracted text to inspect the verified location."
                    End If
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    matchFailure = "Native text geometry is unavailable: " & ex.Message & " The verified extracted-text location remains available."
                    System.Diagnostics.Trace.WriteLine("[TabularReview] native text mapping failed: " & ex.ToString())
                End Try
            ElseIf evidence IsNot Nothing Then
                matchFailure = "This source quote has no verified text anchor. The native document is shown without a claimed source highlight."
            End If

            cancellationToken.ThrowIfCancellationRequested()
            Using pdf As PdfSharp.Pdf.PdfDocument = PdfSharp.Pdf.IO.PdfReader.Open(inputPdf, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)
                If pdf.Pages.Count = 0 Then Throw New System.IO.InvalidDataException("The native PDF contains no pages.")
                If result.PageNumber < 1 OrElse result.PageNumber > pdf.Pages.Count Then result.PageNumber = 1
                If match IsNot Nothing AndAlso matchCount = 1 Then
                    Dim page As PdfSharp.Pdf.PdfPage = pdf.Pages(match.PageNumber - 1)
                    result.PageNumber = match.PageNumber
                    Dim geometryIssue As System.String = ValidateGeometry(page, match)
                    If geometryIssue.Length = 0 Then
                        PaintHighlight(page, match.Boxes)
                        result.Highlighted = True
                    Else
                        matchFailure = geometryIssue & " The verified extracted-text location remains available."
                    End If
                End If
                ' These preview copies are display artifacts. Do not execute document
                ' open actions or page actions when a local source is displayed.
                Dim catalog As PdfSharp.Pdf.PdfDictionary = pdf.Internals.Catalog
                catalog.Elements.Remove("/OpenAction")
                catalog.Elements.Remove("/AA")
                For Each page As PdfSharp.Pdf.PdfPage In pdf.Pages
                    page.Elements.Remove("/AA")
                Next
                cancellationToken.ThrowIfCancellationRequested()
                pdf.Save(outputPdf)
            End Using
            result.FilePath = outputPdf
            ' Routine display success needs no extra row in the source viewer.
            ' Keep actual matching failures visible; rendering details stay available as information.
            result.MessageKind = If(Not result.Highlighted AndAlso matchFailure.Length > 0, "warning", "info")
            result.Message = If(result.MessageKind = "warning", matchFailure,
                If(isOriginalPdf, "", "Read-only PDF rendering produced by local Office. Print areas, hidden slides and fonts follow the source application's rendering."))
            If Not result.Highlighted AndAlso evidence IsNot Nothing Then System.Diagnostics.Trace.WriteLine("[TabularReview] native highlight fallback: " & result.Message)
            Return result
        End Function

        Private Shared Function IsVerifiedEvidence(document As ReviewDocument, evidence As ReviewEvidence,
                                                   ByRef segment As ReviewSourceSegment) As System.Boolean
            segment = Nothing
            If evidence Is Nothing OrElse document.Segments Is Nothing Then Return False
            If evidence.MatchStatus <> "exact" AndAlso evidence.MatchStatus <> "whitespace" Then Return False
            segment = document.Segments.Find(Function(item As ReviewSourceSegment) item.Id = evidence.SegmentId)
            If segment Is Nothing OrElse evidence.Start < 0 OrElse evidence.Length <= 0 OrElse
                CLng(evidence.Start) + evidence.Length > segment.Text.Length OrElse System.String.IsNullOrWhiteSpace(evidence.Quote) Then Return False
            Dim actual As System.String = segment.Text.Substring(evidence.Start, evidence.Length)
            If Not HasCompleteTokenBoundaries(segment.Text, evidence.Start, evidence.Length) Then Return False
            Return System.String.Equals(NormalizeForMatch(actual, Nothing), NormalizeForMatch(evidence.Quote, Nothing), System.StringComparison.Ordinal)
        End Function

        Private Shared Function FindNativeMatch(path As System.String, document As ReviewDocument, evidence As ReviewEvidence,
                segment As ReviewSourceSegment, pageHint As System.Int32, ByRef matchCount As System.Int32,
                ByRef failure As System.String, cancellationToken As System.Threading.CancellationToken) As NativeMatch
            matchCount = 0
            Dim anchors As System.Collections.Generic.List(Of NativeAnchor) = BuildSourceAnchors(document, evidence, segment, pageHint, cancellationToken)
            If anchors.Count = 0 Then
                failure = "The stored quote and surrounding text do not identify a unique source location. Use Extracted text to inspect the verified anchor."
                Return Nothing
            End If
            Using pdf As UglyToad.PdfPig.PdfDocument = UglyToad.PdfPig.PdfDocument.Open(path)
                Dim first As System.Int32 = If(pageHint > 0, pageHint, 1)
                Dim last As System.Int32 = If(pageHint > 0, pageHint, pdf.NumberOfPages)
                If first > pdf.NumberOfPages Then Return Nothing
                For pageNumber As System.Int32 = first To last
                    cancellationToken.ThrowIfCancellationRequested()
                    Dim page As UglyToad.PdfPig.Content.Page = pdf.GetPage(pageNumber)
                    For Each projection As NativeTextProjection In BuildNativeProjections(page, cancellationToken)
                        Dim offsets As New System.Collections.Generic.List(Of System.Int32)()
                        Dim haystack As System.String = NormalizeNativeText(projection.Text, offsets)
                        For Each anchor As NativeAnchor In anchors
                            ' Two distinct locations suffice to mark this context as ambiguous.
                            ' Longer source contexts still get their own independent search.
                            If anchor.Matches.Count > 1 Then Continue For
                            Dim searchFrom As System.Int32 = 0
                            Do While searchFrom <= haystack.Length - anchor.Text.Length
                                cancellationToken.ThrowIfCancellationRequested()
                                Dim index As System.Int32 = haystack.IndexOf(anchor.Text, searchFrom, System.StringComparison.Ordinal)
                                If index < 0 Then Exit Do
                                Dim quoteIndex As System.Int32 = index + anchor.QuoteStart
                                If HasCompleteTokenBoundaries(haystack, index, anchor.Text.Length) AndAlso
                                    HasCompleteTokenBoundaries(haystack, quoteIndex, anchor.QuoteLength) Then
                                    Dim start As System.Int32 = offsets(quoteIndex)
                                    Dim ending As System.Int32 = offsets(quoteIndex + anchor.QuoteLength - 1) + 1
                                    Dim found As NativeMatch = CreateNativeMatch(page, projection, start, ending)
                                    If found IsNot Nothing AndAlso Not anchor.Matches.ContainsKey(found.Identity) Then
                                        anchor.Matches.Add(found.Identity, found)
                                        If anchor.Matches.Count > 1 Then Exit Do
                                    End If
                                End If
                                searchFrom = index + 1
                            Loop
                        Next
                    Next
                Next
            End Using
            Dim resolved As New System.Collections.Generic.Dictionary(Of System.String, NativeMatch)(System.StringComparer.Ordinal)
            For Each anchor As NativeAnchor In anchors
                matchCount = System.Math.Max(matchCount, anchor.Matches.Count)
                If anchor.Matches.Count <> 1 Then Continue For
                For Each item As System.Collections.Generic.KeyValuePair(Of System.String, NativeMatch) In anchor.Matches
                    If Not resolved.ContainsKey(item.Key) Then resolved.Add(item.Key, item.Value)
                Next
            Next
            ' Agreement is by physical glyph identity, never by projection order.
            ' An exact, longer context can disambiguate a repeated quotation.
            If resolved.Count <> 1 Then
                If resolved.Count > 1 Then matchCount = 2
                Return Nothing
            End If
            For Each result As NativeMatch In resolved.Values
                matchCount = 1
                Dim actualSource As System.String = segment.Text.Substring(evidence.Start, evidence.Length)
                result.LayoutNormalized = result.LayoutNormalized OrElse Not System.String.Equals(NormalizeForMatch(actualSource, Nothing), NormalizeNativeText(actualSource, Nothing), System.StringComparison.Ordinal)
                System.Diagnostics.Trace.WriteLine("[TabularReview] native highlight mapped: page=" & result.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) & "; projection=" & result.Projection & "; glyphs=" & result.Boxes.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & "; typographyNormalized=" & result.LayoutNormalized.ToString())
                Return result
            Next
            Return Nothing
        End Function

        Private Shared Function BuildSourceAnchors(document As ReviewDocument, evidence As ReviewEvidence,
                segment As ReviewSourceSegment, pageHint As System.Int32,
                cancellationToken As System.Threading.CancellationToken) As System.Collections.Generic.List(Of NativeAnchor)
            Dim result As New System.Collections.Generic.List(Of NativeAnchor)()
            Dim normalizedSources As New System.Collections.Generic.List(Of System.String)()
            For Each source As ReviewSourceSegment In document.Segments
                cancellationToken.ThrowIfCancellationRequested()
                If pageHint > 0 AndAlso source.PageNumber <> pageHint Then Continue For
                normalizedSources.Add(NormalizeNativeText(source.Text, Nothing))
            Next
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim segmentOffsets As New System.Collections.Generic.List(Of System.Int32)()
            Dim normalizedSegment As System.String = NormalizeNativeText(segment.Text, segmentOffsets)
            Dim quoteEnd As System.Int32 = evidence.Start + evidence.Length
            For Each radius As System.Int32 In New System.Int32() {0, 64, 160, 384, 1024}
                Dim beginning As System.Int32 = System.Math.Max(0, evidence.Start - radius)
                Dim ending As System.Int32 = System.Math.Min(segment.Text.Length, quoteEnd + radius)
                If beginning > 0 AndAlso System.Char.IsLowSurrogate(segment.Text(beginning)) AndAlso System.Char.IsHighSurrogate(segment.Text(beginning - 1)) Then beginning -= 1
                If ending < segment.Text.Length AndAlso System.Char.IsHighSurrogate(segment.Text(ending - 1)) AndAlso System.Char.IsLowSurrogate(segment.Text(ending)) Then ending += 1
                While beginning > 0 AndAlso IsWordCharacter(segment.Text(beginning - 1)) AndAlso IsWordCharacter(segment.Text(beginning))
                    beginning -= 1
                End While
                While ending < segment.Text.Length AndAlso IsWordCharacter(segment.Text(ending - 1)) AndAlso IsWordCharacter(segment.Text(ending))
                    ending += 1
                End While
                Dim offsets As New System.Collections.Generic.List(Of System.Int32)()
                Dim text As System.String = NormalizeNativeText(segment.Text.Substring(beginning, ending - beginning), offsets)
                If text.Length = 0 OrElse Not seen.Add(text) Then Continue For
                Dim anchorStart As System.Int32 = segmentOffsets.IndexOf(beginning + offsets(0))
                ' A clipped context must still start at this actual source offset.
                ' Never borrow another occurrence when a cut splits a signed amount.
                If anchorStart < 0 OrElse CLng(anchorStart) + text.Length > normalizedSegment.Length OrElse
                    System.String.CompareOrdinal(normalizedSegment, anchorStart, text, 0, text.Length) <> 0 OrElse
                    Not HasCompleteTokenBoundaries(normalizedSegment, anchorStart, text.Length) Then Continue For
                Dim quoteStart As System.Int32 = -1
                Dim quoteLast As System.Int32 = -1
                For index As System.Int32 = 0 To offsets.Count - 1
                    Dim original As System.Int32 = beginning + offsets(index)
                    If original < evidence.Start OrElse original >= quoteEnd OrElse System.Char.IsWhiteSpace(text(index)) Then Continue For
                    If quoteStart < 0 Then quoteStart = index
                    quoteLast = index
                Next
                If quoteStart < 0 Then Continue For
                Dim occurrences As System.Int32 = 0
                For Each source As System.String In normalizedSources
                    Dim from As System.Int32 = 0
                    Do While from <= source.Length - text.Length
                        cancellationToken.ThrowIfCancellationRequested()
                        Dim index As System.Int32 = source.IndexOf(text, from, System.StringComparison.Ordinal)
                        If index < 0 Then Exit Do
                        If HasCompleteTokenBoundaries(source, index, text.Length) Then occurrences += 1
                        If occurrences > 1 Then Exit Do
                        from = index + 1
                    Loop
                    If occurrences > 1 Then Exit For
                Next
                If occurrences = 1 Then result.Add(New NativeAnchor With {.Text = text, .QuoteStart = quoteStart, .QuoteLength = quoteLast - quoteStart + 1})
            Next
            Return result
        End Function

        Private Shared Function BuildNativeProjections(page As UglyToad.PdfPig.Content.Page,
                cancellationToken As System.Threading.CancellationToken) As System.Collections.Generic.List(Of NativeTextProjection)
            Dim result As New System.Collections.Generic.List(Of NativeTextProjection)()
            Dim glyphs As New System.Collections.Generic.Dictionary(Of UglyToad.PdfPig.Content.Letter, System.Int32)(New NativeLetterComparer())
            For index As System.Int32 = 0 To page.Letters.Count - 1
                If Not glyphs.ContainsKey(page.Letters(index)) Then glyphs.Add(page.Letters(index), index)
            Next
            ' The shared PDF reader uses this extractor first. Match its output back
            ' to the very same letters instead of guessing word enumeration order.
            Try
                Dim text As System.String = UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor.ContentOrderTextExtractor.GetText(page)
                AddProjection(result, MapNativeProjection(text, page.Letters, glyphs, "content order", cancellationToken))
            Catch ex As System.OperationCanceledException
                Throw
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine("[TabularReview] content-order preview mapping unavailable: " & ex.ToString())
            End Try
            cancellationToken.ThrowIfCancellationRequested()
            Try
                Dim words As New System.Collections.Generic.List(Of UglyToad.PdfPig.Content.Word)(page.GetWords(UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor.NearestNeighbourWordExtractor.Instance))
                If words.Count = 0 Then Return result
                AddProjection(result, MapWordProjection(words, glyphs, "word order", cancellationToken))
                ' Keep the shared reader's line-based fallback: baseline groups, top to
                ' bottom, with words left to right within each group.
                Dim ordered As New System.Collections.Generic.List(Of UglyToad.PdfPig.Content.Word)(words)
                ordered.Sort(Function(left As UglyToad.PdfPig.Content.Word, right As UglyToad.PdfPig.Content.Word) As System.Int32
                                 Dim order As System.Int32 = right.BoundingBox.Bottom.CompareTo(left.BoundingBox.Bottom)
                                 Return If(order <> 0, order, left.BoundingBox.Left.CompareTo(right.BoundingBox.Left))
                             End Function)
                Dim lines As New System.Collections.Generic.List(Of System.Collections.Generic.List(Of UglyToad.PdfPig.Content.Word))()
                Dim tolerance As System.Double = System.Math.Max(0.5, page.Height * 0.002)
                For Each word As UglyToad.PdfPig.Content.Word In ordered
                    cancellationToken.ThrowIfCancellationRequested()
                    Dim destination As System.Collections.Generic.List(Of UglyToad.PdfPig.Content.Word) = Nothing
                    For Each line As System.Collections.Generic.List(Of UglyToad.PdfPig.Content.Word) In lines
                        If System.Math.Abs(word.BoundingBox.Bottom - line(0).BoundingBox.Bottom) <= tolerance Then
                            destination = line
                            Exit For
                        End If
                    Next
                    If destination Is Nothing Then
                        destination = New System.Collections.Generic.List(Of UglyToad.PdfPig.Content.Word)()
                        lines.Add(destination)
                    End If
                    destination.Add(word)
                Next
                ordered.Clear()
                For Each line As System.Collections.Generic.List(Of UglyToad.PdfPig.Content.Word) In lines
                    line.Sort(Function(left As UglyToad.PdfPig.Content.Word, right As UglyToad.PdfPig.Content.Word) left.BoundingBox.Left.CompareTo(right.BoundingBox.Left))
                    ordered.AddRange(line)
                Next
                AddProjection(result, MapWordProjection(ordered, glyphs, "line order", cancellationToken))
                Try
                    ' Reuse the layout primitives already used by PdfMarkdownExtractor.
                    ' Blocks prevent a side column from interrupting a multi-line quote.
                    Dim blocks = UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter.DocstrumBoundingBoxes.Instance.GetBlocks(words)
                    Dim readingOrder = UglyToad.PdfPig.DocumentLayoutAnalysis.ReadingOrderDetector.UnsupervisedReadingOrderDetector.Instance.Get(blocks)
                    ordered.Clear()
                    For Each block In readingOrder
                        cancellationToken.ThrowIfCancellationRequested()
                        For Each line In block.TextLines
                            ordered.AddRange(line.Words)
                        Next
                    Next
                    AddProjection(result, MapWordProjection(ordered, glyphs, "block order", cancellationToken))
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    System.Diagnostics.Trace.WriteLine("[TabularReview] block-order preview mapping unavailable: " & ex.ToString())
                End Try
            Catch ex As System.OperationCanceledException
                Throw
            Catch ex As System.Exception
                ' An optional layout strategy must not discard a fully mapped
                ' content-order projection that the shared reader already supports.
                System.Diagnostics.Trace.WriteLine("[TabularReview] word-layout preview mapping unavailable: " & ex.ToString())
            End Try
            Return result
        End Function

        Private Shared Function MapWordProjection(words As System.Collections.Generic.IEnumerable(Of UglyToad.PdfPig.Content.Word),
                glyphs As System.Collections.Generic.Dictionary(Of UglyToad.PdfPig.Content.Letter, System.Int32), name As System.String,
                cancellationToken As System.Threading.CancellationToken) As NativeTextProjection
            Dim text As New System.Text.StringBuilder()
            Dim letters As New System.Collections.Generic.List(Of UglyToad.PdfPig.Content.Letter)()
            For Each word As UglyToad.PdfPig.Content.Word In words
                cancellationToken.ThrowIfCancellationRequested()
                If System.String.IsNullOrEmpty(word.Text) Then Continue For
                If text.Length > 0 Then text.Append(" "c)
                text.Append(word.Text)
                letters.AddRange(word.Letters)
            Next
            Return MapNativeProjection(text.ToString(), letters, glyphs, name, cancellationToken)
        End Function

        Private Shared Function MapNativeProjection(text As System.String, letters As System.Collections.Generic.IEnumerable(Of UglyToad.PdfPig.Content.Letter),
                glyphs As System.Collections.Generic.Dictionary(Of UglyToad.PdfPig.Content.Letter, System.Int32), name As System.String,
                cancellationToken As System.Threading.CancellationToken) As NativeTextProjection
            Dim result As New NativeTextProjection With {.Text = If(text, System.String.Empty), .Name = name}
            Dim position As System.Int32 = 0
            For Each letter As UglyToad.PdfPig.Content.Letter In letters
                cancellationToken.ThrowIfCancellationRequested()
                Dim beginning As System.Int32 = -1
                For Each character As System.Char In If(letter.Value, System.String.Empty)
                    If System.Char.IsWhiteSpace(character) Then Continue For
                    While position < result.Text.Length AndAlso System.Char.IsWhiteSpace(result.Text(position))
                        position += 1
                    End While
                    If position >= result.Text.Length OrElse result.Text(position) <> character Then
                        System.Diagnostics.Trace.WriteLine("[TabularReview] rejected preview projection without complete glyph alignment: " & name)
                        Return Nothing
                    End If
                    If beginning < 0 Then beginning = position
                    position += 1
                Next
                If beginning < 0 Then Continue For
                Dim glyphIndex As System.Int32
                If Not glyphs.TryGetValue(letter, glyphIndex) Then Throw New System.IO.InvalidDataException("A native projection contains a glyph from a different page.")
                ' PdfRectangle.Width/Height are edge lengths on rotated glyphs;
                ' derive the axis-aligned envelope from all four actual corners.
                Dim rectangle As UglyToad.PdfPig.Core.PdfRectangle = letter.GlyphRectangle
                Dim left As System.Double = System.Math.Min(System.Math.Min(rectangle.TopLeft.X, rectangle.TopRight.X), System.Math.Min(rectangle.BottomLeft.X, rectangle.BottomRight.X))
                Dim right As System.Double = System.Math.Max(System.Math.Max(rectangle.TopLeft.X, rectangle.TopRight.X), System.Math.Max(rectangle.BottomLeft.X, rectangle.BottomRight.X))
                Dim bottom As System.Double = System.Math.Min(System.Math.Min(rectangle.TopLeft.Y, rectangle.TopRight.Y), System.Math.Min(rectangle.BottomLeft.Y, rectangle.BottomRight.Y))
                Dim top As System.Double = System.Math.Max(System.Math.Max(rectangle.TopLeft.Y, rectangle.TopRight.Y), System.Math.Max(rectangle.BottomLeft.Y, rectangle.BottomRight.Y))
                result.Boxes.Add(New NativeWordBox With {
                    .GlyphIndex = glyphIndex, .Start = beginning, .Length = position - beginning,
                    .X = left, .Y = bottom, .Width = right - left, .Height = top - bottom,
                    .Baseline = letter.StartBaseLine.Y
                })
            Next
            While position < result.Text.Length AndAlso System.Char.IsWhiteSpace(result.Text(position))
                position += 1
            End While
            If position <> result.Text.Length Then
                System.Diagnostics.Trace.WriteLine("[TabularReview] rejected preview projection with unmapped text: " & name)
                Return Nothing
            End If
            Return result
        End Function

        Private Shared Sub AddProjection(projections As System.Collections.Generic.List(Of NativeTextProjection), projection As NativeTextProjection)
            If projection IsNot Nothing AndAlso projection.Text.Length > 0 AndAlso projection.Boxes.Count > 0 Then projections.Add(projection)
        End Sub

        Private Shared Function CreateNativeMatch(page As UglyToad.PdfPig.Content.Page, projection As NativeTextProjection,
                start As System.Int32, ending As System.Int32) As NativeMatch
            Dim found As New NativeMatch With {.PageNumber = page.Number, .PageWidth = page.Width, .PageHeight = page.Height, .Projection = projection.Name}
            Dim ids As New System.Collections.Generic.List(Of System.Int32)()
            Dim seen As New System.Collections.Generic.HashSet(Of System.Int32)()
            For Each box As NativeWordBox In projection.Boxes
                If box.Start >= ending OrElse box.Start + box.Length <= start OrElse Not seen.Add(box.GlyphIndex) Then Continue For
                If (box.Width <= 0 OrElse box.Height <= 0) AndAlso projection.Text.Substring(box.Start, box.Length).Trim(Microsoft.VisualBasic.ChrW(&HAD)).Length = 0 Then Continue For
                ids.Add(box.GlyphIndex)
                found.Boxes.Add(box)
            Next
            If found.Boxes.Count = 0 Then Return Nothing
            ids.Sort()
            found.Identity = page.Number.ToString(System.Globalization.CultureInfo.InvariantCulture) & ":" & System.String.Join(",", ids)
            Dim actual As System.String = projection.Text.Substring(start, ending - start)
            found.LayoutNormalized = Not System.String.Equals(NormalizeForMatch(actual, Nothing), NormalizeNativeText(actual, Nothing), System.StringComparison.Ordinal)
            Return found
        End Function

        ''' <summary>Native display only: keep letters, digits, case and punctuation
        ''' significant. Expand alphabetic presentation ligatures and discretionary
        ''' soft hyphens; never apply compatibility/fuzzy or letters-only matching.
        ''' Evidence verification continues to use NormalizeForMatch.</summary>
        Private Shared Function NormalizeNativeText(value As System.String, offsets As System.Collections.Generic.List(Of System.Int32)) As System.String
            Dim text As System.String = If(value, System.String.Empty)
            Dim builder As New System.Text.StringBuilder(text.Length)
            Dim pendingSpace As System.Int32 = -1
            Dim index As System.Int32 = 0
            While index < text.Length
                Dim character As System.Char = text(index)
                If character = Microsoft.VisualBasic.ChrW(&HAD) Then
                    Dim following As System.Int32 = index + 1
                    While following < text.Length AndAlso System.Char.IsWhiteSpace(text(following))
                        following += 1
                    End While
                    If builder.Length > 0 AndAlso pendingSpace < 0 AndAlso System.Char.IsLetter(builder(builder.Length - 1)) AndAlso
                        following < text.Length AndAlso System.Char.IsLetter(text(following)) Then
                        index = following
                        Continue While
                    End If
                    index += 1
                    Continue While
                End If
                If System.Char.IsWhiteSpace(character) Then
                    If builder.Length > 0 AndAlso pendingSpace < 0 Then pendingSpace = index
                    index += 1
                    Continue While
                End If
                If pendingSpace >= 0 Then
                    builder.Append(" "c)
                    If offsets IsNot Nothing Then offsets.Add(pendingSpace)
                    pendingSpace = -1
                End If
                Dim expanded As System.String = character.ToString()
                Select Case System.Convert.ToInt32(character)
                    Case &HFB00 : expanded = "ff"
                    Case &HFB01 : expanded = "fi"
                    Case &HFB02 : expanded = "fl"
                    Case &HFB03 : expanded = "ffi"
                    Case &HFB04 : expanded = "ffl"
                    Case &HFB05, &HFB06 : expanded = "st"
                End Select
                builder.Append(expanded)
                If offsets IsNot Nothing Then
                    For count As System.Int32 = 1 To expanded.Length
                        offsets.Add(index)
                    Next
                End If
                index += 1
            End While
            Return builder.ToString()
        End Function

        Private Shared Function HasCompleteTokenBoundaries(text As System.String, start As System.Int32, length As System.Int32) As System.Boolean
            If length <= 0 OrElse start < 0 OrElse CLng(start) + length > text.Length Then Return False
            Dim ending As System.Int32 = start + length
            ' Padding in a quotation must not hide a partial numeric token.
            ' Check its first and last substantive characters in the original text.
            While start < ending AndAlso System.Char.IsWhiteSpace(text(start))
                start += 1
            End While
            While ending > start AndAlso System.Char.IsWhiteSpace(text(ending - 1))
                ending -= 1
            End While
            If start = ending Then Return False
            length = ending - start
            If start > 0 AndAlso IsWordCharacter(text(start - 1)) AndAlso IsWordCharacter(text(start)) Then Return False
            If ending < text.Length AndAlso IsWordCharacter(text(ending - 1)) AndAlso IsWordCharacter(text(ending)) Then Return False
            ' Adjacent digit groups separated only by whitespace can belong to
            ' one amount. Treat them conservatively as one potential numeric token,
            ' before and after normalization; no locale-specific grouping is assumed.
            If System.Char.IsDigit(text(start)) Then
                Dim preceding As System.Int32 = start - 1
                While preceding >= 0 AndAlso System.Char.IsWhiteSpace(text(preceding))
                    preceding -= 1
                End While
                If preceding >= 0 AndAlso System.Char.IsDigit(text(preceding)) Then Return False
            End If
            If System.Char.IsDigit(text(ending - 1)) Then
                Dim following As System.Int32 = ending
                While following < text.Length AndAlso System.Char.IsWhiteSpace(text(following))
                    following += 1
                End While
                If following < text.Length AndAlso System.Char.IsDigit(text(following)) Then Return False
            End If
            ' An isolated amount must not match a prefix/suffix of a different
            ' decimal/grouped amount in the rendered document (100 vs 100,000).
            If start > 0 AndAlso System.Char.IsDigit(text(start)) AndAlso
                (text(start - 1) = "-"c OrElse text(start - 1) = "+"c OrElse text(start - 1) = Microsoft.VisualBasic.ChrW(&H2212)) Then Return False
            If start > 1 AndAlso System.Char.IsDigit(text(start)) AndAlso IsNumericSeparator(text(start - 1)) AndAlso System.Char.IsDigit(text(start - 2)) Then Return False
            If ending + 1 < text.Length AndAlso System.Char.IsDigit(text(ending - 1)) AndAlso IsNumericSeparator(text(ending)) AndAlso System.Char.IsDigit(text(ending + 1)) Then Return False
            If ending < text.Length AndAlso length > 1 AndAlso IsNumericSeparator(text(ending - 1)) AndAlso System.Char.IsDigit(text(ending - 2)) AndAlso System.Char.IsDigit(text(ending)) Then Return False
            Return True
        End Function

        Private Shared Function IsWordCharacter(character As System.Char) As System.Boolean
            Return System.Char.IsLetterOrDigit(character) OrElse character = "_"c OrElse
                System.Char.GetUnicodeCategory(character) = System.Globalization.UnicodeCategory.NonSpacingMark OrElse
                System.Char.GetUnicodeCategory(character) = System.Globalization.UnicodeCategory.SpacingCombiningMark
        End Function

        Private Shared Function IsNumericSeparator(character As System.Char) As System.Boolean
            Return character = "."c OrElse character = ","c OrElse character = "'"c OrElse character = Microsoft.VisualBasic.ChrW(&H2019)
        End Function

        ''' <summary>Only whitespace is normalized. Numbers, signs, punctuation and
        ''' case remain significant; no fuzzy/letters-only match may attest DD evidence.</summary>
        Private Shared Function NormalizeForMatch(value As System.String, offsets As System.Collections.Generic.List(Of System.Int32)) As System.String
            Dim text As System.String = If(value, System.String.Empty)
            Dim builder As New System.Text.StringBuilder(text.Length)
            Dim pendingSpace As System.Int32 = -1
            For index As System.Int32 = 0 To text.Length - 1
                Dim character As System.Char = text(index)
                If System.Char.IsWhiteSpace(character) Then
                    If builder.Length > 0 AndAlso pendingSpace < 0 Then pendingSpace = index
                    Continue For
                End If
                If pendingSpace >= 0 Then
                    builder.Append(" "c)
                    If offsets IsNot Nothing Then offsets.Add(pendingSpace)
                    pendingSpace = -1
                End If
                builder.Append(character)
                If offsets IsNot Nothing Then offsets.Add(index)
            Next
            Return builder.ToString()
        End Function

        Private Shared Function ValidateGeometry(page As PdfSharp.Pdf.PdfPage, match As NativeMatch) As System.String
            If ((page.Rotate Mod 360) + 360) Mod 360 <> 0 Then Return "Rotated-page text geometry has not been verified; no approximate native highlight was drawn."
            Dim media As PdfSharp.Pdf.PdfRectangle = page.MediaBox
            ' Some PDFsharp versions create a zero-sized CropBox when its getter is
            ' used on a page without one. Reading geometry must never blank a preview.
            Dim crop As PdfSharp.Pdf.PdfRectangle = media
            If page.Elements.ContainsKey("/CropBox") Then crop = page.CropBox
            If System.Math.Abs(media.X1) > 0.01 OrElse System.Math.Abs(media.Y1) > 0.01 Then Return "The native PDF uses an offset page origin; exact text geometry is not verified."
            If crop.Width > 0 AndAlso crop.Height > 0 AndAlso
                (System.Math.Abs(crop.X1 - media.X1) > 0.01 OrElse System.Math.Abs(crop.Y1 - media.Y1) > 0.01 OrElse
                 System.Math.Abs(crop.Width - media.Width) > 0.01 OrElse System.Math.Abs(crop.Height - media.Height) > 0.01) Then Return "The native PDF uses a cropped page; exact text geometry is not verified."
            If System.Math.Abs(page.Width.Point - match.PageWidth) > 1.0 OrElse System.Math.Abs(page.Height.Point - match.PageHeight) > 1.0 Then Return "The text and rendering engines reported different page dimensions."
            For Each box As NativeWordBox In match.Boxes
                If System.Double.IsNaN(box.X) OrElse System.Double.IsNaN(box.Y) OrElse System.Double.IsNaN(box.Width) OrElse System.Double.IsNaN(box.Height) OrElse System.Double.IsNaN(box.Baseline) OrElse
                    System.Double.IsInfinity(box.X) OrElse System.Double.IsInfinity(box.Y) OrElse System.Double.IsInfinity(box.Width) OrElse System.Double.IsInfinity(box.Height) OrElse System.Double.IsInfinity(box.Baseline) OrElse
                    box.Width <= 0 OrElse box.Height <= 0 OrElse box.X < -1.0 OrElse box.Y < -1.0 OrElse
                    box.X + box.Width > match.PageWidth + 1.0 OrElse box.Y + box.Height > match.PageHeight + 1.0 Then Return "A source highlight falls outside the verified page geometry."
            Next
            Return System.String.Empty
        End Function

        Private Shared Sub PaintHighlight(page As PdfSharp.Pdf.PdfPage, boxes As System.Collections.Generic.List(Of NativeWordBox))
            ' Embed the translucent marks in the disposable preview's content stream.
            ' WebView2 therefore displays them without relying on a PDF annotation UI,
            ' JavaScript PDF library, online viewer, or viewer-specific search command.
            Using graphics As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(page, PdfSharp.Drawing.XGraphicsPdfPageOptions.Append)
                Dim brush As New PdfSharp.Drawing.XSolidBrush(PdfSharp.Drawing.XColor.FromArgb(76, 255, 211, 48))
                ' Merge only adjacent glyphs on the same baseline. Keep line breaks
                ' and column gaps separate, and never paint a whole bounding paragraph.
                For Each box As NativeWordBox In MergeHighlightBoxes(boxes)
                    graphics.DrawRectangle(brush, box.X, page.Height.Point - box.Y - box.Height, box.Width, box.Height)
                Next
            End Using
        End Sub

        Private Shared Function MergeHighlightBoxes(boxes As System.Collections.Generic.List(Of NativeWordBox)) As System.Collections.Generic.List(Of NativeWordBox)
            Dim lines As New System.Collections.Generic.List(Of System.Collections.Generic.List(Of NativeWordBox))()
            For Each box As NativeWordBox In boxes
                Dim destination As System.Collections.Generic.List(Of NativeWordBox) = Nothing
                For Each line As System.Collections.Generic.List(Of NativeWordBox) In lines
                    If System.Math.Abs(box.Baseline - line(0).Baseline) <= System.Math.Max(0.5, System.Math.Min(box.Height, line(0).Height) * 0.2) Then
                        destination = line
                        Exit For
                    End If
                Next
                If destination Is Nothing Then
                    destination = New System.Collections.Generic.List(Of NativeWordBox)()
                    lines.Add(destination)
                End If
                destination.Add(box)
            Next
            Dim result As New System.Collections.Generic.List(Of NativeWordBox)()
            For Each line As System.Collections.Generic.List(Of NativeWordBox) In lines
                line.Sort(Function(left As NativeWordBox, right As NativeWordBox) left.X.CompareTo(right.X))
                Dim current As NativeWordBox = Nothing
                For Each box As NativeWordBox In line
                    If current IsNot Nothing AndAlso box.X <= current.X + current.Width + System.Math.Max(2.0, System.Math.Max(current.Height, box.Height) * 0.45) Then
                        Dim right As System.Double = System.Math.Max(current.X + current.Width, box.X + box.Width)
                        Dim top As System.Double = System.Math.Max(current.Y + current.Height, box.Y + box.Height)
                        current.Y = System.Math.Min(current.Y, box.Y)
                        current.Width = right - current.X
                        current.Height = top - current.Y
                    Else
                        current = New NativeWordBox With {.X = box.X, .Y = box.Y, .Width = box.Width, .Height = box.Height, .Baseline = box.Baseline}
                        result.Add(current)
                    End If
                Next
            Next
            Return result
        End Function


        Private Shared Async Function ExportOfficePdfAsync(source As System.String, extension As System.String,
                destination As System.String, cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.String)
            Dim packageIssue As System.String = ValidateOfflineOfficePackage(source, extension)
            If packageIssue.Length > 0 Then Return packageIssue & " The extracted-text viewer remains available."
            If Not OfficePreviewGate.Wait(0) Then Return "A prior Office preview conversion is still finishing. Use the extracted-text viewer now and retry the native preview later."
            Dim attempt As New OfficePreviewAttempt With {
                .WorkDirectory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "RedInkReviewOffice-" & System.Guid.NewGuid().ToString("N")),
                .Extension = extension
            }
            Dim workerStarted As System.Boolean = False
            Try
                System.IO.Directory.CreateDirectory(attempt.WorkDirectory)
                attempt.SourcePath = System.IO.Path.Combine(attempt.WorkDirectory, "source" & extension)
                attempt.OutputPath = System.IO.Path.Combine(attempt.WorkDirectory, "render.pdf")
                System.IO.File.Copy(source, attempt.SourcePath, False)
                cancellationToken.ThrowIfCancellationRequested()
                Dim worker As New System.Threading.Thread(Sub() RunOfficePreview(attempt)) With {.IsBackground = True, .Name = "Red Ink native review preview"}
                worker.SetApartmentState(System.Threading.ApartmentState.STA)
                worker.Start()
                workerStarted = True
                Using deadline As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                    Dim timeout As System.Threading.Tasks.Task = System.Threading.Tasks.Task.Delay(System.TimeSpan.FromSeconds(120), deadline.Token)
                    Dim completed As System.Threading.Tasks.Task = Await System.Threading.Tasks.Task.WhenAny(attempt.Completion.Task, timeout).ConfigureAwait(False)
                    If completed IsNot attempt.Completion.Task OrElse cancellationToken.IsCancellationRequested Then
                        System.Threading.Interlocked.Exchange(attempt.Abandoned, 1)
                        cancellationToken.ThrowIfCancellationRequested()
                        Return "Local Office did not finish the preview within 120 seconds. The extracted-text viewer remains available. A still-running conversion must finish before another can start."
                    End If
                    deadline.Cancel()
                    Dim errorMessage As System.String = Await attempt.Completion.Task.ConfigureAwait(False)
                    If Not System.String.IsNullOrEmpty(errorMessage) Then Return "Local Office could not render this source: " & errorMessage & " The extracted-text viewer remains available."
                End Using
                cancellationToken.ThrowIfCancellationRequested()
                If Not System.IO.File.Exists(attempt.OutputPath) OrElse New System.IO.FileInfo(attempt.OutputPath).Length = 0 Then Return "Local Office produced no PDF preview. The extracted-text viewer remains available."
                Using stream As New System.IO.FileStream(attempt.OutputPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                    Dim signature(4) As System.Byte
                    If stream.Read(signature, 0, signature.Length) <> signature.Length OrElse System.Text.Encoding.ASCII.GetString(signature) <> "%PDF-" Then Return "The Office preview did not contain a valid PDF signature. The extracted-text viewer remains available."
                End Using
                ' The caller only exposes a complete output, never a worker's partial file.
                PublishCopy(attempt.OutputPath, destination)
                Return System.String.Empty
            Finally
                If Not workerStarted Then OfficePreviewGate.Release()
                If Not workerStarted OrElse attempt.Completion.Task.IsCompleted Then DeleteDirectory(attempt.WorkDirectory)
            End Try
        End Function

        ''' <summary>Every COM object is created, used and released on this dedicated
        ''' STA. Never use a live workbook/document as the export carrier.</summary>
        Private Shared Sub RunOfficePreview(attempt As OfficePreviewAttempt)
            Dim application As System.Object = Nothing
            Dim collection As System.Object = Nothing
            Dim document As System.Object = Nothing
            Dim wordOptions As System.Object = Nothing
            Dim originalSettings As New System.Collections.Generic.Dictionary(Of System.String, System.Object)(System.StringComparer.Ordinal)
            Dim originalWordOptions As New System.Collections.Generic.Dictionary(Of System.String, System.Object)(System.StringComparer.Ordinal)
            Dim quitEmptyApplication As System.Boolean = False
            Dim errorMessage As System.String = System.String.Empty
            Dim applicationKind As System.String = If(attempt.Extension = ".xlsx" OrElse attempt.Extension = ".xlsm", "Excel",
                If(attempt.Extension = ".pptx" OrElse attempt.Extension = ".pptm", "PowerPoint", "Word"))
            Dim collectionName As System.String = If(applicationKind = "Excel", "Workbooks", If(applicationKind = "PowerPoint", "Presentations", "Documents"))
            Try
                Dim applicationType As System.Type = System.Type.GetTypeFromProgID(applicationKind & ".Application", True)
                application = System.Activator.CreateInstance(applicationType)
                collection = GetComProperty(application, collectionName)
                Dim originalDocumentCount As System.Int32 = System.Convert.ToInt32(GetComProperty(collection, "Count"), System.Globalization.CultureInfo.InvariantCulture)
                quitEmptyApplication = originalDocumentCount = 0 AndAlso Not System.Convert.ToBoolean(GetComProperty(application, "Visible"), System.Globalization.CultureInfo.InvariantCulture)
                ' PowerPoint may share an application even after CreateInstance. Only
                ' restore settings captured here, and never close unrelated documents.
                RememberAndSet(application, originalSettings, "AutomationSecurity", 3) ' msoAutomationSecurityForceDisable
                RememberAndSet(application, originalSettings, "DisplayAlerts", If(applicationKind = "PowerPoint", 1, 0))
                If originalDocumentCount = 0 AndAlso applicationKind <> "PowerPoint" Then RememberAndSet(application, originalSettings, "Visible", False)
                Dim unavailablePassword As System.String = System.Guid.NewGuid().ToString("N")
                Select Case applicationKind
                    Case "Word"
                        wordOptions = GetComProperty(application, "Options")
                        RememberAndSet(wordOptions, originalWordOptions, "UpdateLinksAtOpen", False)
                        RememberAndSet(wordOptions, originalWordOptions, "UpdateFieldsAtPrint", False)
                        RememberAndSet(wordOptions, originalWordOptions, "UpdateLinksAtPrint", False)
                        document = CallCom(collection, "Open",
                            New System.Object() {attempt.SourcePath, True, False, False, False, unavailablePassword, unavailablePassword, False, True},
                            New System.String() {"FileName", "ReadOnly", "Visible", "AddToRecentFiles", "ConfirmConversions", "PasswordDocument", "WritePasswordDocument", "OpenAndRepair", "NoEncodingDialog"})
                        If System.Threading.Volatile.Read(attempt.Abandoned) = 0 Then
                            CallCom(document, "ExportAsFixedFormat", New System.Object() {attempt.OutputPath, 17, False}, New System.String() {"OutputFileName", "ExportFormat", "OpenAfterExport"})
                        End If
                    Case "Excel"
                        RememberAndSet(application, originalSettings, "EnableEvents", False)
                        RememberAndSet(application, originalSettings, "AskToUpdateLinks", False)
                        RememberAndSet(application, originalSettings, "Calculation", -4135) ' xlCalculationManual; cached source values remain authoritative
                        document = CallCom(collection, "Open",
                            New System.Object() {attempt.SourcePath, 0, True, unavailablePassword, unavailablePassword, True, False, False},
                            New System.String() {"Filename", "UpdateLinks", "ReadOnly", "Password", "WriteResPassword", "IgnoreReadOnlyRecommended", "Notify", "AddToMru"})
                        If System.Threading.Volatile.Read(attempt.Abandoned) = 0 Then
                            CallCom(document, "ExportAsFixedFormat", New System.Object() {0, attempt.OutputPath, 0, False, False, False},
                                New System.String() {"Type", "Filename", "Quality", "IncludeDocProperties", "IgnorePrintAreas", "OpenAfterPublish"})
                        End If
                    Case "PowerPoint"
                        document = CallCom(collection, "Open", New System.Object() {attempt.SourcePath, -1, 0, 0},
                            New System.String() {"FileName", "ReadOnly", "Untitled", "WithWindow"})
                        If System.Threading.Volatile.Read(attempt.Abandoned) = 0 Then
                            CallCom(document, "SaveAs", New System.Object() {attempt.OutputPath, 32, -1}, New System.String() {"FileName", "FileFormat", "EmbedTrueTypeFonts"})
                        End If
                End Select
            Catch ex As System.Exception
                Dim actual As System.Exception = If(TryCast(ex, System.Reflection.TargetInvocationException)?.InnerException, ex)
                errorMessage = actual.Message
                System.Diagnostics.Trace.WriteLine("[TabularReview] Office preview failure: " & ex.ToString())
            Finally
                If document IsNot Nothing Then
                    Try
                        If applicationKind = "PowerPoint" Then
                            CallCom(document, "Close", New System.Object() {})
                        Else
                            CallCom(document, "Close", New System.Object() {False}, New System.String() {"SaveChanges"})
                        End If
                    Catch ex As System.Exception
                        System.Diagnostics.Trace.WriteLine("[TabularReview] closing owned preview document failed: " & ex.Message)
                    End Try
                    ReleaseCom(document)
                End If
                If application IsNot Nothing Then
                    For Each setting As System.Collections.Generic.KeyValuePair(Of System.String, System.Object) In originalWordOptions
                        Try
                            SetComProperty(wordOptions, setting.Key, setting.Value)
                        Catch ex As System.Exception
                            System.Diagnostics.Trace.WriteLine("[TabularReview] restoring Word preview option failed: " & setting.Key & ": " & ex.Message)
                        End Try
                    Next
                    For Each setting As System.Collections.Generic.KeyValuePair(Of System.String, System.Object) In originalSettings
                        Try
                            SetComProperty(application, setting.Key, setting.Value)
                        Catch ex As System.Exception
                            System.Diagnostics.Trace.WriteLine("[TabularReview] restoring preview application setting failed: " & setting.Key & ": " & ex.Message)
                        End Try
                    Next
                    Try
                        ' A user can open a document while conversion runs. Even an
                        ' originally empty application must then remain open.
                        If quitEmptyApplication AndAlso collection IsNot Nothing AndAlso System.Convert.ToInt32(GetComProperty(collection, "Count"), System.Globalization.CultureInfo.InvariantCulture) = 0 Then
                            If applicationKind = "Word" Then
                                CallCom(application, "Quit", New System.Object() {False}, New System.String() {"SaveChanges"})
                            Else
                                CallCom(application, "Quit", New System.Object() {})
                            End If
                        End If
                    Catch ex As System.Exception
                        System.Diagnostics.Trace.WriteLine("[TabularReview] closing empty preview application failed: " & ex.Message)
                    End Try
                End If
                ReleaseCom(wordOptions)
                ReleaseCom(collection)
                ReleaseCom(application)
                OfficePreviewGate.Release()
                attempt.Completion.TrySetResult(errorMessage)
                If System.Threading.Volatile.Read(attempt.Abandoned) <> 0 Then DeleteDirectory(attempt.WorkDirectory)
            End Try
        End Sub

        Private Shared Function ValidateOfflineOfficePackage(path As System.String, extension As System.String) As System.String
            If extension = ".doc" OrElse extension = ".rtf" Then Return System.String.Empty
            ' Native conversion must not fetch a different live resource and then
            ' present it as evidence from the hashed source package. Ordinary
            ' hyperlinks are inert here; externally linked pictures, templates,
            ' OLE items and data refresh connections require the text fallback.
            Using archive As System.IO.Compression.ZipArchive = System.IO.Compression.ZipFile.OpenRead(path)
                For Each entry As System.IO.Compression.ZipArchiveEntry In archive.Entries
                    Dim name As System.String = entry.FullName.Replace("\", "/")
                    If name.StartsWith("xl/macrosheets/", System.StringComparison.OrdinalIgnoreCase) Then Return "The workbook contains Excel 4 macro sheets. Automatic native rendering is disabled for this source."
                    If System.String.Equals(name, "xl/connections.xml", System.StringComparison.OrdinalIgnoreCase) OrElse
                        name.StartsWith("xl/queryTables/", System.StringComparison.OrdinalIgnoreCase) Then Return "The workbook contains external data connections or query tables. Native rendering is disabled to prevent refreshed data from being presented as evidence from the stored source."
                    If Not name.EndsWith(".rels", System.StringComparison.OrdinalIgnoreCase) Then Continue For
                    Using stream As System.IO.Stream = entry.Open()
                        Dim settings As New System.Xml.XmlReaderSettings With {.DtdProcessing = System.Xml.DtdProcessing.Prohibit, .XmlResolver = Nothing}
                        Using reader As System.Xml.XmlReader = System.Xml.XmlReader.Create(stream, settings)
                            While reader.Read()
                                If reader.NodeType <> System.Xml.XmlNodeType.Element OrElse reader.LocalName <> "Relationship" Then Continue While
                                If Not System.String.Equals(reader.GetAttribute("TargetMode"), "External", System.StringComparison.OrdinalIgnoreCase) Then Continue While
                                Dim relationType As System.String = If(reader.GetAttribute("Type"), System.String.Empty)
                                If Not relationType.EndsWith("/hyperlink", System.StringComparison.OrdinalIgnoreCase) Then Return "The Office file includes externally linked content. Native rendering is disabled so that no unpinned external resource can enter the source preview."
                            End While
                        End Using
                    End Using
                Next
            End Using
            Return System.String.Empty
        End Function

        Private Shared Sub PublishCopy(source As System.String, destination As System.String)
            Dim temporary As System.String = destination & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
            Try
                System.IO.File.Copy(source, temporary, False)
                If Not System.IO.File.Exists(destination) Then System.IO.File.Move(temporary, destination)
            Finally
                DeleteFile(temporary)
            End Try
        End Sub

        Private Shared Sub RememberAndSet(target As System.Object, values As System.Collections.Generic.Dictionary(Of System.String, System.Object),
                                          name As System.String, value As System.Object)
            If Not values.ContainsKey(name) Then values.Add(name, GetComProperty(target, name))
            SetComProperty(target, name, value)
        End Sub

        Private Shared Function GetComProperty(target As System.Object, name As System.String) As System.Object
            Return target.GetType().InvokeMember(name,
                System.Reflection.BindingFlags.GetProperty Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance,
                Nothing, target, Nothing, System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Shared Sub SetComProperty(target As System.Object, name As System.String, value As System.Object)
            target.GetType().InvokeMember(name,
                System.Reflection.BindingFlags.SetProperty Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance,
                Nothing, target, New System.Object() {value}, System.Globalization.CultureInfo.InvariantCulture)
        End Sub

        Private Shared Function CallCom(target As System.Object, name As System.String, values As System.Object(),
                                       Optional names As System.String() = Nothing) As System.Object
            Return target.GetType().InvokeMember(name,
                System.Reflection.BindingFlags.InvokeMethod Or System.Reflection.BindingFlags.OptionalParamBinding Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance,
                Nothing, target, values, Nothing, System.Globalization.CultureInfo.InvariantCulture, names)
        End Function

        Private Shared Sub ReleaseCom(value As System.Object)
            Try
                If value IsNot Nothing AndAlso System.Runtime.InteropServices.Marshal.IsComObject(value) Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine("[TabularReview] preview COM release failed: " & ex.Message)
            End Try
        End Sub

        Private Shared Sub DeleteFile(path As System.String)
            Try
                If Not System.String.IsNullOrEmpty(path) AndAlso System.IO.File.Exists(path) Then System.IO.File.Delete(path)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine("[TabularReview] preview snapshot cleanup failed: " & ex.Message)
            End Try
        End Sub

        Private Shared Sub DeleteDirectory(path As System.String)
            Try
                If Not System.String.IsNullOrEmpty(path) AndAlso System.IO.Directory.Exists(path) Then System.IO.Directory.Delete(path, True)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine("[TabularReview] preview worker cleanup failed: " & ex.Message)
            End Try
        End Sub
    End Class
End Namespace
