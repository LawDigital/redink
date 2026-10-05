' Part of "Red Ink" (SharedLibrary)
' Optional explicit literal inspection, independent of metadata/semantic eligibility.

' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: SemanticArchiveSearchService.Literal.vb
' Purpose:
'   Bounded literal-original-text channel over validated UTF-8 streams.
'
' Architecture / Function:
'   Verifies source/representation association and coverage while keeping exact literal
'   matching separate from semantic ranking.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Friend NotInheritable Class SemanticArchiveLiteralScanState
        Friend Property LiteralText As System.String
        Friend Property ArchivePosition As System.Int32
        Friend Property Enumerators As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.IEnumerator(Of SemanticArchiveDocumentRecord))(System.StringComparer.Ordinal)
        Friend Property PendingDocument As SemanticArchiveDocumentRecord
        Friend Property PendingArchiveId As System.String
        Friend Property Completed As System.Boolean
        Friend Property HadOmissions As System.Boolean
        Friend Property TotalByteBudgetUsed As System.Int64
    End Class

    Partial Public NotInheritable Class SemanticArchiveSearchService
        Private Sub RunLiteralTextChannel(scope As SemanticArchiveRunScope, state As SemanticArchiveSearchState,
                         coverage As SemanticArchiveCoverage, timer As System.Diagnostics.Stopwatch,
                         cancellationToken As System.Threading.CancellationToken)
            Dim scan As SemanticArchiveLiteralScanState = state.LiteralScan
            If scan Is Nothing OrElse scan.Completed Then Return
            Dim literalBytes As System.Byte() = New System.Text.UTF8Encoding(False, True).GetBytes(scan.LiteralText)
            Dim candidateCapacity As System.Int32 = System.Math.Max(1, state.Budgets.MaxCandidateFiles \ 3)
            While Not scan.Completed AndAlso coverage.LiteralDocumentsInspected < state.Budgets.MaxLiteralScanDocuments AndAlso
                  state.Candidates.Count < candidateCapacity AndAlso timer.ElapsedMilliseconds < CLng(state.Budgets.MaxElapsedSeconds) * 1000L \ 3L
                cancellationToken.ThrowIfCancellationRequested()
                If scan.PendingDocument Is Nothing Then
                    If scan.ArchivePosition >= state.ArchiveIds.Count Then
                        scan.Completed = True
                        Exit While
                    End If
                    Dim archiveId As System.String = state.ArchiveIds(scan.ArchivePosition)
                    Dim iterator As System.Collections.Generic.IEnumerator(Of SemanticArchiveDocumentRecord) = scan.Enumerators(archiveId)
                    Dim moved As System.Boolean
                    Try
                        moved = iterator.MoveNext()
                    Catch ex As System.Exception
                        scan.HadOmissions = True
                        coverage.Diagnostics.Add("literal_metadata_unavailable: A document shard could not be inspected for the explicit literal request.")
                        moved = False
                    End Try
                    If Not moved Then
                        iterator.Dispose()
                        scan.ArchivePosition += 1
                        Continue While
                    End If
                    scan.PendingArchiveId = archiveId
                    scan.PendingDocument = iterator.Current
                End If
                Dim document As SemanticArchiveDocumentRecord = scan.PendingDocument
                Dim generation As SemanticArchiveGenerationManifest = scope.State.Generations(scan.PendingArchiveId)
                coverage.LiteralDocumentsInspected += 1
                If document Is Nothing OrElse document.Representation Is Nothing OrElse
                   Not _store.CanReadDocument(scope.AccessContext, generation, document) Then
                    coverage.SourcesUnavailable += 1
                    scan.HadOmissions = True
                    scan.PendingDocument = Nothing
                    Continue While
                End If
                cancellationToken.ThrowIfCancellationRequested()
                If Not System.String.Equals(document.Representation.Completeness, "complete", System.StringComparison.OrdinalIgnoreCase) Then
                    scan.HadOmissions = True
                    If Not coverage.Diagnostics.Contains("literal_incomplete_extraction: An inspected source has partial or unknown extraction completeness; a missing literal does not rule out omitted original content.") Then
                        coverage.Diagnostics.Add("literal_incomplete_extraction: An inspected source has partial or unknown extraction completeness; a missing literal does not rule out omitted original content.")
                    End If
                End If

                ' Validation is deliberately included in this budget. ValidateTextPath
                ' hashes the original once and the exported file once. The same-handle
                ' scan below hashes the exported file again while matching it. A tiny
                ' scan budget must never trigger an unbounded full-file validation.
                Dim originalBytes As System.Int64 = document.Fingerprint.Length
                Dim textBytes As System.Int64 = document.Representation.TextByteLength
                If originalBytes < 0 OrElse textBytes < 0 OrElse textBytes > (System.Int64.MaxValue - originalBytes) \ 2L Then
                    Throw New System.IO.InvalidDataException("Invalid literal inspection byte accounting.")
                End If
                Dim validationAndScanBytes As System.Int64 = originalBytes + 2L * textBytes
                If validationAndScanBytes > state.Budgets.MaxLiteralScanBytes Then
                    coverage.LiteralDocumentsSkippedByBudget += 1
                    scan.HadOmissions = True
                    scan.PendingDocument = Nothing
                    Continue While
                End If
                If coverage.LiteralByteBudgetUsed + validationAndScanBytes > state.Budgets.MaxLiteralScanBytes Then
                    ' Keep this exact immutable record. A fresh continuation budget can
                    ' finish it; no partial match or partially validated file is accepted.
                    Exit While
                End If
                coverage.LiteralByteBudgetUsed += validationAndScanBytes
                scan.TotalByteBudgetUsed += validationAndScanBytes
                Try
                    Dim encoding As System.String = document.Representation.EncodingName
                    If Not System.String.Equals(encoding, "utf-8", System.StringComparison.OrdinalIgnoreCase) AndAlso
                       Not System.String.Equals(encoding, "utf8", System.StringComparison.OrdinalIgnoreCase) AndAlso
                       Not System.String.Equals(encoding, "utf-8-bom", System.StringComparison.OrdinalIgnoreCase) Then
                        Throw New System.IO.InvalidDataException("The literal inspection representation is not UTF-8.")
                    End If
                    Dim path As System.String = _store.ValidateTextPath(generation, document, scope.AccessContext)
                    Dim matchStart As System.Nullable(Of System.Int64)
                    Using stream As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(System.IO.Path.GetDirectoryName(path), path)
                        If stream.Length <> textBytes Then Throw New System.IO.InvalidDataException("The immutable literal-search artifact length changed.")
                        matchStart = FindLiteralInValidatedUtf8Stream(stream, literalBytes, document.Representation.TextFileHash,
                            System.String.Equals(encoding, "utf-8-bom", System.StringComparison.OrdinalIgnoreCase), cancellationToken)
                    End Using
                    If Not _store.CanReadDocument(scope.AccessContext, generation, document) Then Throw New System.UnauthorizedAccessException("The source became unavailable during literal inspection.")
                    cancellationToken.ThrowIfCancellationRequested()
                    If matchStart.HasValue Then
                        Dim identityKey As System.String = Identity(scan.PendingArchiveId, document.DocumentId)
                        ' A later exact discovery may refine a file already returned by
                        ' semantic metadata. Publish a fresh hit with its exact position.
                        state.Returned.Remove(identityKey)
                        If Not AddCandidate(state, scan.PendingArchiveId, document.DocumentId, 1.0R,
                            "The explicitly requested case-sensitive literal occurs in the verified extracted text.", "literal_text", matchStart) Then
                            Throw New System.InvalidOperationException("The reserved literal candidate capacity was exhausted.")
                        End If
                    End If
                    scan.PendingDocument = Nothing
                Catch ex As System.OperationCanceledException
                    ' PendingDocument intentionally survives a cancellation. A partial
                    ' stream has neither a validated match nor a completed cursor.
                    Throw
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("SA literal inspection failed: " & ex.GetType().FullName)
                    scan.HadOmissions = True
                    coverage.SourcesUnavailable += 1
                    coverage.Diagnostics.Add("literal_source_unavailable: One source or immutable text could not be validated; its literal result was omitted.")
                    scan.PendingDocument = Nothing
                End Try
            End While
            If coverage.LiteralDocumentsSkippedByBudget > 0 Then coverage.Diagnostics.Add("literal_validation_budget: Oversized sources were skipped before hashing/reading their content because original plus artifact validation could not fit the configured literal byte budget.")
            If scan.HadOmissions Then coverage.Diagnostics.Add("literal_coverage_incomplete: This literal scan has retained validation, access, extraction, or size omissions. It is not an exhaustive negative result.")
            If Not scan.Completed Then coverage.Diagnostics.Add("literal_scan_bounded: Continue the same query/literal/scope to inspect more authorized sources within a fresh bounded scan budget.")
        End Sub

        ''' <summary>
        ''' Bounded-memory KMP matching and strict incremental UTF-8 validation. Only a
        ''' completely hashed stream can produce a match. The hash includes a preamble;
        ''' returned match offsets exclude that preamble, matching extracted-text content.
        ''' </summary>
        Friend Shared Function FindLiteralInValidatedUtf8Stream(stream As System.IO.Stream, literal As System.Byte(), expectedHash As System.String,
                         requireUtf8Bom As System.Boolean, cancellationToken As System.Threading.CancellationToken) As System.Nullable(Of System.Int64)
            If stream Is Nothing OrElse Not stream.CanRead Then Throw New System.ArgumentException("A readable immutable stream is required.", NameOf(stream))
            If literal Is Nothing OrElse literal.Length = 0 OrElse literal.Length > 16384 Then Throw New System.ArgumentException("A bounded nonempty UTF-8 literal is required.", NameOf(literal))
            Dim prefix(literal.Length - 1) As System.Int32
            Dim matched As System.Int32 = 0
            For position As System.Int32 = 1 To literal.Length - 1
                While matched > 0 AndAlso literal(position) <> literal(matched)
                    matched = prefix(matched - 1)
                End While
                If literal(position) = literal(matched) Then matched += 1
                prefix(position) = matched
            Next
            matched = 0
            Dim buffer(65535) As System.Byte
            Dim characters(65537) As System.Char
            Dim decoder As System.Text.Decoder = New System.Text.UTF8Encoding(False, True).GetDecoder()
            Dim bytePosition As System.Int64 = 0
            Dim preambleLength As System.Int32 = 0
            Dim first As System.Boolean = True
            Dim matchStart As System.Nullable(Of System.Int64) = Nothing
            Using hash As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                While True
                    cancellationToken.ThrowIfCancellationRequested()
                    Dim count As System.Int32 = stream.Read(buffer, 0, buffer.Length)
                    If count = 0 Then Exit While
                    If first Then
                        While count < 3
                            cancellationToken.ThrowIfCancellationRequested()
                            Dim more As System.Int32 = stream.Read(buffer, count, 3 - count)
                            If more = 0 Then Exit While
                            count += more
                        End While
                    End If
                    hash.TransformBlock(buffer, 0, count, buffer, 0)
                    decoder.GetChars(buffer, 0, count, characters, 0, False)
                    Dim begin As System.Int32 = 0
                    If first Then
                        first = False
                        If count >= 3 AndAlso buffer(0) = &HEF AndAlso buffer(1) = &HBB AndAlso buffer(2) = &HBF Then
                            preambleLength = 3
                            begin = 3
                        ElseIf requireUtf8Bom Then
                            Throw New System.IO.InvalidDataException("The representation declares a UTF-8 BOM but does not contain it.")
                        End If
                    End If
                    For position As System.Int32 = begin To count - 1
                        While matched > 0 AndAlso buffer(position) <> literal(matched)
                            matched = prefix(matched - 1)
                        End While
                        If buffer(position) = literal(matched) Then matched += 1
                        If matched = literal.Length Then
                            If Not matchStart.HasValue Then matchStart = bytePosition + position - literal.Length + 1L - preambleLength
                            matched = prefix(matched - 1)
                        End If
                    Next
                    bytePosition += count
                End While
                decoder.GetChars(New System.Byte() {}, 0, 0, characters, 0, True)
                If first AndAlso requireUtf8Bom Then Throw New System.IO.InvalidDataException("The representation declares a missing UTF-8 BOM.")
                hash.TransformFinalBlock(New System.Byte() {}, 0, 0)
                Dim actual As System.String = System.BitConverter.ToString(hash.Hash).Replace("-", "").ToLowerInvariant()
                If Not System.String.Equals(actual, expectedHash, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("The literal-search artifact does not match its immutable representation hash.")
            End Using
            cancellationToken.ThrowIfCancellationRequested()
            Return matchStart
        End Function
    End Class
End Namespace
