' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.


' =============================================================================
' File: SharedMethods.SemanticArchive.Navigation.vb
' Purpose:
'   Writes validated archive navigation indexes and references to generated/source
'   artifacts.
'
' Architecture / Function:
'   Uses the existing archive/path/permission contracts rather than creating a separate
'   search or extraction format.
' =============================================================================

Option Strict On
Option Explicit On
Option Infer On

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Serializes existing child metadata as a normal version-1 self-indexed text.
        ''' Complete card boundaries and local byte spans are assigned by the host; an
        ''' external document/node reference never becomes an indexed-file byte offset.
        ''' </summary>
        Friend Shared Async Function WriteSemanticArchiveNavigationIndexAsync(
            node As SemanticArchiveNode,
            outputPath As String,
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of SemanticSearchIndexGenerationResult)

            If node Is Nothing Then Throw New System.ArgumentNullException(NameOf(node))
            If System.IO.File.Exists(outputPath) Then Throw New System.IO.IOException("Immutable archive navigation already exists.")
            cancellationToken.ThrowIfCancellationRequested()
            Dim body As New System.Text.StringBuilder()
            Dim index As New SemanticSearchIndexDocument() With {
                .FormatVersion = SemanticSearchCurrentFormatVersion,
                .Encoding = "utf-8",
                .OffsetUnit = "byte",
                .OffsetBase = "content",
                .CreatedUtc = System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture),
                .GeneratorVersion = SemanticSearchDefaultGeneratorVersion,
                .MetadataProfile = SemanticSearchMetadataProfile.Generic.ToString()
            }
            Dim byteOffset As Long = 0
            Dim localDocumentId As String = "D" & node.NodeId
            For Each card As SemanticArchiveCard In node.Cards
                cancellationToken.ThrowIfCancellationRequested()
                Dim text As String = SemanticArchiveMetadata.RenderCard(card) & vbLf
                card.RetrievalText = text.TrimEnd(ChrW(10))
                card.StartByte = byteOffset
                card.LengthBytes = SemanticSearchUtf8NoBom.GetByteCount(text)
                Dim entry As SemanticSearchIndexEntry = SemanticArchiveMetadata.Clone(If(card.Metadata, New SemanticSearchIndexEntry()))
                entry.Id = "S" & (index.Entries.Count + 1).ToString("0000", System.Globalization.CultureInfo.InvariantCulture)
                entry.StableId = card.CardId
                entry.Order = index.Entries.Count + 1
                entry.Title = If(System.String.IsNullOrWhiteSpace(card.Title), card.Level, card.Title)
                entry.Summary = If(System.String.IsNullOrWhiteSpace(card.Summary), "Navigation metadata for this " & card.Level.ToLowerInvariant() & ".", card.Summary)
                entry.StartByte = byteOffset
                entry.LengthBytes = card.LengthBytes
                entry.PreviousId = If(index.Entries.Count = 0, Nothing, index.Entries(index.Entries.Count - 1).Id)
                entry.NextId = Nothing
                entry.RelatedIds = New System.Collections.Generic.List(Of String)()
                entry.SourceDocuments = New System.Collections.Generic.List(Of String) From {"Archive navigation " & node.NodeId}
                entry.SourceDocumentKeys = New System.Collections.Generic.List(Of String) From {localDocumentId}
                entry.SourceDocumentAttributes = New System.Collections.Generic.List(Of String)()
                entry.DocumentSpans = New System.Collections.Generic.List(Of SemanticSearchDocumentSpan) From {
                    New SemanticSearchDocumentSpan() With {
                        .DocumentId = localDocumentId,
                        .DocumentName = "Archive navigation " & node.NodeId,
                        .StartByte = byteOffset,
                        .LengthBytes = card.LengthBytes,
                        .StartByteInDocument = byteOffset
                    }
                }
                If index.Entries.Count > 0 Then index.Entries(index.Entries.Count - 1).NextId = entry.Id
                index.Entries.Add(entry)
                body.Append(text)
                byteOffset += card.LengthBytes
            Next
            Dim payload As Byte() = SemanticSearchUtf8NoBom.GetBytes(body.ToString())
            index.ContentSha256 = ComputeSemanticSearchSha256Hex(payload)
            index.Documents.Add(New SemanticSearchDocumentDescriptor() With {
                .DocumentId = localDocumentId,
                .StableId = node.NodeId,
                .Name = "Archive navigation " & node.NodeId,
                .WrapperStartByte = 0,
                .WrapperLengthBytes = payload.LongLength,
                .StartByte = 0,
                .LengthBytes = payload.LongLength,
                .ContentSha256 = index.ContentSha256
            })
            index.DocumentCount = index.Documents.Count
            index.SegmentCount = index.Entries.Count
            ValidateGeneratedSemanticSearchIndex(index, payload.LongLength)
            Dim header As Byte() = SemanticSearchUtf8NoBom.GetBytes(
                SemanticSearchIndexStartMarker & vbLf & SerializeSemanticSearchJson(index) & vbLf & SemanticSearchContentStartMarker & vbLf)
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(outputPath))
            Dim temporaryPath As String = outputPath & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
            Try
                Using stream As System.IO.FileStream = SemanticArchiveStore.CreatePrivateFile(temporaryPath)
                    Await stream.WriteAsync(header, 0, header.Length, cancellationToken).ConfigureAwait(False)
                    Await stream.WriteAsync(payload, 0, payload.Length, cancellationToken).ConfigureAwait(False)
                    Await stream.FlushAsync(cancellationToken).ConfigureAwait(False)
                    stream.Flush(True)
                End Using
                ValidateWrittenSemanticSearchFile(temporaryPath, header, payload.LongLength, index.ContentSha256)
                cancellationToken.ThrowIfCancellationRequested()
                ' Generations are immutable and unique; this never overwrites a live file.
                System.IO.File.Move(temporaryPath, outputPath)
            Finally
                If System.IO.File.Exists(temporaryPath) Then System.IO.File.Delete(temporaryPath)
            End Try
            node.IndexPath = outputPath
            Return New SemanticSearchIndexGenerationResult() With {
                .OutputPath = outputPath, .ContentByteLength = payload.LongLength,
                .DocumentCount = index.DocumentCount, .SegmentCount = index.SegmentCount,
                .ContentSha256 = index.ContentSha256, .IndexDocument = index
            }
        End Function
    End Class
End Namespace
