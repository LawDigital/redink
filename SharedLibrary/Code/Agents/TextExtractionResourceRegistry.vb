' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' Session-scoped immutable extraction-resource registry. The registry is transport-,
' model-, provider-, organization- and document-type agnostic. Adapter-specific code
' supplies a configuration fingerprint and performs the actual extraction.
Option Strict On
Option Explicit On
Option Infer On

Namespace Agents

    Public NotInheritable Class TextExtractionProcessedRange
        Public Property StartPage As System.Int32
        Public Property EndPage As System.Int32
        Public Property Association As System.String = "range_only"
    End Class

    Public NotInheritable Class TextExtractionPayload
        Public Property Success As System.Boolean
        Public Property Content As System.String = System.String.Empty
        Public Property ErrorCode As System.String = System.String.Empty
        Public Property Message As System.String = System.String.Empty
        Public Property PageCount As System.Nullable(Of System.Int32) = Nothing
        Public Property OcrUsed As System.Nullable(Of System.Boolean) = Nothing
        Public Property OcrAttempted As System.Nullable(Of System.Boolean) = Nothing
        Public Property OcrSkipped As System.Nullable(Of System.Boolean) = Nothing
        Public Property OcrDurationMilliseconds As System.Nullable(Of System.Int64) = Nothing
        Public Property ExtractionComplete As System.Nullable(Of System.Boolean) = Nothing
        Public Property ContentFormat As System.String = "unknown"
        Public Property ProcessedRanges As New System.Collections.Generic.List(Of TextExtractionProcessedRange)()
    End Class

    Public NotInheritable Class TextExtractionResource
        Public Property ResourceId As System.String = System.String.Empty
        Public Property SourceIdentifier As System.String = System.String.Empty
        Public Property SourceSha256 As System.String = System.String.Empty
        Public Property AdapterId As System.String = System.String.Empty
        Public Property AdapterVersion As System.String = System.String.Empty
        Public Property ConfigurationFingerprint As System.String = System.String.Empty
        Public Property OptionsFingerprint As System.String = System.String.Empty
        Public Property CreatedUtc As System.DateTime
        Public Property ReuseStatus As System.String = "created"
        Public Property Payload As TextExtractionPayload
    End Class

    Public NotInheritable Class TextExtractionResourceRegistry

        Private NotInheritable Class CapturedSource
            Public Property OriginalPath As System.String = System.String.Empty
            Public Property SnapshotPath As System.String = System.String.Empty
            Public Property Sha256 As System.String = System.String.Empty
        End Class

        Private Shared ReadOnly SessionId As System.String = System.Guid.NewGuid().ToString("N")
        Private Shared ReadOnly Cache As New System.Collections.Concurrent.ConcurrentDictionary(Of System.String, TextExtractionResource)(System.StringComparer.Ordinal)
        Private Shared ReadOnly Flights As New System.Collections.Concurrent.ConcurrentDictionary(Of System.String, System.Lazy(Of System.Threading.Tasks.Task(Of TextExtractionResource)))(System.StringComparer.Ordinal)
        Private Shared ReadOnly PublishedOutputs As New System.Collections.Concurrent.ConcurrentDictionary(Of System.String, PublishedOutputAssociation)(System.StringComparer.OrdinalIgnoreCase)

        Private NotInheritable Class PublishedOutputAssociation
            Public Property Resource As TextExtractionResource
            Public Property SnapshotSha256 As System.String = System.String.Empty
        End Class

        Private Sub New()
        End Sub

        Public Shared Async Function ResolveAsync(
                source As System.String,
                workingDirectory As System.String,
                adapterId As System.String,
                adapterVersion As System.String,
                configurationFingerprint As System.String,
                optionsFingerprint As System.String,
                extractor As System.Func(Of System.String, System.Threading.CancellationToken, System.Threading.Tasks.Task(Of TextExtractionPayload)),
                cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of TextExtractionResource)

            If extractor Is Nothing Then Throw New System.ArgumentNullException(NameOf(extractor))
            Dim captured As CapturedSource = CaptureSource(source, workingDirectory)
            Dim key As System.String = BuildKey(captured.Sha256,
                                                System.IO.Path.GetExtension(captured.OriginalPath),
                                                System.IO.Path.GetFullPath(PathPolicy.GetDefaultWritableRoot()),
                                                adapterId,
                                                adapterVersion,
                                                configurationFingerprint,
                                                optionsFingerprint)
            Try
                Dim cached As TextExtractionResource = Nothing
                If Cache.TryGetValue(key, cached) AndAlso cached IsNot Nothing AndAlso cached.Payload IsNot Nothing AndAlso cached.Payload.Success Then
                    DeleteCapturedSource(captured)
                    Return CloneWithReuse(cached, "session_cache_hit")
                End If

                Dim createdLazy As New System.Lazy(Of System.Threading.Tasks.Task(Of TextExtractionResource))(
                    Function() CreateResourceAsync(key, captured, adapterId, adapterVersion, configurationFingerprint, optionsFingerprint, extractor, cancellationToken),
                    System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)
                Dim activeLazy As System.Lazy(Of System.Threading.Tasks.Task(Of TextExtractionResource)) = Flights.GetOrAdd(key, createdLazy)
                Dim ownsFlight As System.Boolean = System.Object.ReferenceEquals(activeLazy, createdLazy)

                If Not ownsFlight Then
                    DeleteCapturedSource(captured)
                End If

                Dim resource As TextExtractionResource
                Try
                    resource = Await activeLazy.Value.ConfigureAwait(False)
                Finally
                    If ownsFlight Then
                        Dim ignored As System.Lazy(Of System.Threading.Tasks.Task(Of TextExtractionResource)) = Nothing
                        Flights.TryRemove(key, ignored)
                    End If
                End Try

                If resource Is Nothing Then Return Nothing
                Return CloneWithReuse(resource, If(ownsFlight, resource.ReuseStatus, "singleflight_join"))
            Catch
                DeleteCapturedSource(captured)
                Throw
            End Try
        End Function

        Private Shared Async Function CreateResourceAsync(
                key As System.String,
                captured As CapturedSource,
                adapterId As System.String,
                adapterVersion As System.String,
                configurationFingerprint As System.String,
                optionsFingerprint As System.String,
                extractor As System.Func(Of System.String, System.Threading.CancellationToken, System.Threading.Tasks.Task(Of TextExtractionPayload)),
                cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of TextExtractionResource)
            Try
                cancellationToken.ThrowIfCancellationRequested()
                Dim payload As TextExtractionPayload = Await extractor(captured.SnapshotPath, cancellationToken).ConfigureAwait(False)
                cancellationToken.ThrowIfCancellationRequested()

                Dim resource As New TextExtractionResource With {
                    .ResourceId = "txr-" & key.Substring(0, 24),
                    .SourceIdentifier = captured.OriginalPath,
                    .SourceSha256 = captured.Sha256,
                    .AdapterId = If(adapterId, System.String.Empty),
                    .AdapterVersion = If(adapterVersion, System.String.Empty),
                    .ConfigurationFingerprint = If(configurationFingerprint, System.String.Empty),
                    .OptionsFingerprint = If(optionsFingerprint, System.String.Empty),
                    .CreatedUtc = System.DateTime.UtcNow,
                    .ReuseStatus = "created",
                    .Payload = payload
                }

                If payload IsNot Nothing AndAlso payload.Success Then
                    Cache(key) = resource
                    CleanupSessionCache()
                End If
                Return resource
            Finally
                DeleteCapturedSource(captured)
            End Try
        End Function

        Private Shared Function CaptureSource(source As System.String, workingDirectory As System.String) As CapturedSource
            Dim sourcePath As System.String = PathPolicy.Resolve(source, PathAccess.Read)
            If Not System.IO.File.Exists(sourcePath) Then
                Throw New TextFileInputException("not_found", "Extraction source file was not found.")
            End If
            Dim directory As System.String = PathPolicy.Resolve(workingDirectory, PathAccess.Write)
            System.IO.Directory.CreateDirectory(directory)
            Dim extension As System.String = System.IO.Path.GetExtension(sourcePath)
            Dim snapshotPath As System.String = PathPolicy.Resolve(
                System.IO.Path.Combine(directory, ".ri-source-" & System.Guid.NewGuid().ToString("N") & extension), PathAccess.Write)

            Dim hash As System.String
            Try
                Using input As New System.IO.FileStream(sourcePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                    Using output As New System.IO.FileStream(snapshotPath, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None)
                        input.CopyTo(output)
                        output.Flush(True)
                    End Using
                End Using
                Using snapshotStream As New System.IO.FileStream(snapshotPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                    Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                        hash = System.BitConverter.ToString(algorithm.ComputeHash(snapshotStream)).Replace("-", System.String.Empty).ToLowerInvariant()
                    End Using
                End Using
                Return New CapturedSource With {.OriginalPath = sourcePath, .SnapshotPath = snapshotPath, .Sha256 = hash}
            Catch
                Try
                    If System.IO.File.Exists(snapshotPath) Then System.IO.File.Delete(snapshotPath)
                Catch
                End Try
                Throw
            End Try
        End Function

        Private Shared Sub DeleteCapturedSource(captured As CapturedSource)
            If captured Is Nothing OrElse System.String.IsNullOrWhiteSpace(captured.SnapshotPath) Then Return
            Try
                If System.IO.File.Exists(captured.SnapshotPath) Then System.IO.File.Delete(captured.SnapshotPath)
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Extraction source snapshot cleanup failed: " & ex.GetType().FullName)
            End Try
        End Sub

        Private Shared Function BuildKey(sourceHash As System.String,
                                         sourceExtension As System.String,
                                         sessionScope As System.String,
                                         adapterId As System.String,
                                         adapterVersion As System.String,
                                         configurationFingerprint As System.String,
                                         optionsFingerprint As System.String) As System.String
            Dim canonical As System.String = String.Join("|", New System.String() {
                "v1", SessionId, If(sessionScope, System.String.Empty).ToLowerInvariant(),
                If(sourceHash, System.String.Empty), If(sourceExtension, System.String.Empty).ToLowerInvariant(),
                If(adapterId, System.String.Empty), If(adapterVersion, System.String.Empty),
                If(configurationFingerprint, System.String.Empty), If(optionsFingerprint, System.String.Empty)
            })
            Return HashString(canonical)
        End Function

        Private Shared Sub CleanupSessionCache()
            Dim cutoff As System.DateTime = System.DateTime.UtcNow.AddHours(-4)
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, TextExtractionResource) In Cache
                If pair.Value Is Nothing OrElse pair.Value.CreatedUtc < cutoff Then
                    Dim ignored As TextExtractionResource = Nothing
                    Cache.TryRemove(pair.Key, ignored)
                End If
            Next
            If Cache.Count <= 64 Then Return
            Dim ordered = System.Linq.Enumerable.ToArray(
                System.Linq.Enumerable.Take(
                    System.Linq.Enumerable.OrderBy(Cache, Function(entry) If(entry.Value Is Nothing, System.DateTime.MinValue, entry.Value.CreatedUtc)),
                    Cache.Count - 64))
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, TextExtractionResource) In ordered
                Dim ignored As TextExtractionResource = Nothing
                Cache.TryRemove(pair.Key, ignored)
            Next
        End Sub

        Public Shared Function HashString(value As System.String) As System.String
            Dim bytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(If(value, System.String.Empty))
            Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Return System.BitConverter.ToString(algorithm.ComputeHash(bytes)).Replace("-", System.String.Empty).ToLowerInvariant()
            End Using
        End Function

        Private Shared Function CloneWithReuse(source As TextExtractionResource, reuseStatus As System.String) As TextExtractionResource
            If source Is Nothing Then Return Nothing
            Return New TextExtractionResource With {
                .ResourceId = source.ResourceId,
                .SourceIdentifier = source.SourceIdentifier,
                .SourceSha256 = source.SourceSha256,
                .AdapterId = source.AdapterId,
                .AdapterVersion = source.AdapterVersion,
                .ConfigurationFingerprint = source.ConfigurationFingerprint,
                .OptionsFingerprint = source.OptionsFingerprint,
                .CreatedUtc = source.CreatedUtc,
                .ReuseStatus = reuseStatus,
                .Payload = source.Payload
            }
        End Function


        Public Shared Function HasPublishedOutput(outputPath As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(outputPath) OrElse Not System.IO.File.Exists(outputPath) Then Return False
            Try
                Dim path As System.String = PathPolicy.Resolve(outputPath, PathAccess.Read)
                Return PublishedOutputs.ContainsKey(path)
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Public Shared Sub RegisterPublishedOutput(outputPath As System.String,
                                                  resource As TextExtractionResource,
                                                  snapshotSha256 As System.String)
            If resource Is Nothing OrElse resource.Payload Is Nothing OrElse Not resource.Payload.Success Then Return
            If System.String.IsNullOrWhiteSpace(outputPath) OrElse System.String.IsNullOrWhiteSpace(snapshotSha256) Then Return
            Dim path As System.String = PathPolicy.Resolve(outputPath, PathAccess.Read)
            PublishedOutputs(path) = New PublishedOutputAssociation With {
                .Resource = resource,
                .SnapshotSha256 = snapshotSha256
            }
        End Sub

        Public Shared Function TryGetVerifiedPublishedOutput(outputPath As System.String,
                                                             sourceSha256 As System.String,
                                                             adapterId As System.String,
                                                             adapterVersion As System.String,
                                                             configurationFingerprint As System.String,
                                                             optionsFingerprint As System.String,
                                                             ByRef resource As TextExtractionResource) As System.Boolean
            resource = Nothing
            If System.String.IsNullOrWhiteSpace(outputPath) OrElse Not System.IO.File.Exists(outputPath) Then Return False
            Dim path As System.String = PathPolicy.Resolve(outputPath, PathAccess.Read)
            Dim association As PublishedOutputAssociation = Nothing
            If Not PublishedOutputs.TryGetValue(path, association) OrElse association Is Nothing OrElse association.Resource Is Nothing Then Return False
            Dim candidate As TextExtractionResource = association.Resource
            If Not System.String.Equals(candidate.SourceSha256, sourceSha256, System.StringComparison.Ordinal) OrElse
               Not System.String.Equals(candidate.AdapterId, adapterId, System.StringComparison.Ordinal) OrElse
               Not System.String.Equals(candidate.AdapterVersion, adapterVersion, System.StringComparison.Ordinal) OrElse
               Not System.String.Equals(candidate.ConfigurationFingerprint, configurationFingerprint, System.StringComparison.Ordinal) OrElse
               Not System.String.Equals(candidate.OptionsFingerprint, optionsFingerprint, System.StringComparison.Ordinal) Then
                Return False
            End If
            Try
                Dim snapshot As TextFileSnapshot = TextFileSnapshot.Read(path)
                If Not System.String.Equals(snapshot.Sha256, association.SnapshotSha256, System.StringComparison.Ordinal) Then Return False
            Catch ex As System.Exception
                Return False
            End Try
            resource = CloneWithReuse(candidate, "published_output_reuse")
            Return True
        End Function

    End Class

End Namespace
