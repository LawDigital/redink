' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
Option Strict On
Option Explicit On

Namespace SharedLibrary

    Public NotInheritable Class SemanticArchiveCooperativeManifest
        Public Property Format As System.String = "redink-semantic-artifacts"
        Public Property SchemaVersion As System.Int32 = 1
        Public Property SourceIdentity As System.String = ""
        Public Property SourcePermissionSignature As System.String = ""
        Public Property Revision As System.Int64
        Public Property WriterClaimId As System.String = ""
        Public Property Entries As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveCooperativeEntry)(System.StringComparer.Ordinal)
    End Class

    Public NotInheritable Class SemanticArchiveCooperativeEntry
        Public Property Key As System.String = ""
        Public Property Stage As System.String = ""
        Public Property SourceHash As System.String = ""
        Public Property ExtractionSignature As System.String = ""
        Public Property SemanticSignature As System.String = ""
        Public Property RecordPath As System.String = ""
        Public Property RecordHash As System.String = ""
        Public Property RecordLength As System.Int64
        Public Property PublishedUtc As System.DateTimeOffset = System.DateTimeOffset.UtcNow
    End Class

    Public NotInheritable Class SemanticArchiveCooperativeImport
        Public Property Stage As System.String = ""
        Public Property Document As SemanticArchiveDocumentRecord
    End Class

    ''' <summary>
    ''' A per-source cooperative cache, separate from every user's private aggregate.
    ''' Immutable payloads and record JSON are committed before one small manifest switch.
    ''' The shared file handle serializes all profiles for one source, so compatible
    ''' semantic variants can also reuse an extraction completed by another producer.
    ''' No elapsed-time lease takeover is permitted. Source writers are trusted producers;
    ''' source hashes bind versions, not the semantic truth of a contributor's extraction.
    ''' </summary>
    Public NotInheritable Class SemanticArchiveCooperativeStore
        Implements System.IDisposable

        Public Const ExtractedStage As System.String = "extracted"
        Public Const IndexedStage As System.String = "indexed"
        Public Const CompleteStage As System.String = "complete"
        Private Const MaximumManifestBytes As System.Int32 = 1024 * 1024
        Private Const MaximumRecordBytes As System.Int32 = 16 * 1024 * 1024
        Private Const MaximumEntries As System.Int32 = 256
        Private ReadOnly _location As SemanticArchiveArtifactLocation
        Private ReadOnly _claimId As System.String = System.Guid.NewGuid().ToString("N")
        Private ReadOnly _payloads As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
        Private _claim As System.IO.FileStream

        Public Sub New(location As SemanticArchiveArtifactLocation)
            If location Is Nothing Then Throw New System.ArgumentNullException(NameOf(location))
            _location = location
        End Sub

        Public ReadOnly Property HasClaim As System.Boolean
            Get
                Return _claim IsNot Nothing
            End Get
        End Property

        Public Shared Function BuildKey(stage As System.String, sourceHash As System.String,
                                        extractionSignature As System.String, semanticSignature As System.String) As System.String
            If stage <> ExtractedStage AndAlso stage <> IndexedStage AndAlso stage <> CompleteStage Then Throw New System.ArgumentException("Unknown cooperative artifact stage.", NameOf(stage))
            If Not IsHash(sourceHash) OrElse System.String.IsNullOrWhiteSpace(extractionSignature) Then Throw New System.ArgumentException("A source hash and extraction signature are required.")
            If stage <> ExtractedStage AndAlso System.String.IsNullOrWhiteSpace(semanticSignature) Then Throw New System.ArgumentException("A semantic stage requires its exact configuration signature.")
            ' Length-delimited JSON avoids separator collisions in externally supplied profiles.
            Return SemanticArchiveIdentity.StableId("artifact", Newtonsoft.Json.JsonConvert.SerializeObject(
                New System.String() {stage, sourceHash.ToLowerInvariant(), extractionSignature, If(stage = ExtractedStage, "", semanticSignature)}))
        End Function

        Public Sub AcquireClaim(cancellationToken As System.Threading.CancellationToken, Optional maximumWait As System.TimeSpan? = Nothing)
            If Not _location.IsShared Then Throw New System.InvalidOperationException("A private placement has no cooperative writer claim.")
            If _claim IsNot Nothing Then Return
            Dim budget As System.TimeSpan = If(maximumWait.HasValue, maximumWait.Value, System.TimeSpan.FromSeconds(30))
            If budget < System.TimeSpan.Zero OrElse budget > System.TimeSpan.FromSeconds(30) Then Throw New System.ArgumentOutOfRangeException(NameOf(maximumWait))
            SemanticArchiveArtifactPlanner.CreateArtifactDirectory(_location, _location.ArtifactDirectory)
            Dim claimPath As System.String = System.IO.Path.Combine(_location.ArtifactDirectory, "claim.lock")
            Dim clock As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Do
                cancellationToken.ThrowIfCancellationRequested()
                Dim opened As System.IO.FileStream = Nothing
                Try
                    opened = SemanticArchiveArtifactPlanner.OpenClaim(_location, claimPath)
                    _claim = opened
                    AssertClaim()
                    Return
                Catch ex As System.IO.IOException When (ex.HResult And &HFFFF) = 32 OrElse (ex.HResult And &HFFFF) = 33
                    If opened IsNot Nothing Then opened.Dispose()
                    _claim = Nothing
                    If clock.Elapsed >= budget Then Throw New SemanticArchiveCooperativeBusyException("Another contributor is processing this source; its shared claim was not replaced.", ex)
                    cancellationToken.WaitHandle.WaitOne(System.Math.Min(100, System.Math.Max(1, CInt((budget - clock.Elapsed).TotalMilliseconds))))
                Catch
                    If opened IsNot Nothing Then opened.Dispose()
                    _claim = Nothing
                    Throw
                End Try
            Loop
        End Sub

        ''' <summary>Read-only reuse does not require acquiring or writing a shared claim.</summary>
        Public Function TryRead(sourceHash As System.String, extractionSignature As System.String, semanticSignature As System.String,
                                allowSemantic As System.Boolean, cancellationToken As System.Threading.CancellationToken) As SemanticArchiveCooperativeImport
            If Not _location.IsShared Then Return Nothing
            Dim manifest As SemanticArchiveCooperativeManifest = ReadManifest()
            Dim stages As System.String() = If(allowSemantic, New System.String() {CompleteStage, IndexedStage, ExtractedStage}, New System.String() {ExtractedStage})
            For Each stage As System.String In stages
                cancellationToken.ThrowIfCancellationRequested()
                Dim key As System.String = BuildKey(stage, sourceHash, extractionSignature, semanticSignature)
                Dim descriptor As SemanticArchiveCooperativeEntry = Nothing
                If Not manifest.Entries.TryGetValue(key, descriptor) Then Continue For
                ValidateEntry(descriptor, key, stage, sourceHash, extractionSignature, semanticSignature)
                Dim recordPath As System.String = ResolveRelative(descriptor.RecordPath)
                Dim record As SemanticArchiveDocumentRecord = ReadJson(Of SemanticArchiveDocumentRecord)(recordPath, MaximumRecordBytes, descriptor.RecordHash, descriptor.RecordLength)
                If record Is Nothing OrElse record.CanonicalSourceKey <> _location.SourceIdentity Then Throw New System.IO.InvalidDataException("The shared document belongs to a different physical source.")
                ValidateRecord(record, stage, sourceHash, extractionSignature, semanticSignature)
                Dim textRelative As System.String = record.Representation.TextPath
                Dim textPath As System.String = ResolveRelative(textRelative)
                RequireHash(textPath, record.Representation.TextFileHash, record.Representation.TextByteLength)
                _payloads(record.Representation.TextFileHash.ToLowerInvariant()) = textRelative
                record.Representation.TextPath = textPath
                If record.Index IsNot Nothing Then
                    Dim indexRelative As System.String = record.Index.Path
                    Dim indexPath As System.String = ResolveRelative(indexRelative)
                    RequireHash(indexPath, record.Index.FileHash, -1)
                    _payloads(record.Index.FileHash.ToLowerInvariant()) = indexRelative
                    record.Index.Path = indexPath
                End If
                Return New SemanticArchiveCooperativeImport With {.Stage = stage, .Document = record}
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Copy an accepted shared record into a private immutable generation location.
        ''' Aggregate readers consequently never rely on a contributor's mutable pathname.
        ''' Paths and card identity are rebound by the builder to its independently readable
        ''' original source after this copy; shared record objects are never mutated in place.
        ''' </summary>
        Public Function Materialize(imported As SemanticArchiveCooperativeImport, privateVersionDirectory As System.String,
                                    cancellationToken As System.Threading.CancellationToken) As SemanticArchiveDocumentRecord
            If imported Is Nothing OrElse imported.Document Is Nothing Then Throw New System.ArgumentNullException(NameOf(imported))
            Dim document As SemanticArchiveDocumentRecord = SemanticArchiveMetadata.Clone(imported.Document)
            SemanticArchiveStore.CreatePrivateDirectory(privateVersionDirectory)
            Dim textPath As System.String = System.IO.Path.Combine(privateVersionDirectory, PayloadName(".txt"))
            CopySharedToPrivate(document.Representation.TextPath, textPath, document.Representation.TextFileHash, cancellationToken)
            document.Representation.TextPath = textPath
            If document.Index IsNot Nothing Then
                Dim indexPath As System.String = System.IO.Path.Combine(privateVersionDirectory, PayloadName(".indexed.txt"))
                CopySharedToPrivate(document.Index.Path, indexPath, document.Index.FileHash, cancellationToken)
                document.Index.Path = indexPath
            End If
            Return document
        End Function

        Public Sub Publish(document As SemanticArchiveDocumentRecord, stage As System.String, semanticSignature As System.String,
                           cancellationToken As System.Threading.CancellationToken)
            If _claim Is Nothing Then Throw New System.InvalidOperationException("A cooperative source claim is required before publication.")
            AssertClaim()
            cancellationToken.ThrowIfCancellationRequested()
            If document Is Nothing OrElse document.Representation Is Nothing Then Throw New System.ArgumentException("No reusable extraction is present.", NameOf(document))
            Dim record As SemanticArchiveDocumentRecord = SemanticArchiveMetadata.Clone(document)
            record.SemanticSignature = If(stage = ExtractedStage, "", semanticSignature)
            If stage = ExtractedStage Then record.Index = Nothing
            If stage <> CompleteStage Then record.Card = Nothing
            ValidateRecord(record, stage, record.Fingerprint.Sha256, record.ExtractionSignature, semanticSignature)
            Dim manifest As SemanticArchiveCooperativeManifest = ReadManifest()
            Dim previousRevision As System.Int64 = manifest.Revision
            Dim previousWriter As System.String = manifest.WriterClaimId
            Dim key As System.String = BuildKey(stage, record.Fingerprint.Sha256, record.ExtractionSignature, semanticSignature)
            If Not manifest.Entries.ContainsKey(key) AndAlso manifest.Entries.Count >= MaximumEntries Then
                Dim oldestKey As System.String = Nothing
                Dim oldestTime As System.DateTimeOffset = System.DateTimeOffset.MaxValue
                For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, SemanticArchiveCooperativeEntry) In manifest.Entries
                    If pair.Value IsNot Nothing AndAlso pair.Value.SourceHash <> record.Fingerprint.Sha256 AndAlso pair.Value.PublishedUtc <= oldestTime Then
                        oldestKey = pair.Key
                        oldestTime = pair.Value.PublishedUtc
                    End If
                Next
                If oldestKey Is Nothing Then Throw New System.IO.InvalidDataException("The bounded shared source manifest already holds the maximum current-version processing variants; private artifacts remain usable.")
                ' Prune only an old source-version lookup, never its immutable files.
                ' Pinned readers and already imported private generations stay valid.
                manifest.Entries.Remove(oldestKey)
            End If
            Dim version As System.String = "versions/" & System.Guid.NewGuid().ToString("N") & "/"
            Dim versionPath As System.String = ResolveRelative(version.TrimEnd("/"c), False)
            SemanticArchiveArtifactPlanner.CreateArtifactDirectory(_location, versionPath)
            record.Representation.TextPath = PublishPayload(record.Representation.TextPath, record.Representation.TextFileHash,
                version & PayloadName(".txt"), cancellationToken)
            If record.Index IsNot Nothing Then
                record.Index.Path = PublishPayload(record.Index.Path, record.Index.FileHash, version & PayloadName(".indexed.txt"), cancellationToken)
            End If
            ' Only this source's content is contributed, never the user's bindings,
            ' partition membership, personal artifact paths or aggregate summaries.
            record.DocumentId = ""
            record.SourceItemId = ""
            record.CanonicalSourceKey = _location.SourceIdentity
            record.SourcePath = ""
            record.RelativePath = System.IO.Path.GetFileName(_location.SourcePath)
            record.PartitionKey = ""
            record.BindingIds = New System.Collections.Generic.List(Of System.String)()
            record.Diagnostic = ""
            record.CooperativeState = ""
            record.CooperativeDiagnostic = ""
            If record.Card IsNot Nothing Then
                record.Card.CardId = ""
                record.Card.TargetId = ""
                record.Card.PartitionKey = ""
                record.Card.FullMetadataReference = ""
                record.Card.RetrievalText = SemanticArchiveMetadata.RenderCard(record.Card)
            End If
            Dim recordRelative As System.String = version & "record.json"
            Dim recordPath As System.String = ResolveRelative(recordRelative, False)
            Dim encodedLength As System.Int32 = System.Text.Encoding.UTF8.GetByteCount(Newtonsoft.Json.JsonConvert.SerializeObject(record))
            If encodedLength > MaximumRecordBytes Then Throw New System.IO.InvalidDataException("The shared document metadata exceeds its record budget.")
            SemanticArchiveArtifactPlanner.AtomicWriteJson(_location, recordPath, record)
            Dim descriptor As New SemanticArchiveCooperativeEntry With {
                .Key = key, .Stage = stage, .SourceHash = record.Fingerprint.Sha256,
                .ExtractionSignature = record.ExtractionSignature, .SemanticSignature = If(stage = ExtractedStage, "", semanticSignature),
                .RecordPath = recordRelative, .RecordHash = SemanticArchiveIdentity.ComputeFileHash(recordPath),
                .RecordLength = New System.IO.FileInfo(recordPath).Length}
            manifest.Entries(key) = descriptor
            If manifest.Revision = System.Int64.MaxValue Then Throw New System.IO.InvalidDataException("The shared source manifest revision is exhausted.")
            manifest.Revision += 1
            manifest.WriterClaimId = _claimId
            cancellationToken.ThrowIfCancellationRequested()
            AssertClaim()
            Dim beforeCommit As SemanticArchiveCooperativeManifest = ReadManifest()
            If beforeCommit.Revision <> previousRevision OrElse beforeCommit.WriterClaimId <> previousWriter Then Throw New System.IO.IOException("The shared source manifest changed outside this producer's claim; publication was rejected.")
            ' This is the only commit point. Previously committed immutable artifacts
            ' remain available if any copy, cancellation or manifest switch fails.
            SemanticArchiveArtifactPlanner.AtomicWriteJson(_location, _location.ManifestPath, manifest)
            Dim committed As SemanticArchiveCooperativeManifest = ReadManifest()
            If committed.Revision <> manifest.Revision OrElse committed.WriterClaimId <> _claimId OrElse
                Not committed.Entries.ContainsKey(key) OrElse committed.Entries(key).RecordHash <> descriptor.RecordHash Then
                Throw New System.IO.IOException("Shared artifact publication could not be verified.")
            End If
        End Sub

        Private Sub AssertClaim()
            If _claim Is Nothing Then Throw New System.InvalidOperationException("The cooperative source claim is not held.")
            ' Exercise the still-open exclusive handle before each commit. A lost SMB
            ' handle must fail here; a stale in-memory IDisposable is not a lease.
            Dim bytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(_claimId)
            _claim.Position = 0
            _claim.Write(bytes, 0, bytes.Length)
            _claim.SetLength(bytes.Length)
            _claim.Flush(True)
        End Sub

        ''' <summary>
        ''' Called only by the permission reconciler after every generated artifact has
        ''' been checked/repaired under this same exclusive source claim. Rebinds the
        ''' access-policy commit record last; it never regenerates text or metadata.
        ''' </summary>
        Friend Sub RebindPermissionsAfterRepair(exclusiveClaim As System.IO.FileStream,
                                                cancellationToken As System.Threading.CancellationToken)
            If exclusiveClaim Is Nothing OrElse Not exclusiveClaim.CanWrite OrElse Not exclusiveClaim.CanSeek Then Throw New System.InvalidOperationException("Permission rebinding requires the live exclusive source claim.")
            Dim expectedClaim As System.String = SemanticArchivePathGuard.CanonicalPath(System.IO.Path.Combine(_location.ArtifactDirectory, "claim.lock"))
            If Not System.String.Equals(SemanticArchivePathGuard.CanonicalPath(exclusiveClaim.Name), expectedClaim, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.InvalidOperationException("The permission-repair claim belongs to a different source.")
            cancellationToken.ThrowIfCancellationRequested()
            If Not ExistsChecked(_location.ManifestPath) Then Return
            Dim manifest As SemanticArchiveCooperativeManifest = ReadManifest(True)
            If manifest.SourcePermissionSignature = _location.SourcePermissionSignature Then Return
            Dim oldRevision As System.Int64 = manifest.Revision
            Dim oldWriter As System.String = manifest.WriterClaimId
            Dim oldPermission As System.String = manifest.SourcePermissionSignature
            If oldRevision = System.Int64.MaxValue Then Throw New System.IO.InvalidDataException("The shared source manifest revision is exhausted.")
            Dim claimBytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(_claimId)
            exclusiveClaim.Position = 0
            exclusiveClaim.Write(claimBytes, 0, claimBytes.Length)
            exclusiveClaim.SetLength(claimBytes.Length)
            exclusiveClaim.Flush(True)
            Dim beforeCommit As SemanticArchiveCooperativeManifest = ReadManifest(True)
            If beforeCommit.Revision <> oldRevision OrElse beforeCommit.WriterClaimId <> oldWriter OrElse beforeCommit.SourcePermissionSignature <> oldPermission Then Throw New System.IO.IOException("The shared manifest changed during permission repair.")
            manifest.Revision += 1
            manifest.WriterClaimId = _claimId
            manifest.SourcePermissionSignature = _location.SourcePermissionSignature
            cancellationToken.ThrowIfCancellationRequested()
            SemanticArchiveArtifactPlanner.AtomicWriteJson(_location, _location.ManifestPath, manifest)
            Dim committed As SemanticArchiveCooperativeManifest = ReadManifest()
            If committed.Revision <> manifest.Revision OrElse committed.WriterClaimId <> _claimId Then Throw New System.IO.IOException("The repaired shared permission manifest could not be verified.")
        End Sub

        Private Function ReadManifest(Optional allowPreviousPermissionSignature As System.Boolean = False) As SemanticArchiveCooperativeManifest
            Dim manifest As SemanticArchiveCooperativeManifest
            If Not ExistsChecked(_location.ManifestPath) Then
                manifest = New SemanticArchiveCooperativeManifest With {.SourceIdentity = _location.SourceIdentity, .SourcePermissionSignature = _location.SourcePermissionSignature}
            Else
                SemanticArchiveArtifactPlanner.RequireArtifact(_location, _location.ManifestPath)
                manifest = ReadJson(Of SemanticArchiveCooperativeManifest)(_location.ManifestPath, MaximumManifestBytes)
            End If
            If manifest Is Nothing OrElse manifest.Format <> "redink-semantic-artifacts" OrElse manifest.SchemaVersion <> 1 OrElse
                manifest.SourceIdentity <> _location.SourceIdentity OrElse Not IsHash(manifest.SourcePermissionSignature) OrElse
                (Not allowPreviousPermissionSignature AndAlso manifest.SourcePermissionSignature <> _location.SourcePermissionSignature) OrElse
                manifest.Revision < 0 OrElse manifest.Entries Is Nothing OrElse manifest.Entries.Count > MaximumEntries Then
                Throw New System.IO.InvalidDataException("The shared manifest identity, access policy or schema is not valid for this source.")
            End If
            Return manifest
        End Function

        Private Shared Sub ValidateEntry(entry As SemanticArchiveCooperativeEntry, key As System.String, stage As System.String,
                                         sourceHash As System.String, extractionSignature As System.String, semanticSignature As System.String)
            If entry Is Nothing OrElse entry.Key <> key OrElse entry.Stage <> stage OrElse
                Not System.String.Equals(entry.SourceHash, sourceHash, System.StringComparison.OrdinalIgnoreCase) OrElse entry.ExtractionSignature <> extractionSignature OrElse
                entry.SemanticSignature <> If(stage = ExtractedStage, "", semanticSignature) OrElse Not IsHash(entry.RecordHash) OrElse
                entry.RecordLength < 1 OrElse entry.RecordLength > MaximumRecordBytes Then Throw New System.IO.InvalidDataException("A shared processing variant has inconsistent provenance.")
        End Sub

        Private Shared Sub ValidateRecord(record As SemanticArchiveDocumentRecord, stage As System.String, sourceHash As System.String,
                                          extractionSignature As System.String, semanticSignature As System.String)
            If record Is Nothing OrElse record.Fingerprint Is Nothing OrElse record.Representation Is Nothing OrElse
                Not System.String.Equals(record.Fingerprint.Sha256, sourceHash, System.StringComparison.OrdinalIgnoreCase) OrElse
                Not System.String.Equals(record.Representation.SourceHash, sourceHash, System.StringComparison.OrdinalIgnoreCase) OrElse
                record.ExtractionSignature <> extractionSignature OrElse record.Representation.OptionsSignature <> extractionSignature OrElse
                Not IsHash(record.Representation.TextFileHash) OrElse record.Representation.TextByteLength < 1 OrElse
                System.String.IsNullOrWhiteSpace(record.Representation.RepresentationId) OrElse System.String.IsNullOrWhiteSpace(record.Representation.ExtractorVersion) OrElse
                (record.Representation.Completeness <> "complete" AndAlso record.Representation.Completeness <> "incomplete" AndAlso record.Representation.Completeness <> "unknown") Then
                Throw New System.IO.InvalidDataException("The shared extraction does not match the source and processing signatures.")
            End If
            If stage <> ExtractedStage AndAlso record.SemanticSignature <> semanticSignature Then Throw New System.IO.InvalidDataException("The shared semantic record has a different processing signature.")
            If record.Index IsNot Nothing AndAlso (record.Index.RepresentationId <> record.Representation.RepresentationId OrElse
                record.Index.ModelIdentity <> semanticSignature OrElse Not IsHash(record.Index.FileHash) OrElse Not IsHash(record.Index.PayloadHash) OrElse
                record.Index.FormatVersion <> SharedMethods.SemanticSearchCurrentFormatVersion OrElse record.Index.GeneratorVersion <> SharedMethods.SemanticSearchDefaultGeneratorVersion) Then
                Throw New System.IO.InvalidDataException("The shared section index has incompatible provenance.")
            End If
            If stage = IndexedStage AndAlso record.Index Is Nothing Then Throw New System.IO.InvalidDataException("An indexed checkpoint has no section index.")
            If stage = CompleteStage AndAlso (record.Card Is Nothing OrElse record.Card.Metadata Is Nothing OrElse record.Card.Level <> "DOCUMENT" OrElse
                record.Card.SourceVersion <> sourceHash OrElse record.Card.RepresentationId <> record.Representation.RepresentationId) Then
                Throw New System.IO.InvalidDataException("A completed shared document has no matching document card.")
            End If
        End Sub

        Private Function PublishPayload(sourcePath As System.String, hash As System.String, relativePath As System.String,
                                        cancellationToken As System.Threading.CancellationToken) As System.String
            Dim existing As System.String = Nothing
            If _payloads.TryGetValue(hash.ToLowerInvariant(), existing) Then
                RequireHash(ResolveRelative(existing), hash, -1)
                Return existing
            End If
            SemanticArchiveStore.RequirePrivateArtifact(sourcePath)
            Dim target As System.String = ResolveRelative(relativePath, False)
            Using input As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(System.IO.Path.GetDirectoryName(sourcePath), sourcePath)
                Using output As System.IO.FileStream = SemanticArchiveArtifactPlanner.CreateArtifactFile(_location, target)
                    CopyBytes(input, output, cancellationToken)
                End Using
            End Using
            RequireHash(target, hash, -1)
            _payloads(hash.ToLowerInvariant()) = relativePath
            Return relativePath
        End Function

        Private Sub CopySharedToPrivate(source As System.String, target As System.String, hash As System.String,
                                        cancellationToken As System.Threading.CancellationToken)
            RequireHash(source, hash, -1)
            Using input As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(_location.ArtifactDirectory, source)
                Using output As System.IO.FileStream = SemanticArchiveStore.CreatePrivateFile(target)
                    CopyBytes(input, output, cancellationToken)
                End Using
            End Using
            SemanticArchiveStore.RequirePrivateArtifact(target)
            If Not System.String.Equals(SemanticArchiveIdentity.ComputeFileHash(target), hash, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("A shared payload changed during private import.")
        End Sub

        Private Shared Sub CopyBytes(input As System.IO.Stream, output As System.IO.FileStream, cancellationToken As System.Threading.CancellationToken)
            Dim buffer(81919) As System.Byte
            Do
                cancellationToken.ThrowIfCancellationRequested()
                Dim read As System.Int32 = input.Read(buffer, 0, buffer.Length)
                If read = 0 Then Exit Do
                output.Write(buffer, 0, read)
            Loop
            output.Flush(True)
        End Sub

        Private Function ResolveRelative(relativePath As System.String, Optional mustExist As System.Boolean = True) As System.String
            If System.String.IsNullOrWhiteSpace(relativePath) OrElse System.IO.Path.IsPathRooted(relativePath) OrElse relativePath.IndexOf(":"c) >= 0 Then Throw New System.IO.InvalidDataException("A shared artifact must use a relative immutable path.")
            Return SemanticArchivePathGuard.ValidateContainedPath(_location.ArtifactDirectory,
                System.IO.Path.Combine(_location.ArtifactDirectory, relativePath.Replace("/"c, System.IO.Path.DirectorySeparatorChar)), mustExist)
        End Function

        Private Sub RequireHash(path As System.String, hash As System.String, length As System.Int64)
            If Not IsHash(hash) Then Throw New System.IO.InvalidDataException("A shared artifact hash is missing.")
            SemanticArchiveArtifactPlanner.RequireArtifact(_location, path)
            If length >= 0 AndAlso New System.IO.FileInfo(path).Length <> length Then Throw New System.IO.InvalidDataException("A shared artifact length changed.")
            Using stream As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(_location.ArtifactDirectory, path)
                Using hasher As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                    Dim actual As System.String = System.BitConverter.ToString(hasher.ComputeHash(stream)).Replace("-", "").ToLowerInvariant()
                    If Not System.String.Equals(actual, hash, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("A shared artifact hash changed.")
                End Using
            End Using
        End Sub

        Private Function ReadJson(Of T As Class)(path As System.String, maximumBytes As System.Int32,
                                                 Optional expectedHash As System.String = Nothing, Optional expectedLength As System.Int64 = -1) As T
            SemanticArchiveArtifactPlanner.RequireArtifact(_location, path)
            Dim bytes As System.Byte()
            Using input As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(_location.ArtifactDirectory, path)
                If input.Length < 1 OrElse input.Length > maximumBytes OrElse (expectedLength >= 0 AndAlso input.Length <> expectedLength) Then Throw New System.IO.InvalidDataException("The shared JSON record exceeds its budget or differs from its committed length.")
                Using buffer As New System.IO.MemoryStream()
                    Dim block(8191) As System.Byte
                    Do
                        Dim count As System.Int32 = input.Read(block, 0, block.Length)
                        If count = 0 Then Exit Do
                        If buffer.Length + count > maximumBytes Then Throw New System.IO.InvalidDataException("The shared JSON record grew beyond its bounded size.")
                        buffer.Write(block, 0, count)
                    Loop
                    bytes = buffer.ToArray()
                End Using
            End Using
            If expectedLength >= 0 AndAlso bytes.LongLength <> expectedLength Then Throw New System.IO.InvalidDataException("The shared JSON record differs from its committed length.")
            If expectedHash IsNot Nothing AndAlso Not System.String.Equals(SemanticArchiveIdentity.HashBytes(bytes), expectedHash, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("The shared JSON bytes differ from their committed hash.")
            Dim json As System.String = New System.Text.UTF8Encoding(False, True).GetString(bytes)
            If json.Length > 0 AndAlso json(0) = System.Convert.ToChar(&HFEFF) Then json = json.Substring(1)
            Return Newtonsoft.Json.JsonConvert.DeserializeObject(Of T)(json,
                New Newtonsoft.Json.JsonSerializerSettings With {.TypeNameHandling = Newtonsoft.Json.TypeNameHandling.None,
                    .MetadataPropertyHandling = Newtonsoft.Json.MetadataPropertyHandling.Ignore, .MaxDepth = 80, .CheckAdditionalContent = True})
        End Function

        Private Shared Function ExistsChecked(path As System.String) As System.Boolean
            Try
                Dim attributes As System.IO.FileAttributes = System.IO.File.GetAttributes(path)
                If (attributes And System.IO.FileAttributes.Directory) <> 0 Then Throw New System.IO.InvalidDataException("A shared manifest path is occupied by a directory.")
                Return True
            Catch ex As System.IO.FileNotFoundException
                Return False
            Catch ex As System.IO.DirectoryNotFoundException
                Return False
            End Try
        End Function

        Private Function PayloadName(suffix As System.String) As System.String
            ' Source identity and version are already encoded by the containing slot.
            ' Original full names remain in document metadata; bounded payload names
            ' leave room for long Windows user-profile and configured parent paths.
            Return "content" & suffix
        End Function

        Private Shared Function IsHash(value As System.String) As System.Boolean
            If value Is Nothing OrElse value.Length <> 64 Then Return False
            For Each character As System.Char In value
                If Not ((character >= "0"c AndAlso character <= "9"c) OrElse (character >= "a"c AndAlso character <= "f"c) OrElse (character >= "A"c AndAlso character <= "F"c)) Then Return False
            Next
            Return True
        End Function

        Public Sub Dispose() Implements System.IDisposable.Dispose
            If _claim IsNot Nothing Then
                _claim.Dispose()
                _claim = Nothing
            End If
        End Sub
    End Class

    Public NotInheritable Class SemanticArchiveCooperativeBusyException
        Inherits System.IO.IOException
        Public Sub New(message As System.String, innerException As System.Exception)
            MyBase.New(message, innerException)
        End Sub
    End Class
End Namespace
