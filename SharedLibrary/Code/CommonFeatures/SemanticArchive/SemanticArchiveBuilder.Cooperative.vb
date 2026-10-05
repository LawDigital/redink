' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Partial Class SemanticArchiveBuilder
        Private NotInheritable Class CooperativePreparation
            Public Property Store As SemanticArchiveCooperativeStore
            Public Property Document As SemanticArchiveDocumentRecord
            Public Property ValidatedLocalExtract As System.Boolean
        End Class

        Private Async Function PrepareCooperativeAsync(archive As SemanticArchiveDefinition, binding As SemanticArchiveSourceBinding,
                                                        item As SemanticArchiveWorkItem, document As SemanticArchiveDocumentRecord,
                                                        queue As SemanticArchiveWorkQueue,
                                                        options As SemanticArchiveBuildOptions, result As SemanticArchiveBuildResult,
                                                        cancellationToken As System.Threading.CancellationToken,
                                                        Optional validatedLocalExtract As System.Boolean = False) As System.Threading.Tasks.Task(Of CooperativePreparation)
            Dim preparation As New CooperativePreparation With {.Document = document, .ValidatedLocalExtract = validatedLocalExtract}
            Dim indexOnly As System.Boolean = options.IndexOnlyRebuild OrElse item.IndexOnlyRebuild
            If indexOnly AndAlso Not item.ForceExtractionRebuild AndAlso (validatedLocalExtract OrElse ValidateReusableRepresentation(binding, document.Representation, item.SourceHash, item.ExtractionSignature)) Then
                ' Never replace a valid local extract or wait for a shared writer merely
                ' to regenerate private semantic derivatives.
                preparation.ValidatedLocalExtract = True
                Return preparation
            End If
            Try
                Dim location As SemanticArchiveArtifactLocation = SemanticArchiveArtifactPlanner.Plan(binding, item.SourcePath)
                If Not System.String.Equals(location.SourceIdentity, item.CanonicalSourceKey, System.StringComparison.Ordinal) Then
                    Throw New System.InvalidOperationException("source_identity_changed: The planned source identity differs from the physically verified source job. Refresh discovery before processing the remapped original.")
                End If
                Dim imported As SemanticArchiveCooperativeImport = Nothing
                If Not item.ForceExtractionRebuild Then
                    ' Preferred and legacy locations are read-only candidates. An old
                    ' source-parent cache remains reusable even when the new namespace
                    ' is not writable. Only the current preferred Plan can be claimed.
                    For Each readLocation As SemanticArchiveArtifactLocation In SemanticArchiveArtifactPlanner.GetReadLocations(binding, item.SourcePath)
                        cancellationToken.ThrowIfCancellationRequested()
                        If Not System.String.Equals(readLocation.SourceIdentity, item.CanonicalSourceKey, System.StringComparison.Ordinal) Then
                            Throw New System.InvalidOperationException("source_identity_changed: A cache candidate belongs to a different physically verified source.")
                        End If
                        Try
                            Using reader As New SemanticArchiveCooperativeStore(readLocation)
                                imported = FilterCooperativeImport(reader.TryRead(item.SourceHash, item.ExtractionSignature, item.SemanticSignature,
                                    Not item.ForceSemanticRebuild, cancellationToken), item.SourceLength, archive.SectionIndexThresholdBytes)
                                If imported IsNot Nothing Then
                                    preparation.Document = Await ImportIfUsefulAsync(reader, imported, archive, binding, item,
                                        preparation.Document, result, cancellationToken).ConfigureAwait(False)
                                    CheckpointCooperativeImport(preparation.Document, item, document.Fingerprint, queue)
                                    If imported.Stage = SemanticArchiveCooperativeStore.CompleteStage AndAlso Not item.ForceSemanticRebuild Then Return preparation
                                End If
                            End Using
                        Catch ex As System.Exception When IsCooperativeAvailabilityFailure(ex)
                            result.Diagnostics.Add("cooperative_candidate_unavailable: " & item.DocumentId & "; " & ex.Message)
                        End Try
                    Next
                End If
                If indexOnly Then Return preparation
                If Not location.IsShared Then
                    preparation.Document.CooperativeState = "local_only"
                    preparation.Document.CooperativeDiagnostic = location.Diagnostic
                    result.Diagnostics.Add("cooperative_local_only: " & item.DocumentId & "; " & location.Diagnostic)
                    Return preparation
                End If
                preparation.Store = New SemanticArchiveCooperativeStore(location)
                preparation.Store.AcquireClaim(cancellationToken, If(options.IsBackground, System.TimeSpan.FromMilliseconds(250), System.TimeSpan.FromSeconds(30)))
                ' A producer may have completed while this process waited for its claim.
                ' Recheck after acquisition before invoking any extractor or model.
                If Not item.ForceExtractionRebuild Then
                    imported = FilterCooperativeImport(preparation.Store.TryRead(item.SourceHash, item.ExtractionSignature, item.SemanticSignature,
                        Not item.ForceSemanticRebuild, cancellationToken), item.SourceLength, archive.SectionIndexThresholdBytes)
                    If imported IsNot Nothing Then
                        preparation.Document = Await ImportIfUsefulAsync(preparation.Store, imported, archive, binding, item,
                            preparation.Document, result, cancellationToken).ConfigureAwait(False)
                        CheckpointCooperativeImport(preparation.Document, item, document.Fingerprint, queue)
                        If imported.Stage = SemanticArchiveCooperativeStore.CompleteStage AndAlso Not item.ForceSemanticRebuild Then preparation.Store.Dispose()
                    End If
                End If
                Return preparation
            Catch ex As SemanticArchiveCooperativeBusyException
                If preparation.Store IsNot Nothing Then preparation.Store.Dispose()
                Throw
            Catch ex As System.OperationCanceledException
                If preparation.Store IsNot Nothing Then preparation.Store.Dispose()
                Throw
            Catch ex As System.Exception When IsCooperativeAvailabilityFailure(ex)
                If preparation.Store IsNot Nothing Then preparation.Store.Dispose()
                preparation.Store = Nothing
                preparation.Document.CooperativeState = "local_only"
                preparation.Document.CooperativeDiagnostic = ex.Message
                result.Diagnostics.Add("cooperative_local_only: " & item.DocumentId & "; shared reuse/contribution is unavailable (" & ex.Message & "). Processing continues in the private writable shadow.")
                Return preparation
            Catch
                If preparation.Store IsNot Nothing Then preparation.Store.Dispose()
                Throw
            End Try
        End Function

        ''' <summary>
        ''' Apply the original-byte policy before copying any shared index into the
        ''' private projection. A mismatched semantic record can still supply its
        ''' verified extraction; its card and section index cannot satisfy this policy.
        ''' </summary>
        Friend Shared Function FilterCooperativeImport(imported As SemanticArchiveCooperativeImport,
                                                       originalByteLength As System.Int64, thresholdBytes As System.Int64) As SemanticArchiveCooperativeImport
            If imported Is Nothing Then Return Nothing
            If imported.Document Is Nothing OrElse imported.Document.Fingerprint Is Nothing OrElse
                imported.Document.Fingerprint.Length <> originalByteLength Then
                Throw New System.IO.InvalidDataException("The shared record's original byte length does not match the independently verified source.")
            End If
            Dim needsIndex As System.Boolean = SemanticArchiveIndexPolicy.RequiresDocumentSectionIndex(originalByteLength, thresholdBytes)
            Dim hasIndex As System.Boolean = imported.Document.Index IsNot Nothing
            Dim incompatible As System.Boolean = (imported.Stage = SemanticArchiveCooperativeStore.CompleteStage AndAlso hasIndex <> needsIndex) OrElse
                (imported.Stage = SemanticArchiveCooperativeStore.IndexedStage AndAlso Not needsIndex)
            If Not incompatible Then Return imported
            Dim extraction As SemanticArchiveDocumentRecord = SemanticArchiveMetadata.Clone(imported.Document)
            extraction.Index = Nothing
            extraction.Card = Nothing
            extraction.SemanticSignature = ""
            extraction.SemanticModelIdentity = ""
            extraction.Active = False
            extraction.ProcessingStatus = "extracted"
            Return New SemanticArchiveCooperativeImport With {.Stage = SemanticArchiveCooperativeStore.ExtractedStage, .Document = extraction}
        End Function

        Private Shared Sub CheckpointCooperativeImport(document As SemanticArchiveDocumentRecord, item As SemanticArchiveWorkItem,
                                                       fingerprint As SemanticArchiveSourceFingerprint, queue As SemanticArchiveWorkQueue)
            BindDocumentToSource(document, item, item.SourcePath, fingerprint)
            item.CachedDocument = document
            item.State = If(document.Index Is Nothing, "Extracted", "Indexed")
            queue.Save(item)
        End Sub

        Private Shared Async Function ImportIfUsefulAsync(cooperative As SemanticArchiveCooperativeStore,
                                                          imported As SemanticArchiveCooperativeImport,
                                                          archive As SemanticArchiveDefinition, binding As SemanticArchiveSourceBinding,
                                                          item As SemanticArchiveWorkItem, existing As SemanticArchiveDocumentRecord,
                                                          result As SemanticArchiveBuildResult,
                                                          cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of SemanticArchiveDocumentRecord)
            Dim needsIndex As System.Boolean = SemanticArchiveIndexPolicy.RequiresDocumentSectionIndex(item.SourceLength, archive.SectionIndexThresholdBytes)
            Dim existingExtraction As System.Boolean = ValidateReusableRepresentation(binding, existing.Representation, item.SourceHash, item.ExtractionSignature)
            If existingExtraction Then
                If imported.Stage = SemanticArchiveCooperativeStore.ExtractedStage Then Return existing
                If Not item.ForceSemanticRebuild AndAlso existing.Card IsNot Nothing AndAlso existing.SemanticSignature = item.SemanticSignature AndAlso
                    (existing.Index IsNot Nothing) = needsIndex AndAlso
                    (existing.Index Is Nothing OrElse ValidateReusableIndex(binding, existing.Index)) Then
                    existing.CooperativeState = "imported"
                    existing.CooperativeDiagnostic = "A matching committed shared document is available; the validated private document was retained."
                    Return existing
                End If
                If needsIndex AndAlso imported.Stage = SemanticArchiveCooperativeStore.IndexedStage AndAlso existing.Index IsNot Nothing AndAlso
                    existing.Index.ModelIdentity = item.SemanticSignature AndAlso ValidateReusableIndex(binding, existing.Index) Then Return existing
            End If
            Dim destination As System.String = SemanticArchiveArtifactPlanner.CreatePrivateVersionDirectory(binding, item.CanonicalSourceKey, SemanticArchiveIdentity.NewId())
            Dim document As SemanticArchiveDocumentRecord = cooperative.Materialize(imported, destination, cancellationToken)
            Dim text As System.String = System.IO.File.ReadAllText(document.Representation.TextPath, New System.Text.UTF8Encoding(False, True))
            Dim bytes As System.Byte() = New System.Text.UTF8Encoding(False, True).GetBytes(text)
            Dim payloadHash As System.String = SemanticArchiveIdentity.HashBytes(bytes)
            If document.Index IsNot Nothing Then
                Dim cached As SharedMethods.SemanticSearchIndexCacheItem = Await SharedMethods.TryGetSemanticSearchIndexAsync(document.Index.Path, cancellationToken).ConfigureAwait(False)
                If cached Is Nothing OrElse cached.IndexDocument.ContentSha256 <> payloadHash OrElse document.Index.PayloadHash <> payloadHash OrElse
                    document.Index.ProfileVersion <> archive.SemanticProfileVersion OrElse document.Index.ModelIdentity <> item.SemanticSignature Then
                    Throw New System.IO.InvalidDataException("The shared v1 index does not describe the exact extracted UTF-8 payload and requested profile.")
                End If
            ElseIf imported.Stage = SemanticArchiveCooperativeStore.CompleteStage AndAlso needsIndex Then
                Throw New System.IO.InvalidDataException("A completed shared record is missing the section index required by its original source byte length.")
            End If
            result.Diagnostics.Add("cooperative_reuse: " & item.DocumentId & "; imported validated " & imported.Stage & " artifacts without repeating completed processing stages.")
            document.CooperativeState = "imported"
            document.CooperativeDiagnostic = "Validated " & imported.Stage & " artifacts copied into the private projection."
            Return document
        End Function

        Private Shared Sub BindDocumentToSource(document As SemanticArchiveDocumentRecord, item As SemanticArchiveWorkItem,
                                               sourcePath As System.String, fingerprint As SemanticArchiveSourceFingerprint)
            document.DocumentId = item.DocumentId
            document.SourceItemId = item.SourceItemId
            document.CanonicalSourceKey = item.CanonicalSourceKey
            document.PartitionKey = item.PartitionKey
            document.SourcePath = sourcePath
            document.RelativePath = item.RelativePath
            document.DisplayName = System.IO.Path.GetFileName(sourcePath)
            document.BindingIds = New System.Collections.Generic.List(Of System.String)(item.BindingIds)
            document.Fingerprint = fingerprint
            document.ExtractionSignature = item.ExtractionSignature
            If document.Card IsNot Nothing Then
                document.Card.CardId = SemanticArchiveIdentity.StableId("card", "DOCUMENT:" & document.DocumentId)
                document.Card.TargetId = document.DocumentId
                document.Card.PartitionKey = document.PartitionKey
                document.Card.FullMetadataReference = document.DocumentId
                document.Card.SourceVersion = fingerprint.Sha256
                document.Card.RepresentationId = document.Representation.RepresentationId
                document.Card.RetrievalText = SemanticArchiveMetadata.RenderCard(document.Card)
            End If
        End Sub

        Private Shared Sub PublishCooperativeCheckpoint(cooperative As SemanticArchiveCooperativeStore,
                                                        document As SemanticArchiveDocumentRecord, stage As System.String,
                                                        semanticSignature As System.String, result As SemanticArchiveBuildResult,
                                                        cancellationToken As System.Threading.CancellationToken)
            If cooperative Is Nothing OrElse Not cooperative.HasClaim Then Return
            Try
                cooperative.Publish(document, stage, semanticSignature, cancellationToken)
                If stage = SemanticArchiveCooperativeStore.CompleteStage Then
                    document.CooperativeState = "contributed"
                    document.CooperativeDiagnostic = "The shared immutable document is committed."
                    result.Diagnostics.Add("cooperative_contributed: " & document.DocumentId & "; the shared immutable document is committed.")
                End If
            Catch ex As System.Exception When IsCooperativeAvailabilityFailure(ex)
                cooperative.Dispose()
                document.CooperativeState = "contribution_pending"
                document.CooperativeDiagnostic = ex.Message
                result.Diagnostics.Add("cooperative_contribution_pending: " & document.DocumentId & "; " & ex.Message & ". The private checkpoint remains usable; a refresh can retry sharing.")
            End Try
        End Sub

        Private Shared Function IsCooperativeAvailabilityFailure(exception As System.Exception) As System.Boolean
            Return TypeOf exception Is System.IO.IOException OrElse TypeOf exception Is System.UnauthorizedAccessException OrElse
                TypeOf exception Is System.IO.InvalidDataException OrElse
                TypeOf exception Is System.Security.SecurityException OrElse TypeOf exception Is System.NotSupportedException OrElse
                TypeOf exception Is Newtonsoft.Json.JsonException
        End Function
    End Class
End Namespace
