' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Strict On
Option Explicit On
Option Infer On

Namespace SharedLibrary

    ''' <summary>Published inventory, not a grant of current source access. Each changed
    ''' bounded document shard computes its counters once; publication aggregates those
    ''' counters without rereading unchanged source files or extracted text.</summary>
    Public NotInheritable Class SemanticArchiveInventory
        Public Property CurrentSources As System.Int32
        Public Property TextRepresentations As System.Int32
        Public Property SearchableDocuments As System.Int32
        Public Property CompleteExtractions As System.Int32
        Public Property UnknownExtractions As System.Int32
        Public Property IncompleteExtractions As System.Int32
        Public Property ExcludedUnknown As System.Int32
        Public Property ExcludedIncomplete As System.Int32
        Public Property EmptySources As System.Int32
        Public Property FailedSources As System.Int32
        Public Property PendingSources As System.Int32
        Public Property RemovedSources As System.Int32

        ' Derived from existing counters, including already-published current-format shards.
        ' No schema migration or zero-filled historical status counter is required.
        Public ReadOnly Property ExcludedComplete As System.Int32
            Get
                Return System.Math.Max(0, CompleteExtractions - (SearchableDocuments -
                    (UnknownExtractions - ExcludedUnknown) - (IncompleteExtractions - ExcludedIncomplete)))
            End Get
        End Property

        Public Shared Function IsSearchable(document As SemanticArchiveDocumentRecord) As System.Boolean
            Return document IsNot Nothing AndAlso document.Active AndAlso document.Card IsNot Nothing AndAlso
                document.Representation IsNot Nothing AndAlso document.Representation.Completeness <> "empty"
        End Function

        Public Shared Function FromDocuments(documents As System.Collections.Generic.IEnumerable(Of SemanticArchiveDocumentRecord)) As SemanticArchiveInventory
            If documents Is Nothing Then Throw New System.ArgumentNullException(NameOf(documents))
            Dim counts As New SemanticArchiveInventory()
            For Each document As SemanticArchiveDocumentRecord In documents
                If document Is Nothing Then Throw New System.IO.InvalidDataException("A document record is missing.")
                If document.ProcessingStatus = "removed" Then
                    counts.RemovedSources += 1
                    Continue For
                End If
                counts.CurrentSources += 1
                If IsSearchable(document) Then counts.SearchableDocuments += 1
                If document.ProcessingStatus = "failed" OrElse document.ProcessingStatus = "unavailable" Then counts.FailedSources += 1
                If document.ProcessingStatus = "pending_host" OrElse document.ProcessingStatus = "pending" OrElse document.ProcessingStatus = "extracted" Then counts.PendingSources += 1
                If document.ProcessingStatus = "empty" OrElse
                    (document.Representation IsNot Nothing AndAlso document.Representation.Completeness = "empty") Then
                    counts.EmptySources += 1
                End If
                If document.Representation Is Nothing OrElse document.Representation.Completeness = "empty" Then Continue For
                counts.TextRepresentations += 1
                Select Case document.Representation.Completeness
                    Case "complete"
                        counts.CompleteExtractions += 1
                    Case "unknown"
                        counts.UnknownExtractions += 1
                        If Not IsSearchable(document) Then counts.ExcludedUnknown += 1
                    Case "incomplete"
                        counts.IncompleteExtractions += 1
                        If Not IsSearchable(document) Then counts.ExcludedIncomplete += 1
                End Select
            Next
            counts.Validate()
            Return counts
        End Function

        Public Sub Add(other As SemanticArchiveInventory)
            If other Is Nothing Then Throw New System.IO.InvalidDataException("The document inventory is missing; rebuild the archive.")
            other.Validate()
            CurrentSources += other.CurrentSources
            TextRepresentations += other.TextRepresentations
            SearchableDocuments += other.SearchableDocuments
            CompleteExtractions += other.CompleteExtractions
            UnknownExtractions += other.UnknownExtractions
            IncompleteExtractions += other.IncompleteExtractions
            ExcludedUnknown += other.ExcludedUnknown
            ExcludedIncomplete += other.ExcludedIncomplete
            EmptySources += other.EmptySources
            FailedSources += other.FailedSources
            PendingSources += other.PendingSources
            RemovedSources += other.RemovedSources
        End Sub

        Public Sub Validate()
            For Each value As System.Int32 In New System.Int32() {CurrentSources, TextRepresentations, SearchableDocuments, CompleteExtractions, UnknownExtractions, IncompleteExtractions, ExcludedUnknown, ExcludedIncomplete, EmptySources, FailedSources, PendingSources, RemovedSources}
                If value < 0 Then Throw New System.IO.InvalidDataException("Negative archive inventory count.")
            Next
            If TextRepresentations > CurrentSources OrElse SearchableDocuments > TextRepresentations OrElse
                CompleteExtractions + UnknownExtractions + IncompleteExtractions <> TextRepresentations OrElse
                ExcludedUnknown > UnknownExtractions OrElse ExcludedIncomplete > IncompleteExtractions OrElse
                EmptySources > CurrentSources OrElse FailedSources > CurrentSources OrElse PendingSources > CurrentSources Then
                Throw New System.IO.InvalidDataException("Inconsistent archive inventory counts.")
            End If
        End Sub

        Public Function SameAs(other As SemanticArchiveInventory) As System.Boolean
            If other Is Nothing Then Return False
            Return CurrentSources = other.CurrentSources AndAlso
                TextRepresentations = other.TextRepresentations AndAlso
                SearchableDocuments = other.SearchableDocuments AndAlso
                CompleteExtractions = other.CompleteExtractions AndAlso
                UnknownExtractions = other.UnknownExtractions AndAlso
                IncompleteExtractions = other.IncompleteExtractions AndAlso
                ExcludedUnknown = other.ExcludedUnknown AndAlso
                ExcludedIncomplete = other.ExcludedIncomplete AndAlso
                EmptySources = other.EmptySources AndAlso
                FailedSources = other.FailedSources AndAlso
                PendingSources = other.PendingSources AndAlso
                RemovedSources = other.RemovedSources
        End Function

        Public Function ToDiagnosticText() As System.String
            Dim attention As System.Int32 = CurrentSources - SearchableDocuments
            Dim lines As New System.Collections.Generic.List(Of System.String) From {
                "Archive contents: " & CurrentSources.ToString(System.Globalization.CultureInfo.InvariantCulture) & " document(s).",
                "Searchable now: " & SearchableDocuments.ToString(System.Globalization.CultureInfo.InvariantCulture) & ".",
                "Text available: " & TextRepresentations.ToString(System.Globalization.CultureInfo.InvariantCulture) & ".",
                "Need attention: " & attention.ToString(System.Globalization.CultureInfo.InvariantCulture) & "."
            }
            If attention > 0 Then
                lines.Add("  Complete text without searchable metadata: " & ExcludedComplete.ToString(System.Globalization.CultureInfo.InvariantCulture) & "; incomplete extraction: " & ExcludedIncomplete.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; unverified extraction: " & ExcludedUnknown.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; no readable text: " & EmptySources.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; failed/unavailable: " & FailedSources.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; pending: " & PendingSources.ToString(System.Globalization.CultureInfo.InvariantCulture) & ".")
            End If
            If ExcludedComplete > 0 Then lines.Add("Complete extracted text is retained but some documents require extraction-contract validation or semantic metadata repair. Run Retry failed or repair; permissions alone do not rebuild content.")
            Return System.String.Join(System.Environment.NewLine, lines)
        End Function

        Public Function ToTechnicalDiagnosticText() As System.String
            Return "Published inventory (current source rights are checked again at retrieval):" & System.Environment.NewLine &
                "Current source records: " & CurrentSources.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; extracted text: " & TextRepresentations.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; searchable documents: " & SearchableDocuments.ToString(System.Globalization.CultureInfo.InvariantCulture) & System.Environment.NewLine &
                "Extraction complete: " & CompleteExtractions.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; unknown: " & UnknownExtractions.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; incomplete: " & IncompleteExtractions.ToString(System.Globalization.CultureInfo.InvariantCulture) & System.Environment.NewLine &
                "Complete extraction but not searchable: " & ExcludedComplete.ToString(System.Globalization.CultureInfo.InvariantCulture) & System.Environment.NewLine &
                "Not searchable with unknown coverage: " & ExcludedUnknown.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; not searchable with incomplete coverage: " & ExcludedIncomplete.ToString(System.Globalization.CultureInfo.InvariantCulture) & System.Environment.NewLine &
                "Empty: " & EmptySources.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; failed/unavailable: " & FailedSources.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; pending in published records: " & PendingSources.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; retired records retained in this generation: " & RemovedSources.ToString(System.Globalization.CultureInfo.InvariantCulture) & System.Environment.NewLine &
                "Discovered but filtered files are not source records. Live queue counts are reported separately. " &
                "Unknown/incomplete text is excluded when Allow incomplete extracted text is off; other validity checks may also suppress it."
        End Function
    End Class
End Namespace
