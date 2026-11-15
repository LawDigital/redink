' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. For license to use see https://redink.ai.
' Combined reviews reuse the full-source analysis contract. Each synthetic segment
' maps to exactly one unchanged original segment; only the host maps citations back.
Option Explicit On
Option Strict On

Namespace Global.SharedLibrary.TabularReview
    Public NotInheritable Partial Class ReviewContradictionEngine
        Private NotInheritable Class PromptSource
            Public Property Document As ReviewDocument
            Public Property Segment As ReviewSourceSegment
        End Class

        Private NotInheritable Class PreparedPrompt
            Public Property Run As New ReviewConsistencyRun With {.Mode = "prompt"}
            Public Property Combined As New ReviewDocument With {.FileName = "Selected document set", .ExtractionStatus = "complete"}
            Public Property Criterion As New ReviewCriterion With {.Title = "Combined document review", .Kind = "analysis"}
            Public Property Sources As New System.Collections.Generic.Dictionary(Of System.String, PromptSource)(System.StringComparer.Ordinal)
            Public Property Engine As ReviewAnalysisEngine
        End Class

        Public Function PlanPrompt(documents As System.Collections.Generic.IEnumerable(Of ReviewDocument), instruction As System.String,
                                   outputLanguage As System.String, operationModelKey As System.String,
                                   cancellationToken As System.Threading.CancellationToken, Optional criterion As ReviewCriterion = Nothing) As ReviewConsistencyRun
            ValidateArguments(documents, operationModelKey)
            Dim prepared As New PreparedPrompt()
            If criterion IsNot Nothing Then prepared.Criterion = criterion
            prepared.Run.StartedUtc = ""
            PreparePrompt(prepared, documents, instruction, outputLanguage, operationModelKey, cancellationToken)
            Return CloneRun(prepared.Run)
        End Function

        Private Sub PreparePrompt(prepared As PreparedPrompt, documents As System.Collections.Generic.IEnumerable(Of ReviewDocument),
                                  instruction As System.String, outputLanguage As System.String, operationModelKey As System.String,
                                  cancellationToken As System.Threading.CancellationToken)
            If System.String.IsNullOrWhiteSpace(instruction) Then Throw New System.IO.InvalidDataException("Enter the prompt to run over the selected documents.")
            Dim run As ReviewConsistencyRun = prepared.Run
            run.Instruction = instruction
            run.ModelKey = operationModelKey
            If prepared.Criterion.Instruction.Length = 0 Then prepared.Criterion.Instruction = instruction
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each original As ReviewDocument In documents
                cancellationToken.ThrowIfCancellationRequested()
                If original Is Nothing Then Throw New System.IO.InvalidDataException("The selected scope contains a missing document record.")
                ReviewContract.RequireId(original.Id)
                If Not seen.Add(original.Id) Then Throw New System.IO.InvalidDataException("The selected scope contains the same document twice.")
                Dim coverage As New ReviewConsistencyDocumentCoverage With {
                    .DocumentId = original.Id, .FileName = If(original.FileName, ""), .SourceHash = If(original.SourceHash, ""),
                    .TextHash = If(original.TextHash, ""), .ExtractionStatus = If(original.ExtractionStatus, ""), .Status = "unavailable"
                }
                run.Documents.Add(coverage)
                Dim source As ReviewDocument = SnapshotSource(original, coverage, cancellationToken)
                If source Is Nothing Then Throw New System.IO.InvalidDataException("Refresh or relink " & If(original.FileName, "the selected document") & " before starting the combined review. " & coverage.Message)
                coverage.Status = "included"
                coverage.SourceParts = source.Segments.Count
                coverage.Message = If(source.ExtractionStatus = "complete", "The review includes the complete saved extract.", "The review includes the available extract only. Check the original for missing content.")
                If source.ExtractionStatus <> "complete" Then prepared.Combined.ExtractionStatus = "partial"
                Dim sourceLabel As System.String = "Document " & run.Documents.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & ": " & source.FileName
                If Not System.String.IsNullOrWhiteSpace(original.Number) Then sourceLabel &= " (" & original.Number & ")"
                For Each segment As ReviewSourceSegment In source.Segments
                    cancellationToken.ThrowIfCancellationRequested()
                    Dim key As System.String = source.Id & ":" & segment.Id
                    prepared.Sources.Add(key, New PromptSource With {.Document = source, .Segment = segment})
                    prepared.Combined.Segments.Add(New ReviewSourceSegment With {
                        .Id = key, .Locator = sourceLabel & If(System.String.IsNullOrWhiteSpace(segment.Locator), "", " — " & segment.Locator),
                        .Kind = segment.Kind, .PageNumber = segment.PageNumber, .Text = segment.Text
                    })
                Next
            Next
            If seen.Count < 2 Then Throw New System.IO.InvalidDataException("Select at least two documents for a combined review.")
            prepared.Combined.SourceHash = ReviewContract.Digest(ReviewContract.Serialize(run.Documents))
            prepared.Combined.TextHash = ReviewContract.Digest(ReviewContract.Serialize(prepared.Combined.Segments))
            prepared.Combined.ExtractionMessage = If(prepared.Combined.ExtractionStatus = "complete", "", "At least one selected document has only a partial extract.")
            prepared.Engine = New ReviewAnalysisEngine(
                Function(systemPrompt As System.String, request As System.String, cancellation As System.Threading.CancellationToken)
                    Return _invoke(operationModelKey, systemPrompt, request, cancellation)
                End Function, _requestBudget)
            run.PlannedSourcePasses = prepared.Engine.PlanSourcePasses(prepared.Combined, prepared.Criterion, If(outputLanguage, ""), cancellationToken)
            If run.PlannedSourcePasses = 0 Then Throw New System.IO.InvalidDataException("The selected set has no readable source text.")
            run.CoverageMessage = "The prompt covers " & run.Documents.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " selected documents. " &
                If(run.PlannedSourcePasses = 1, "One source pass reads their saved text together.",
                   "The full saved text requires " & run.PlannedSourcePasses.ToString(System.Globalization.CultureInfo.InvariantCulture) & " source passes, followed by synthesis. Synthesis work depends on the responses. The review retains every validated working note and its quotations.") &
                " " & PromptSourceCoverage(run) & " The model can miss information; request completion does not establish substantive completeness."
        End Sub

        Public Async Function CheckPromptAsync(documents As System.Collections.Generic.IEnumerable(Of ReviewDocument), instruction As System.String,
                                              outputLanguage As System.String, operationModelKey As System.String,
                                              cancellationToken As System.Threading.CancellationToken,
                                              persist As System.Func(Of ReviewConsistencyRun, ReviewFinding, System.Threading.Tasks.Task),
                                              Optional progress As System.Action(Of System.String) = Nothing, Optional criterion As ReviewCriterion = Nothing) As System.Threading.Tasks.Task(Of ReviewConsistencyRun)
            ValidateArguments(documents, operationModelKey)
            If persist Is Nothing Then Throw New System.ArgumentNullException(NameOf(persist))
            If System.String.IsNullOrWhiteSpace(instruction) Then Throw New System.IO.InvalidDataException("Enter the prompt to run over the selected documents.")
            Dim prepared As New PreparedPrompt()
            If criterion IsNot Nothing Then prepared.Criterion = criterion
            Dim run As ReviewConsistencyRun = prepared.Run
            run.Instruction = instruction
            run.ModelKey = operationModelKey
            run.StartedUtc = UtcNow()
            Dim storageFailed As System.Boolean = False
            Dim cancelled As System.Boolean = False
            Dim failed As System.Boolean = False
            Dim complete As System.Boolean = False
            Try
                PreparePrompt(prepared, documents, instruction, outputLanguage, operationModelKey, cancellationToken)
                Await PublishAsync(persist, run, Nothing).ConfigureAwait(False)
                Report(progress, run.CoverageMessage)
                Dim result As ReviewAnswer = Await prepared.Engine.GenerateAsync(prepared.Combined, prepared.Criterion, If(outputLanguage, ""),
                    operationModelKey, "", cancellationToken, progress,
                    Async Function(checkpoint As ReviewAnalysisCheckpoint) As System.Threading.Tasks.Task
                        run.CompletedSourcePasses = checkpoint.CompletedSourcePasses
                        run.PlannedSourcePasses = checkpoint.PlannedSourcePasses
                        run.CompletedSynthesisPasses = checkpoint.CompletedSynthesisPasses
                        run.PlannedSynthesisPasses = checkpoint.PlannedSynthesisPasses
                        If checkpoint.Result IsNot Nothing Then
                            run.WorkingNotes.Add(New ReviewPromptNote With {
                                .Stage = checkpoint.Stage, .Description = checkpoint.Result.Text, .Assessment = checkpoint.Result.Assessment,
                                .NoRelevantEvidence = checkpoint.Result.Evidence.Count = 0, .Evidence = MapPromptEvidence(prepared, checkpoint.Result.Evidence)
                            })
                        End If
                        run.CoverageMessage = PromptCoverage(run, False)
                        ' Stop retains a validated result; host shutdown still cancels
                        ' persistence through the owning form's existing lifetime token.
                        Await PublishAsync(persist, run, Nothing).ConfigureAwait(False)
                    End Function).ConfigureAwait(False)
                cancellationToken.ThrowIfCancellationRequested()
                Dim finding As New ReviewFinding With {
                    .RunId = run.Id, .Kind = "review_finding", .Title = PromptTitle(instruction), .Description = result.Text,
                    .Assessment = result.Assessment, .ModelKey = operationModelKey, .Status = "draft",
                    .NoRelevantEvidence = result.Evidence.Count = 0, .Evidence = MapPromptEvidence(prepared, result.Evidence),
                    .CreatedUtc = UtcNow(), .UpdatedUtc = UtcNow()
                }
                run.GroupFlag = result.RiskFlag
                run.FindingIds.Add(finding.Id)
                Await PublishAsync(persist, run, finding).ConfigureAwait(False)
                complete = True
            Catch ex As PersistenceFailure
                storageFailed = True
                failed = True
                run.PersistenceFailed = True
                run.Error = ex.Message
                System.Diagnostics.Trace.WriteLine("[TabularReview] combined review persistence failed; run=" & run.Id & "; " & ex.ToString())
            Catch ex As System.OperationCanceledException
                cancelled = cancellationToken.IsCancellationRequested
                failed = Not cancelled
                If failed Then run.Error = "A model request ended before the combined review finished. Check model availability and run the prompt again. Completed working notes remain available."
                System.Diagnostics.Trace.WriteLine("[TabularReview] combined review interrupted; run=" & run.Id & "; " & ex.ToString())
            Catch ex As System.Exception
                failed = True
                run.Error = If(TypeOf ex Is System.IO.InvalidDataException OrElse TypeOf ex Is System.InvalidOperationException,
                    ex.Message, "The combined review stopped before it could finish. Check the selected sources and model, then run the prompt again. Completed working notes remain available.")
                System.Diagnostics.Trace.WriteLine("[TabularReview] combined review failed; run=" & run.Id & "; " & ex.ToString())
            End Try
            run.CompletedUtc = UtcNow()
            If cancelled Then
                run.Status = "cancelled"
            ElseIf failed OrElse Not complete Then
                run.Status = If(run.WorkingNotes.Count > 0 OrElse run.FindingIds.Count > 0, "partial", "error")
            Else
                run.Status = If(run.Documents.TrueForAll(Function(item As ReviewConsistencyDocumentCoverage) item.Status = "included" AndAlso item.ExtractionStatus = "complete"), "complete", "partial")
            End If
            run.CoverageMessage = PromptCoverage(run, True)
            Report(progress, run.CoverageMessage)
            If Not storageFailed Then
                Try
                    Await PublishAsync(persist, run, Nothing).ConfigureAwait(False)
                Catch ex As PersistenceFailure
                    run.PersistenceFailed = True
                    run.Status = "partial"
                    run.Error = ex.Message
                    System.Diagnostics.Trace.WriteLine("[TabularReview] combined review terminal persistence failed; run=" & run.Id & "; " & ex.ToString())
                End Try
            End If
            Return CloneRun(run)
        End Function

        Private Shared Function MapPromptEvidence(prepared As PreparedPrompt, values As System.Collections.Generic.List(Of ReviewEvidence)) As System.Collections.Generic.List(Of ReviewCrossDocumentEvidence)
            Dim result As New System.Collections.Generic.List(Of ReviewCrossDocumentEvidence)()
            For Each evidence As ReviewEvidence In values
                Dim source As PromptSource = Nothing
                If Not prepared.Sources.TryGetValue(evidence.SegmentId, source) Then Throw New System.IO.InvalidDataException("A combined review cites a source outside its selected set.")
                Dim mapped As ReviewEvidence = ReviewAnalysisEngine.ValidateEvidence(source.Document, source.Segment.Id, evidence.Quote, evidence.Start)
                If mapped.MatchStatus = "unverified" OrElse mapped.Start <> evidence.Start OrElse mapped.Length <> evidence.Length Then Throw New System.IO.InvalidDataException("A combined-review quotation does not match its original source span.")
                ' Preserve the evidence identity across the source and synthesis passes.
                mapped.Id = evidence.Id
                result.Add(New ReviewCrossDocumentEvidence With {
                    .DocumentId = source.Document.Id, .SourceHash = source.Document.SourceHash, .TextHash = source.Document.TextHash, .Evidence = mapped
                })
            Next
            Return result
        End Function

        Private Shared Function PromptTitle(instruction As System.String) As System.String
            Dim value As System.String = instruction.Trim().Split(New System.Char() {Microsoft.VisualBasic.ControlChars.Cr, Microsoft.VisualBasic.ControlChars.Lf}, System.StringSplitOptions.RemoveEmptyEntries)(0)
            If value.Length <= 120 Then Return value
            Dim length As System.Int32 = If(System.Char.IsHighSurrogate(value(119)), 119, 120)
            Return value.Substring(0, length) & "…"
        End Function

        Private Shared Function PromptSourceCoverage(run As ReviewConsistencyRun) As System.String
            Dim partialCount As System.Int32 = run.Documents.FindAll(Function(item As ReviewConsistencyDocumentCoverage) item.ExtractionStatus <> "complete" OrElse item.Status <> "included").Count
            If partialCount = 0 Then Return "Every selected document has a complete saved extract."
            Return partialCount.ToString(System.Globalization.CultureInfo.InvariantCulture) & " selected document(s) have incomplete or unavailable extracts; conclusions cover only the saved text."
        End Function

        Private Shared Function PromptCoverage(run As ReviewConsistencyRun, terminal As System.Boolean) As System.String
            Dim message As System.String = "Read " & run.CompletedSourcePasses.ToString(System.Globalization.CultureInfo.InvariantCulture) & " of " & run.PlannedSourcePasses.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                " source passes; completed " & run.CompletedSynthesisPasses.ToString(System.Globalization.CultureInfo.InvariantCulture) & " of " & run.PlannedSynthesisPasses.ToString(System.Globalization.CultureInfo.InvariantCulture) & " planned synthesis passes. " & PromptSourceCoverage(run)
            If run.PlannedSourcePasses > 1 Then message &= " This review synthesizes validated partial answers; every completed working note and its source quotations remain available."
            If terminal Then
                If run.FindingIds.Count = 0 Then
                    message &= " No final answer is available. Inspect the completed working notes and run the prompt again to finish."
                Else
                    message &= " The review retained its final response. Check its meaning and supporting sources before marking it reviewed."
                End If
                If run.Status <> "complete" Then message &= " This run has incomplete coverage."
            End If
            Return message
        End Function
    End Class
End Namespace
