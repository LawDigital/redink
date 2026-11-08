' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. For license to use see https://redink.ai.
' The persisted review contract contains data only. Office, browser and model adapters
' validate this contract; no provider, organization or document template owns it.
Option Explicit On
Option Strict On

Namespace Global.SharedLibrary.TabularReview
    Public NotInheritable Class ReviewProject
        Public Property SchemaVersion As System.Int32 = 3
        Public Property Id As System.String = System.Guid.NewGuid().ToString("N")
        Public Property Title As System.String = "Tabular Review"
        Public Property Revision As System.Int64
        Public Property CreatedUtc As System.String = System.DateTime.UtcNow.ToString("o")
        Public Property UpdatedUtc As System.String = System.DateTime.UtcNow.ToString("o")
        Public Property SourceRoot As System.String = ""
        ' Relative to the workbook directory; only contained roots form a portable folder.
        Public Property SourceRootRelativePath As System.String = ""
        Public Property IncludeSubfolders As System.Boolean = True
        Public Property UseOcr As System.Boolean
        Public Property OutputLanguage As System.String = ""
        Public Property ModelKey As System.String = ""
        Public Property Criteria As New System.Collections.Generic.List(Of ReviewCriterion)()
        Public Property Documents As New System.Collections.Generic.List(Of ReviewDocument)()
        Public Property Audit As New System.Collections.Generic.List(Of ReviewAuditEntry)()
        Public Property FlagColorKeyVersion As System.Int32
        Public Property FlagColors As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
        Public Property Findings As New System.Collections.Generic.List(Of ReviewFinding)()
        Public Property ConsistencyRuns As New System.Collections.Generic.List(Of ReviewConsistencyRun)()
        Public Property ReportDefinitions As New System.Collections.Generic.List(Of ReviewReportDefinition)()
        Public Property ReportFiles As New System.Collections.Generic.List(Of ReviewReportFile)()
        Public Property View As New ReviewViewState()
    End Class

    Public NotInheritable Class ReviewCriterion
        Public Property Id As System.String = System.Guid.NewGuid().ToString("N")
        Public Property Title As System.String = ""
        Public Property Instruction As System.String = ""
        ' analysis | text | priority | flag. The values of custom columns remain user-defined.
        Public Property Kind As System.String = "analysis"
        Public Property Options As New System.Collections.Generic.List(Of System.String)()
        Public Property Revision As System.Int32 = 1
        Public Property Archived As System.Boolean
    End Class

    Public NotInheritable Class ReviewDocument
        Public Property Id As System.String = System.Guid.NewGuid().ToString("N")
        Public Property Number As System.String = ""
        Public Property FileName As System.String = ""
        Public Property RelativePath As System.String = ""
        Public Property SourcePath As System.String = ""
        Public Property SourceHash As System.String = ""
        Public Property SourceLength As System.Int64
        Public Property SourceLastWriteUtc As System.String = ""
        Public Property TextHash As System.String = ""
        ' pending | complete | partial | error | missing | changed | imported
        Public Property ExtractionStatus As System.String = "pending"
        Public Property ExtractionMessage As System.String = ""
        ' Cached coverage survives temporary original-file availability failures.
        Public Property StoredExtractionStatus As System.String = ""
        Public Property StoredExtractionMessage As System.String = ""
        Public Property Segments As New System.Collections.Generic.List(Of ReviewSourceSegment)()
        Public Property SourceHistory As New System.Collections.Generic.List(Of ReviewSourceVersion)()
        Public Property ReviewStatus As System.String = "unreviewed"
        Public Property Priority As System.String = ""
        Public Property Flags As New System.Collections.Generic.List(Of System.String)()
        Public Property Reviewer As System.String = ""
        Public Property Notes As System.String = ""
        Public Property LastReviewedUtc As System.String = ""
        Public Property ImportedFrom As System.String = ""
        Public Property Archived As System.Boolean
        Public Property Answers As New System.Collections.Generic.Dictionary(Of System.String, ReviewAnswer)(System.StringComparer.Ordinal)
        Public Property UpdatedBy As System.String = ""
        Public Property ReviewedBy As System.String = ""
        Public Property ReviewedUtc As System.String = ""
        Public Property UpdatedUtc As System.String = ""
    End Class

    Public NotInheritable Class ReviewSourceSegment
        Public Property Id As System.String = ""
        Public Property Locator As System.String = ""
        Public Property Kind As System.String = "text"
        Public Property PageNumber As System.Int32
        Public Property Text As System.String = ""
    End Class

    Public NotInheritable Class ReviewSourceVersion
        Public Property SourcePath As System.String = ""
        Public Property FileName As System.String = ""
        Public Property RelativePath As System.String = ""
        Public Property SourceHash As System.String = ""
        Public Property TextHash As System.String = ""
        Public Property ExtractionStatus As System.String = ""
        Public Property ExtractionMessage As System.String = ""
        Public Property ArchivedUtc As System.String = System.DateTime.UtcNow.ToString("o")
        Public Property Segments As New System.Collections.Generic.List(Of ReviewSourceSegment)()
    End Class

    Public NotInheritable Class ReviewEvidence
        Public Property Id As System.String = System.Guid.NewGuid().ToString("N")
        Public Property SegmentId As System.String = ""
        Public Property Quote As System.String = ""
        Public Property Locator As System.String = ""
        Public Property Start As System.Int32 = -1
        Public Property Length As System.Int32
        ' exact | whitespace | unverified; only the host sets these fields.
        Public Property MatchStatus As System.String = "unverified"
    End Class

    Public NotInheritable Class ReviewAnswer
        Public Property Text As System.String = ""
        Public Property Assessment As System.String = ""
        ' pending | draft | reviewed | stale | error | imported
        Public Property Status As System.String = "pending"
        Public Property CriterionRevision As System.Int32
        Public Property SourceHash As System.String = ""
        Public Property ModelKey As System.String = ""
        Public Property UpdatedUtc As System.String = ""
        Public Property UserEdited As System.Boolean
        Public Property [Error] As System.String = ""
        Public Property Evidence As New System.Collections.Generic.List(Of ReviewEvidence)()
        Public Property History As New System.Collections.Generic.List(Of ReviewAnswerVersion)()
        Public Property FollowUps As New System.Collections.Generic.List(Of ReviewTurn)()
        Public Property UpdatedBy As System.String = ""
        Public Property ReviewedBy As System.String = ""
        Public Property ReviewedUtc As System.String = ""
    End Class

    Public NotInheritable Class ReviewAnswerVersion
        Public Property Text As System.String = ""
        Public Property Assessment As System.String = ""
        Public Property Status As System.String = ""
        Public Property UpdatedUtc As System.String = ""
        Public Property Actor As System.String = ""
        Public Property ModelKey As System.String = ""
        Public Property SourceHash As System.String = ""
        Public Property CriterionRevision As System.Int32
        Public Property UserEdited As System.Boolean
        Public Property [Error] As System.String = ""
        Public Property Evidence As New System.Collections.Generic.List(Of ReviewEvidence)()
        Public Property UpdatedBy As System.String = ""
        Public Property ReviewedBy As System.String = ""
        Public Property ReviewedUtc As System.String = ""
    End Class

    Public NotInheritable Class ReviewTurn
        Public Property Instruction As System.String = ""
        Public Property Answer As System.String = ""
        Public Property ModelKey As System.String = ""
        Public Property CreatedUtc As System.String = System.DateTime.UtcNow.ToString("o")
    End Class

    Public NotInheritable Class ReviewAuditEntry
        Public Property TimeUtc As System.String = System.DateTime.UtcNow.ToString("o")
        Public Property Action As System.String = ""
        Public Property DocumentId As System.String = ""
        Public Property CriterionId As System.String = ""
        Public Property Detail As System.String = ""
        Public Property Actor As System.String = ""
    End Class

    Public NotInheritable Class ReviewViewState
        Public Property DocumentId As System.String = ""
        Public Property CriterionId As System.String = ""
        Public Property Theme As System.String = "light"
        Public Property Mode As System.String = "matrix"
        Public Property LayoutVersion As System.Int32 = 1
        Public Property ViewerSide As System.String = "right"
        Public Property ViewerWidth As System.Double = 40.0R
        Public Property AutoSave As System.Boolean = True
        Public Property ColumnOrder As New System.Collections.Generic.List(Of System.String)()
        Public Property ColumnWidths As New System.Collections.Generic.Dictionary(Of System.String, System.Int32)(System.StringComparer.Ordinal)
        Public Property RowHeights As New System.Collections.Generic.Dictionary(Of System.String, System.Int32)(System.StringComparer.Ordinal)
        Public Property Search As System.String = ""
        Public Property StatusFilter As System.String = ""
        Public Property PriorityFilter As System.String = ""
        Public Property FlagFilter As System.String = ""
        Public Property SortColumn As System.String = ""
        Public Property SortDirection As System.String = ""
    End Class

    Public NotInheritable Class ReviewPreviewResult
        ' Local, temporary file owned by the preview adapter; never a persisted dependency.
        Public Property FilePath As System.String = ""
        Public Property PageNumber As System.Int32 = 1
        Public Property Highlighted As System.Boolean
        Public Property Message As System.String = ""
        Public Property MessageKind As System.String = "warning"
    End Class

    Public NotInheritable Class ReviewWorkbookSaveResult
        ' saved | cancelled | skipped | failed. Only an actual picker cancellation is cancelled.
        Public Property Status As System.String = "skipped"
        Public Property Message As System.String = ""
    End Class

    Public NotInheritable Class ReviewWorkbookSaveState
        Public Property Saved As System.Boolean
        Public Property CanAutoSave As System.Boolean
        Public Property Message As System.String = ""
    End Class

    Public NotInheritable Class ReviewCrossDocumentEvidence
        Public Property DocumentId As System.String = ""
        Public Property SourceHash As System.String = ""
        Public Property TextHash As System.String = ""
        Public Property Evidence As New ReviewEvidence()
    End Class

    Public NotInheritable Class ReviewFindingContext
        Public Property Description As System.String = ""
        Public Property Relationship As System.String = ""
        Public Property Evidence As New System.Collections.Generic.List(Of ReviewCrossDocumentEvidence)()
    End Class

    Public NotInheritable Class ReviewFinding
        Public Property Id As System.String = System.Guid.NewGuid().ToString("N")
        Public Property RunId As System.String = ""
        Public Property Title As System.String = ""
        Public Property Description As System.String = ""
        Public Property Assessment As System.String = ""
        Public Property Kind As System.String = "potential_conflict"
        Public Property Priority As System.String = "unspecified"
        Public Property Status As System.String = "draft"
        Public Property Question As System.String = ""
        Public Property RequestedAction As System.String = ""
        Public Property Notes As System.String = ""
        Public Property Reviewer As System.String = ""
        Public Property ModelKey As System.String = ""
        Public Property CreatedUtc As System.String = System.DateTime.UtcNow.ToString("o")
        Public Property UpdatedUtc As System.String = System.DateTime.UtcNow.ToString("o")
        Public Property UserEdited As System.Boolean
        Public Property Archived As System.Boolean
        Public Property NoRelevantEvidence As System.Boolean
        Public Property Evidence As New System.Collections.Generic.List(Of ReviewCrossDocumentEvidence)()
        Public Property Context As New System.Collections.Generic.List(Of ReviewFindingContext)()
        Public Property History As New System.Collections.Generic.List(Of ReviewFindingVersion)()
        Public Property UpdatedBy As System.String = ""
        Public Property ReviewedBy As System.String = ""
        Public Property ReviewedUtc As System.String = ""
    End Class

    Public NotInheritable Class ReviewFindingVersion
        Public Property Title As System.String = ""
        Public Property Description As System.String = ""
        Public Property Assessment As System.String = ""
        Public Property Kind As System.String = ""
        Public Property Priority As System.String = ""
        Public Property Status As System.String = ""
        Public Property Question As System.String = ""
        Public Property RequestedAction As System.String = ""
        Public Property Notes As System.String = ""
        Public Property Reviewer As System.String = ""
        Public Property ModelKey As System.String = ""
        Public Property UpdatedUtc As System.String = ""
        Public Property Actor As System.String = ""
        Public Property UserEdited As System.Boolean
        Public Property Archived As System.Boolean
        Public Property NoRelevantEvidence As System.Boolean
        Public Property Evidence As New System.Collections.Generic.List(Of ReviewCrossDocumentEvidence)()
        Public Property Context As New System.Collections.Generic.List(Of ReviewFindingContext)()
        Public Property UpdatedBy As System.String = ""
        Public Property ReviewedBy As System.String = ""
        Public Property ReviewedUtc As System.String = ""
    End Class

    Public NotInheritable Class ReviewConsistencyDocumentCoverage
        Public Property DocumentId As System.String = ""
        Public Property FileName As System.String = ""
        Public Property SourceHash As System.String = ""
        Public Property TextHash As System.String = ""
        Public Property ExtractionStatus As System.String = ""
        Public Property ExtractionMessage As System.String = ""
        Public Property Status As System.String = "included"
        Public Property Message As System.String = ""
        Public Property SourceCharacters As System.Int64
        Public Property SourceParts As System.Int32
    End Class

    Public NotInheritable Class ReviewPromptNote
        Public Property Stage As System.String = "source"
        Public Property Description As System.String = ""
        Public Property Assessment As System.String = ""
        Public Property NoRelevantEvidence As System.Boolean
        Public Property Evidence As New System.Collections.Generic.List(Of ReviewCrossDocumentEvidence)()
    End Class

    ' Optional, host-owned progress for combined reviews. Ordinary cell analysis
    ' keeps its existing call contract and does not persist intermediate results.
    Public NotInheritable Class ReviewAnalysisCheckpoint
        Public Property Result As New ReviewAnswer()
        Public Property Stage As System.String = "source"
        Public Property CompletedSourcePasses As System.Int32
        Public Property PlannedSourcePasses As System.Int32
        Public Property CompletedSynthesisPasses As System.Int32
        Public Property PlannedSynthesisPasses As System.Int32
    End Class

    Public NotInheritable Class ReviewConsistencyRun
        Public Property Mode As System.String = "comparison"
        Public Property Instruction As System.String = ""
        Public Property PlannedSourcePasses As System.Int32
        Public Property CompletedSourcePasses As System.Int32
        Public Property PlannedSynthesisPasses As System.Int32
        Public Property CompletedSynthesisPasses As System.Int32
        Public Property WorkingNotes As New System.Collections.Generic.List(Of ReviewPromptNote)()
        Public Property Id As System.String = System.Guid.NewGuid().ToString("N")
        Public Property StartedUtc As System.String = System.DateTime.UtcNow.ToString("o")
        Public Property CompletedUtc As System.String = ""
        Public Property ModelKey As System.String = ""
        Public Property Status As System.String = "running"
        Public Property CoverageMessage As System.String = ""
        Public Property [Error] As System.String = ""
        Public Property Documents As New System.Collections.Generic.List(Of ReviewConsistencyDocumentCoverage)()
        <Newtonsoft.Json.JsonIgnore>
        Public Property PersistenceFailed As System.Boolean
        Public Property PlannedComparisons As System.Int64
        Public Property CompletedComparisons As System.Int64
        Public Property PlannedContextChecks As System.Int64
        Public Property CompletedContextChecks As System.Int64
        Public Property FindingIds As New System.Collections.Generic.List(Of System.String)()
    End Class

    ' Detached export values. The report renderer never reaches into a live workbook.
    Public NotInheritable Class ReviewReportDefinition
        Public Property Id As System.String = System.Guid.NewGuid().ToString("N")
        Public Property Name As System.String = ""
        Public Property Title As System.String = ""
        Public Property DocumentScope As System.String = "filtered"
        Public Property Kind As System.String = "matrix_xlsx"
        Public Property PaperSize As System.String = "A4"
        Public Property Design As System.String = "redink"
        Public Property Orientation As System.String = "landscape"
        Public Property Fields As New System.Collections.Generic.List(Of System.String)()
    End Class

    Public NotInheritable Class ReviewReportColumn
        Public Property Key As System.String = ""
        Public Property Title As System.String = ""
        Public Property Width As System.Double = 40.0R
    End Class

    Public NotInheritable Class ReviewReportRow
        Public Property DocumentId As System.String = ""
        Public Property Values As New System.Collections.Generic.List(Of System.String)()
    End Class

    Public NotInheritable Class ReviewReportData
        Public Property Title As System.String = ""
        Public Property ProjectId As System.String = ""
        Public Property ProjectRevision As System.Int64
        Public Property GeneratedUtc As System.String = System.DateTime.UtcNow.ToString("o")
        Public Property PaperSize As System.String = "A4"
        Public Property Design As System.String = "redink"
        Public Property Orientation As System.String = "landscape"
        Public Property Columns As New System.Collections.Generic.List(Of ReviewReportColumn)()
        Public Property Rows As New System.Collections.Generic.List(Of ReviewReportRow)()
        Public Property Kind As System.String = "matrix_xlsx"
        Public Property Documents As New System.Collections.Generic.List(Of ReviewReportDocument)()
        Public Property Items As New System.Collections.Generic.List(Of ReviewReportItem)()
    End Class

    Public NotInheritable Class ReviewReportDocument
        Public Property Id As System.String = ""
        Public Property Number As System.String = ""
        Public Property FileName As System.String = ""
        Public Property RelativePath As System.String = ""
        Public Property SourcePath As System.String = ""
        Public Property SourceHash As System.String = ""
        Public Property ExtractionStatus As System.String = ""
        Public Property ExtractionMessage As System.String = ""
    End Class

    Public NotInheritable Class ReviewReportItem
        Public Property Id As System.String = ""
        Public Property Kind As System.String = "answer"
        Public Property Title As System.String = ""
        Public Property Question As System.String = ""
        Public Property Answer As System.String = ""
        Public Property Assessment As System.String = ""
        Public Property Priority As System.String = ""
        Public Property Status As System.String = ""
        Public Property Reviewer As System.String = ""
        Public Property Notes As System.String = ""
        Public Property ModelKey As System.String = ""
        Public Property ModelLabel As System.String = ""
        Public Property RequestedAction As System.String = ""
        Public Property Context As System.String = ""
        Public Property Coverage As System.String = ""
        Public Property Prompt As System.String = ""
        Public Property DocumentIds As New System.Collections.Generic.List(Of System.String)()
        Public Property Evidence As New System.Collections.Generic.List(Of ReviewReportEvidence)()
    End Class

    Public NotInheritable Class ReviewReportEvidence
        Public Property Role As System.String = "Source evidence"
        Public Property DocumentId As System.String = ""
        Public Property SourceHash As System.String = ""
        Public Property SegmentId As System.String = ""
        Public Property Locator As System.String = ""
        Public Property Quote As System.String = ""
        Public Property MatchStatus As System.String = ""
        Public Property Start As System.Int32 = -1
        Public Property Length As System.Int32
    End Class

    Public NotInheritable Class ReviewReportFile
        Public Property Path As System.String = ""
        Public Property RelativePath As System.String = ""
        Public Property Kind As System.String = ""
        Public Property CreatedUtc As System.String = System.DateTime.UtcNow.ToString("o")
    End Class

    Public NotInheritable Class ReviewReportExportResult
        Public Property Path As System.String = ""
        Public Property ItemCount As System.Int32
        Public Property EvidenceCount As System.Int32
        Public Property Message As System.String = ""
    End Class

    Public NotInheritable Class ReviewHostCallbacks
        Public Property EnsureAvailable As System.Action
        Public Property SaveProject As System.Action(Of ReviewProject)
        Public Property ReloadProject As System.Func(Of ReviewProject)
        Public Property ImportSelection As System.Func(Of ReviewProject)
        Public Property SaveWorkbook As System.Func(Of ReviewWorkbookSaveResult)
        Public Property SaveWorkbookAs As System.Func(Of ReviewWorkbookSaveResult)
        Public Property AutoSaveWorkbook As System.Func(Of ReviewWorkbookSaveResult)
        Public Property WorkbookSaveState As System.Func(Of ReviewWorkbookSaveState)
        Public Property WorkbookLabel As System.Func(Of System.String)
        Public Property WorkbookPath As System.Func(Of System.String)
    End Class

    Public NotInheritable Class ReviewContract
        Private Sub New()
        End Sub

        Public Shared ReadOnly JsonSettings As New Newtonsoft.Json.JsonSerializerSettings With {
            .TypeNameHandling = Newtonsoft.Json.TypeNameHandling.None,
            .MaxDepth = 64,
            .ContractResolver = New Newtonsoft.Json.Serialization.DefaultContractResolver With {
                .NamingStrategy = New Newtonsoft.Json.Serialization.CamelCaseNamingStrategy With {.ProcessDictionaryKeys = False, .OverrideSpecifiedNames = False}
            },
            .DateParseHandling = Newtonsoft.Json.DateParseHandling.None
        }

        Public Shared Function ShortDocumentReference(value As System.String, identity As System.String) As System.String
            Dim text As System.String = If(value, "")
            Const maximum As System.Int32 = 48
            If text.Length <= maximum Then Return text
            Dim suffix As System.String = If(identity, "").Replace("-", "")
            If suffix.Length < 8 Then suffix = Digest(text)
            suffix = suffix.Substring(0, 8)
            Dim length As System.Int32 = maximum - suffix.Length - 3
            If System.Char.IsHighSurrogate(text(length - 1)) Then length -= 1
            Dim prefix As System.String = text.Substring(0, length)
            Dim boundary As System.Int32 = prefix.LastIndexOfAny(New System.Char() {" "c, "/"c, "\"c, "_"c, "-"c, "."c})
            If boundary >= length \ 2 Then prefix = prefix.Substring(0, boundary)
            Return prefix.TrimEnd() & " … " & suffix
        End Function

        Public Shared Function Serialize(value As System.Object) As System.String
            Return Newtonsoft.Json.JsonConvert.SerializeObject(value, Newtonsoft.Json.Formatting.None, JsonSettings)
        End Function

        Public Shared Function Clone(project As ReviewProject) As ReviewProject
            Return Deserialize(Serialize(project))
        End Function

        Public Shared Function Deserialize(json As System.String) As ReviewProject
            Dim project As ReviewProject = Newtonsoft.Json.JsonConvert.DeserializeObject(Of ReviewProject)(json, JsonSettings)
            Validate(project)
            Return project
        End Function

        Public Shared Function Digest(value As System.String) As System.String
            Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Return System.BitConverter.ToString(algorithm.ComputeHash(System.Text.Encoding.UTF8.GetBytes(If(value, "")))).Replace("-", "").ToLowerInvariant()
            End Using
        End Function

        Public Shared Sub Validate(project As ReviewProject)
            If project Is Nothing OrElse (project.SchemaVersion <> 1 AndAlso project.SchemaVersion <> 2 AndAlso project.SchemaVersion <> 3) Then Throw New System.IO.InvalidDataException("Unsupported or missing Tabular Review data. Keep the workbook and open it with a compatible Red Ink version.")
            ' Version 3 retains authored prompts, completed working notes and removal state.
            ' Older supported data migrates without deleting its sources or findings.
            project.SchemaVersion = 3
            RequireId(project.Id)
            If project.Criteria Is Nothing OrElse project.Documents Is Nothing OrElse project.Audit Is Nothing OrElse project.View Is Nothing Then Throw New System.IO.InvalidDataException("Incomplete review data.")
            Dim criterionIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each criterion As ReviewCriterion In project.Criteria
                If criterion Is Nothing Then Throw New System.IO.InvalidDataException("Missing criterion.")
                RequireId(criterion.Id)
                If Not criterionIds.Add(criterion.Id) Then Throw New System.IO.InvalidDataException("Duplicate criterion identity.")
                If System.String.IsNullOrWhiteSpace(criterion.Title) OrElse criterion.Revision < 1 OrElse criterion.Options Is Nothing Then Throw New System.IO.InvalidDataException("Every criterion needs a title, revision and options list.")
                If Not Contains(New System.String() {"analysis", "text", "priority", "flag"}, criterion.Kind) Then Throw New System.IO.InvalidDataException("Unknown column type.")
                If criterion.Kind = "analysis" AndAlso System.String.IsNullOrWhiteSpace(criterion.Instruction) Then Throw New System.IO.InvalidDataException("Every analysis criterion needs its detailed question.")
            Next
            Dim documentIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each document As ReviewDocument In project.Documents
                If document Is Nothing Then Throw New System.IO.InvalidDataException("Missing document.")
                RequireId(document.Id)
                If Not documentIds.Add(document.Id) Then Throw New System.IO.InvalidDataException("Duplicate document identity.")
                If document.Segments Is Nothing OrElse document.SourceHistory Is Nothing OrElse document.Answers Is Nothing OrElse document.Flags Is Nothing Then Throw New System.IO.InvalidDataException("Incomplete document data.")
                If Not Contains(New System.String() {"unreviewed", "in_review", "reviewed"}, document.ReviewStatus) Then Throw New System.IO.InvalidDataException("Unknown document review status. Use unreviewed, in_review or reviewed.")
                If Not Contains(New System.String() {"pending", "complete", "partial", "error", "missing", "changed", "imported"}, document.ExtractionStatus) Then Throw New System.IO.InvalidDataException("Unknown source extraction status.")
                If System.String.IsNullOrEmpty(document.StoredExtractionStatus) AndAlso (document.ExtractionStatus = "complete" OrElse document.ExtractionStatus = "partial") Then
                    document.StoredExtractionStatus = document.ExtractionStatus
                    document.StoredExtractionMessage = document.ExtractionMessage
                End If
                If Not Contains(New System.String() {"", "complete", "partial"}, If(document.StoredExtractionStatus, "")) Then Throw New System.IO.InvalidDataException("Unknown stored extraction coverage.")
                Dim segmentIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                For Each segment As ReviewSourceSegment In document.Segments
                    If segment Is Nothing OrElse System.String.IsNullOrEmpty(segment.Id) OrElse segment.Text Is Nothing OrElse Not segmentIds.Add(segment.Id) Then Throw New System.IO.InvalidDataException("Invalid source segments.")
                Next
                For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, ReviewAnswer) In document.Answers
                    If Not criterionIds.Contains(pair.Key) OrElse pair.Value Is Nothing Then Throw New System.IO.InvalidDataException("An answer has no matching criterion.")
                    Dim answer As ReviewAnswer = pair.Value
                    If Not Contains(New System.String() {"pending", "draft", "reviewed", "stale", "error", "imported"}, answer.Status) Then Throw New System.IO.InvalidDataException("Unknown answer status.")
                    If answer.Text Is Nothing OrElse answer.Evidence Is Nothing OrElse answer.History Is Nothing OrElse answer.FollowUps Is Nothing Then Throw New System.IO.InvalidDataException("Incomplete answer data.")
                    For Each evidence As ReviewEvidence In answer.Evidence
                        If evidence Is Nothing Then Throw New System.IO.InvalidDataException("Missing evidence.")
                        RequireId(evidence.Id)
                        If evidence.MatchStatus = "exact" OrElse evidence.MatchStatus = "whitespace" Then
                            Dim segment As ReviewSourceSegment = document.Segments.Find(Function(item As ReviewSourceSegment) item.Id = evidence.SegmentId)
                            If segment Is Nothing OrElse evidence.Start < 0 OrElse evidence.Length <= 0 OrElse CLng(evidence.Start) + evidence.Length > segment.Text.Length Then Throw New System.IO.InvalidDataException("An evidence anchor falls outside its stored source.")
                            If Not System.String.Equals(segment.Text.Substring(evidence.Start, evidence.Length), evidence.Quote, System.StringComparison.Ordinal) Then Throw New System.IO.InvalidDataException("A verified quotation differs from its stored source span.")
                        ElseIf evidence.MatchStatus <> "unverified" Then
                            Throw New System.IO.InvalidDataException("Unknown evidence validation state.")
                        End If
                    Next
                    For Each version As ReviewAnswerVersion In answer.History
                        If version Is Nothing OrElse version.Text Is Nothing OrElse version.Evidence Is Nothing Then Throw New System.IO.InvalidDataException("Incomplete answer history.")
                    Next
                    For Each turn As ReviewTurn In answer.FollowUps
                        If turn Is Nothing OrElse turn.Instruction Is Nothing OrElse turn.Answer Is Nothing Then Throw New System.IO.InvalidDataException("Incomplete follow-up history.")
                    Next
                Next
                For Each version As ReviewSourceVersion In document.SourceHistory
                    If version Is Nothing OrElse version.Segments Is Nothing OrElse System.String.IsNullOrEmpty(version.SourceHash) Then Throw New System.IO.InvalidDataException("Incomplete source history.")
                Next
            Next
            If project.ReportFiles Is Nothing Then project.ReportFiles = New System.Collections.Generic.List(Of ReviewReportFile)()
            For Each report As ReviewReportFile In project.ReportFiles
                If report Is Nothing OrElse System.String.IsNullOrWhiteSpace(report.Path) OrElse Not Contains(New System.String() {"issues", "qa", "report", "matrix_xlsx", "matrix_docx", "document_report"}, report.Kind) Then Throw New System.IO.InvalidDataException("Invalid exported-report reference.")
                If report.RelativePath Is Nothing Then report.RelativePath = ""
                If report.RelativePath.Length > 0 AndAlso Not IsSafeRelativePath(report.RelativePath) Then Throw New System.IO.InvalidDataException("Invalid relative report path.")
            Next
            If project.ReportDefinitions Is Nothing Then project.ReportDefinitions = New System.Collections.Generic.List(Of ReviewReportDefinition)()
            ValidateReportDefinitions(project.ReportDefinitions, criterionIds)
            NormalizeFlagColors(project)
            ValidateFindings(project)
            If project.SourceRootRelativePath Is Nothing Then project.SourceRootRelativePath = ""
            project.SourceRootRelativePath = project.SourceRootRelativePath.Replace("\", "/").Trim()
            If project.SourceRootRelativePath.Length > 0 AndAlso Not IsSafeRelativePath(project.SourceRootRelativePath) Then Throw New System.IO.InvalidDataException("The portable source root must stay inside the workbook folder.")
            NormalizeView(project)
        End Sub

        ' Matrix and Criteria share one order after the explicit version-3 migration.
        ' Accept older/missing preferences and normalize bounded values deterministically.
        Public Shared Sub ValidateReportDefinitions(definitions As System.Collections.Generic.List(Of ReviewReportDefinition),
                                                    criterionIds As System.Collections.Generic.IEnumerable(Of System.String))
            If definitions Is Nothing Then Throw New System.IO.InvalidDataException("Report definitions must be an array.")
            Dim allowed As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal) From {
                "number", "document", "location", "status", "priority", "flags", "reviewer", "notes", "sourceStatus"}
            For Each id As System.String In criterionIds
                For Each prefix As System.String In New System.String() {"answer:", "assessment:", "answerStatus:", "evidence:"}
                    allowed.Add(prefix & id)
                Next
            Next
            Dim ids As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each definition As ReviewReportDefinition In definitions
                If definition Is Nothing OrElse System.String.IsNullOrWhiteSpace(definition.Id) OrElse Not ids.Add(definition.Id) OrElse
                   System.String.IsNullOrWhiteSpace(definition.Name) Then Throw New System.IO.InvalidDataException("Give each report a name and a unique identity.")
                If Not Contains(New System.String() {"matrix_xlsx", "matrix_docx", "document_report"}, definition.Kind) OrElse
                   Not Contains(New System.String() {"portrait", "landscape"}, definition.Orientation) Then Throw New System.IO.InvalidDataException("Choose a supported report format and orientation.")
                If Not Contains(New System.String() {"A4", "A3"}, definition.PaperSize) Then Throw New System.IO.InvalidDataException("Choose A4 or A3 for the report.")
                If definition.Title Is Nothing Then definition.Title = ""
                If Not Contains(New System.String() {"filtered", "selected", "all"}, definition.DocumentScope) Then Throw New System.IO.InvalidDataException("Choose a supported report document scope.")
                If Not Contains(New System.String() {"redink", "neutral"}, definition.Design) Then Throw New System.IO.InvalidDataException("Choose Red Ink or Neutral as the report design.")
                If definition.Fields Is Nothing OrElse definition.Fields.Count = 0 Then Throw New System.IO.InvalidDataException("Choose at least one report field.")
                Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                For Each key As System.String In definition.Fields
                    If key Is Nothing OrElse Not allowed.Contains(key) OrElse Not seen.Add(key) Then Throw New System.IO.InvalidDataException("A report contains an unknown or repeated field. Review its Columns settings.")
                Next
            Next
        End Sub

        Public Shared Sub NormalizeView(project As ReviewProject)
            Dim view As ReviewViewState = project.View
            If view Is Nothing Then Throw New System.IO.InvalidDataException("Missing review view.")
            If Not Contains(New System.String() {"light", "dark"}, view.Theme) OrElse
               Not Contains(New System.String() {"review", "matrix", "activity"}, view.Mode) OrElse
               Not Contains(New System.String() {"left", "right"}, view.ViewerSide) Then
                Throw New System.IO.InvalidDataException("Invalid review view.")
            End If
            If System.Double.IsNaN(view.ViewerWidth) OrElse System.Double.IsInfinity(view.ViewerWidth) Then view.ViewerWidth = 40.0R
            view.ViewerWidth = System.Math.Max(25.0R, System.Math.Min(65.0R, view.ViewerWidth))
            Dim validColumns As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal) From {"document"}
            For Each criterion As ReviewCriterion In project.Criteria
                If Not criterion.Archived Then validColumns.Add(criterion.Id)
            Next
            Dim sortColumn As System.String = If(view.SortColumn, "")
            Dim sortDirection As System.String = If(view.SortDirection, "")
            Dim sortable As System.Boolean = sortColumn = "document" OrElse project.Criteria.Exists(
                Function(criterion As ReviewCriterion) Not criterion.Archived AndAlso criterion.Id = sortColumn AndAlso
                    (criterion.Kind = "flag" OrElse criterion.Kind = "priority"))
            If Not sortable OrElse (sortDirection <> "asc" AndAlso sortDirection <> "desc") Then
                view.SortColumn = ""
                view.SortDirection = ""
            Else
                view.SortColumn = sortColumn
                view.SortDirection = sortDirection
            End If
            Dim order As New System.Collections.Generic.List(Of System.String)()
            If view.ColumnOrder IsNot Nothing Then
                For Each key As System.String In view.ColumnOrder
                    If key IsNot Nothing AndAlso key <> "document" AndAlso validColumns.Contains(key) AndAlso Not order.Contains(key) Then order.Add(key)
                Next
            End If
            For Each criterion As ReviewCriterion In project.Criteria
                If Not criterion.Archived AndAlso Not order.Contains(criterion.Id) Then order.Add(criterion.Id)
            Next
            If view.LayoutVersion >= 3 Then
                order.Clear()
                For Each criterion As ReviewCriterion In project.Criteria
                    If Not criterion.Archived Then order.Add(criterion.Id)
                Next
            End If
            view.ColumnOrder = order
            Dim widths As New System.Collections.Generic.Dictionary(Of System.String, System.Int32)(System.StringComparer.Ordinal)
            If view.ColumnWidths IsNot Nothing Then
                For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Int32) In view.ColumnWidths
                    If validColumns.Contains(pair.Key) Then
                        widths(pair.Key) = If(pair.Key = "document", System.Math.Max(160, System.Math.Min(600, pair.Value)), System.Math.Max(140, System.Math.Min(1200, pair.Value)))
                    End If
                Next
            End If
            view.ColumnWidths = widths
            Dim heights As New System.Collections.Generic.Dictionary(Of System.String, System.Int32)(System.StringComparer.Ordinal)
            If view.RowHeights IsNot Nothing Then
                For Each document As ReviewDocument In project.Documents
                    Dim height As System.Int32
                    If Not document.Archived AndAlso view.RowHeights.TryGetValue(document.Id, height) Then heights(document.Id) = System.Math.Max(64, System.Math.Min(900, height))
                Next
            End If
            view.RowHeights = heights
        End Sub

        Public Shared Function IsSafeRelativePath(value As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(value) OrElse System.IO.Path.IsPathRooted(value) OrElse value.IndexOf(":"c) >= 0 Then Return False
            If value.StartsWith("/", System.StringComparison.Ordinal) OrElse value.StartsWith("\", System.StringComparison.Ordinal) Then Return False
            For Each part As System.String In value.Replace("\", "/").Split("/"c)
                If part = ".." Then Return False
            Next
            Return value.IndexOfAny(System.IO.Path.GetInvalidPathChars()) < 0
        End Function

        Public Shared Sub NormalizeFlagColors(project As ReviewProject)
            Dim colors As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
            If project.FlagColors IsNot Nothing Then
                For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In project.FlagColors
                    Dim name As System.String = If(pair.Key, "")
                    Dim color As System.String = If(pair.Value, "").Trim()
                    If System.String.IsNullOrWhiteSpace(name) OrElse color.Length = 0 Then Continue For
                    If System.Text.RegularExpressions.Regex.IsMatch(color, "^#[0-9a-fA-F]{3}$") Then
                        color = "#" & New System.String(color(1), 2) & New System.String(color(2), 2) & New System.String(color(3), 2)
                    End If
                    If Not System.Text.RegularExpressions.Regex.IsMatch(color, "^#[0-9a-fA-F]{6}$") Then Throw New System.IO.InvalidDataException("Choose a valid flag color.")
                    colors(name) = color.ToUpperInvariant()
                Next
            End If
            ' Recover only an unambiguous application of the previous serializer's
            ' exact key transformation. Never infer a label from case-insensitive text.
            If project.FlagColorKeyVersion = 0 Then
                Dim labels As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                For Each document As ReviewDocument In project.Documents
                    For Each label As System.String In document.Flags
                        labels.Add(label)
                    Next
                Next
                For Each criterion As ReviewCriterion In project.Criteria
                    If criterion.Kind <> "flag" Then Continue For
                    For Each label As System.String In criterion.Options
                        labels.Add(label)
                    Next
                    For Each document As ReviewDocument In project.Documents
                        Dim answer As ReviewAnswer = Nothing
                        If document.Answers.TryGetValue(criterion.Id, answer) AndAlso Not System.String.IsNullOrWhiteSpace(answer.Text) Then labels.Add(answer.Text)
                    Next
                Next
                Dim legacy As New Newtonsoft.Json.Serialization.CamelCaseNamingStrategy With {.ProcessDictionaryKeys = True}
                Dim restored As New System.Collections.Generic.Dictionary(Of System.String, System.String)(colors, System.StringComparer.Ordinal)
                For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In colors
                    If labels.Contains(pair.Key) Then Continue For
                    Dim matches As New System.Collections.Generic.List(Of System.String)()
                    For Each label As System.String In labels
                        If legacy.GetDictionaryKey(label) = pair.Key Then matches.Add(label)
                    Next
                    If matches.Count = 1 AndAlso Not colors.ContainsKey(matches(0)) Then
                        restored.Remove(pair.Key)
                        restored(matches(0)) = pair.Value
                    End If
                Next
                colors = restored
                project.FlagColorKeyVersion = 1
            End If
            project.FlagColors = colors
        End Sub

        Private Shared Sub ValidateFindings(project As ReviewProject)
            If project.Findings Is Nothing Then project.Findings = New System.Collections.Generic.List(Of ReviewFinding)()
            If project.ConsistencyRuns Is Nothing Then project.ConsistencyRuns = New System.Collections.Generic.List(Of ReviewConsistencyRun)()
            Dim ids As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each finding As ReviewFinding In project.Findings
                If finding Is Nothing Then Throw New System.IO.InvalidDataException("Missing review finding.")
                RequireId(finding.Id)
                RequireId(finding.RunId)
                If Not ids.Add(finding.Id) Then Throw New System.IO.InvalidDataException("Duplicate finding identity.")
                If System.String.IsNullOrWhiteSpace(finding.Title) OrElse System.String.IsNullOrWhiteSpace(finding.Description) Then Throw New System.IO.InvalidDataException("Every finding needs its title and description.")
                If Not Contains(New System.String() {"potential_conflict", "ambiguity", "apparent_change", "information_gap", "review_finding"}, finding.Kind) Then Throw New System.IO.InvalidDataException("Unknown finding kind.")
                If Not Contains(New System.String() {"unspecified", "low", "medium", "high"}, finding.Priority) Then Throw New System.IO.InvalidDataException("Unknown finding priority.")
                If Not Contains(New System.String() {"draft", "reviewed", "dismissed", "stale"}, finding.Status) Then Throw New System.IO.InvalidDataException("Unknown finding status.")
                ValidateFindingEvidence(project, finding.Kind, finding.NoRelevantEvidence, finding.Evidence)
                ValidateFindingContext(project, finding.Context)
                If finding.History Is Nothing Then Throw New System.IO.InvalidDataException("Missing finding history.")
                For Each version As ReviewFindingVersion In finding.History
                    If version Is Nothing Then Throw New System.IO.InvalidDataException("Missing finding version.")
                    ValidateFindingEvidence(project, version.Kind, version.NoRelevantEvidence, version.Evidence)
                    ValidateFindingContext(project, version.Context)
                Next
            Next
            Dim runIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each run As ReviewConsistencyRun In project.ConsistencyRuns
                If run Is Nothing Then Throw New System.IO.InvalidDataException("Missing consistency run.")
                RequireId(run.Id)
                If Not runIds.Add(run.Id) Then Throw New System.IO.InvalidDataException("Duplicate consistency-run identity.")
                If Not Contains(New System.String() {"running", "complete", "partial", "cancelled", "error"}, run.Status) Then Throw New System.IO.InvalidDataException("Unknown consistency-run status.")
                If Not Contains(New System.String() {"comparison", "prompt"}, run.Mode) Then Throw New System.IO.InvalidDataException("Unknown cross-document review mode.")
                If run.Mode = "prompt" AndAlso System.String.IsNullOrWhiteSpace(run.Instruction) Then Throw New System.IO.InvalidDataException("A combined review needs its original prompt.")
                If run.WorkingNotes Is Nothing OrElse run.CompletedSourcePasses < 0 OrElse run.PlannedSourcePasses < run.CompletedSourcePasses OrElse run.CompletedSynthesisPasses < 0 OrElse run.PlannedSynthesisPasses < run.CompletedSynthesisPasses Then Throw New System.IO.InvalidDataException("Invalid combined-review coverage.")
                For Each note As ReviewPromptNote In run.WorkingNotes
                    If note Is Nothing OrElse Not Contains(New System.String() {"source", "synthesis"}, note.Stage) OrElse System.String.IsNullOrWhiteSpace(note.Description) Then Throw New System.IO.InvalidDataException("Invalid combined-review working note.")
                    ValidateFindingEvidence(project, "review_finding", note.NoRelevantEvidence, note.Evidence)
                Next
                If run.Documents Is Nothing OrElse run.FindingIds Is Nothing OrElse run.CompletedComparisons < 0 OrElse run.PlannedComparisons < run.CompletedComparisons OrElse run.CompletedContextChecks < 0 OrElse run.PlannedContextChecks < run.CompletedContextChecks Then Throw New System.IO.InvalidDataException("Invalid consistency-run coverage.")
                Dim covered As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                For Each document As ReviewConsistencyDocumentCoverage In run.Documents
                    If document Is Nothing OrElse Not project.Documents.Exists(Function(item As ReviewDocument) item.Id = document.DocumentId) OrElse Not covered.Add(document.DocumentId) Then Throw New System.IO.InvalidDataException("Unknown or duplicate consistency-run document.")
                    If Not Contains(New System.String() {"included", "unavailable"}, document.Status) OrElse document.SourceCharacters < 0 OrElse document.SourceParts < 0 Then Throw New System.IO.InvalidDataException("Invalid source coverage.")
                Next
                Dim found As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                For Each id As System.String In run.FindingIds
                    If Not ids.Contains(id) OrElse Not found.Add(id) OrElse Not project.Findings.Exists(Function(item As ReviewFinding) item.Id = id AndAlso item.RunId = run.Id) Then Throw New System.IO.InvalidDataException("Invalid finding membership in a consistency run.")
                Next
            Next
            For Each finding As ReviewFinding In project.Findings
                If Not runIds.Contains(finding.RunId) Then Throw New System.IO.InvalidDataException("A finding has no consistency run.")
            Next
        End Sub

        Private Shared Sub ValidateFindingEvidence(project As ReviewProject, kind As System.String, noRelevantEvidence As System.Boolean,
                                                   evidence As System.Collections.Generic.List(Of ReviewCrossDocumentEvidence))
            If noRelevantEvidence Then
                If kind <> "review_finding" OrElse evidence Is Nothing OrElse evidence.Count <> 0 Then Throw New System.IO.InvalidDataException("A negative combined review must retain an explicit empty evidence list.")
            Else
                ValidateCrossEvidence(project, evidence, kind <> "review_finding")
            End If
        End Sub

        Private Shared Sub ValidateFindingContext(project As ReviewProject, contexts As System.Collections.Generic.List(Of ReviewFindingContext))
            If contexts Is Nothing Then Throw New System.IO.InvalidDataException("Missing finding context.")
            For Each context As ReviewFindingContext In contexts
                If context Is Nothing OrElse System.String.IsNullOrWhiteSpace(context.Description) OrElse Not Contains(New System.String() {"supports", "qualifies", "apparent_change", "ambiguity"}, context.Relationship) Then Throw New System.IO.InvalidDataException("Invalid finding context.")
                ValidateCrossEvidence(project, context.Evidence, False)
            Next
        End Sub

        Public Shared Sub ValidateCrossEvidence(project As ReviewProject, citations As System.Collections.Generic.List(Of ReviewCrossDocumentEvidence), requireTwoDocuments As System.Boolean)
            If citations Is Nothing OrElse citations.Count = 0 Then Throw New System.IO.InvalidDataException("A finding needs verified source quotations.")
            Dim documentIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each citation As ReviewCrossDocumentEvidence In citations
                If citation Is Nothing OrElse citation.Evidence Is Nothing Then Throw New System.IO.InvalidDataException("Missing finding evidence.")
                Dim document As ReviewDocument = project.Documents.Find(Function(item As ReviewDocument) item.Id = citation.DocumentId)
                If document Is Nothing OrElse System.String.IsNullOrEmpty(citation.SourceHash) Then Throw New System.IO.InvalidDataException("A finding cites an unknown source.")
                Dim segments As System.Collections.Generic.List(Of ReviewSourceSegment) = Nothing
                If document.SourceHash = citation.SourceHash AndAlso document.TextHash = citation.TextHash Then
                    segments = document.Segments
                Else
                    Dim version As ReviewSourceVersion = document.SourceHistory.Find(Function(item As ReviewSourceVersion) item.SourceHash = citation.SourceHash AndAlso item.TextHash = citation.TextHash)
                    If version IsNot Nothing Then segments = version.Segments
                End If
                If segments Is Nothing Then Throw New System.IO.InvalidDataException("A finding cites a source snapshot that this review does not retain.")
                Dim evidence As ReviewEvidence = citation.Evidence
                RequireId(evidence.Id)
                Dim segment As ReviewSourceSegment = segments.Find(Function(item As ReviewSourceSegment) item.Id = evidence.SegmentId)
                If segment Is Nothing OrElse (evidence.MatchStatus <> "exact" AndAlso evidence.MatchStatus <> "whitespace") OrElse evidence.Start < 0 OrElse evidence.Length <= 0 OrElse CLng(evidence.Start) + evidence.Length > segment.Text.Length Then Throw New System.IO.InvalidDataException("Invalid finding source span.")
                If evidence.Locator <> segment.Locator OrElse Not System.String.Equals(segment.Text.Substring(evidence.Start, evidence.Length), evidence.Quote, System.StringComparison.Ordinal) Then Throw New System.IO.InvalidDataException("A finding quotation differs from its stored source span.")
                documentIds.Add(document.Id)
            Next
            If requireTwoDocuments AndAlso documentIds.Count < 2 Then Throw New System.IO.InvalidDataException("A cross-document finding needs verified quotations from two different documents.")
        End Sub

        Public Shared Sub ArchiveFinding(finding As ReviewFinding, actor As System.String)
            Dim version As ReviewFindingVersion = Newtonsoft.Json.JsonConvert.DeserializeObject(Of ReviewFindingVersion)(Serialize(finding), JsonSettings)
            version.Actor = actor
            finding.History.Add(version)
        End Sub

        Public Shared Sub RequireId(value As System.String)
            Dim parsed As System.Guid
            If System.String.IsNullOrEmpty(value) OrElse Not System.Guid.TryParseExact(value, "N", parsed) Then Throw New System.IO.InvalidDataException("Invalid review identity.")
        End Sub

        Public Shared Function Contains(values As System.Collections.Generic.IEnumerable(Of System.String), value As System.String) As System.Boolean
            For Each candidate As System.String In values
                If System.String.Equals(candidate, value, System.StringComparison.Ordinal) Then Return True
            Next
            Return False
        End Function

        Public Shared Sub ArchiveAnswer(answer As ReviewAnswer, actor As System.String)
            If answer Is Nothing Then Return
            If System.String.IsNullOrEmpty(answer.Text) AndAlso System.String.IsNullOrEmpty(answer.Assessment) AndAlso System.String.IsNullOrEmpty(answer.Error) AndAlso answer.Evidence.Count = 0 AndAlso Not answer.UserEdited AndAlso answer.Status = "pending" AndAlso System.String.IsNullOrEmpty(answer.UpdatedUtc) Then Return
            answer.History.Add(New ReviewAnswerVersion With {
                .Text = answer.Text, .Assessment = answer.Assessment, .Status = answer.Status,
                .UpdatedUtc = answer.UpdatedUtc, .UpdatedBy = answer.UpdatedBy, .ReviewedBy = answer.ReviewedBy, .ReviewedUtc = answer.ReviewedUtc, .Actor = actor, .ModelKey = answer.ModelKey,
                .SourceHash = answer.SourceHash, .CriterionRevision = answer.CriterionRevision,
                .UserEdited = answer.UserEdited, .Error = answer.Error,
                .Evidence = Newtonsoft.Json.JsonConvert.DeserializeObject(Of System.Collections.Generic.List(Of ReviewEvidence))(Serialize(answer.Evidence), JsonSettings)
            })
        End Sub

        Public Shared Sub AddAudit(project As ReviewProject, action As System.String, detail As System.String,
                                  Optional documentId As System.String = "", Optional criterionId As System.String = "", Optional actor As System.String = "")
            project.Audit.Add(New ReviewAuditEntry With {.Action = action, .Detail = detail, .DocumentId = documentId, .CriterionId = criterionId, .Actor = actor})
            project.UpdatedUtc = System.DateTime.UtcNow.ToString("o")
        End Sub

        Public Shared Sub Invalidate(document As ReviewDocument, Optional criterionId As System.String = "")
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, ReviewAnswer) In document.Answers
                If criterionId.Length = 0 OrElse pair.Key = criterionId Then pair.Value.Status = "stale"
            Next
            document.ReviewStatus = "unreviewed"
            document.LastReviewedUtc = ""
        End Sub
    End Class
End Namespace
