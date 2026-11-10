' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. For license to use see https://redink.ai.
' Render a detached, host-validated selection as a separate Open XML report.
' This service never opens Office COM or writes a review workbook/source document.
Option Explicit On
Option Strict On

Namespace Global.SharedLibrary.TabularReview
    Public NotInheritable Class TabularReviewReportService
        Private Const SpreadsheetNamespace As System.String = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        Private Const WordNamespace As System.String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        Private Const RelationshipNamespace As System.String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        Private Const XmlNamespace As System.String = "http://www.w3.org/XML/1998/namespace"
        Private Const ExcelPrintSideMarginInches As System.Double = 0.3R
        Private Const MaximumCellCharacters As System.Int32 = 32767
        Private Const ContinuationCharacters As System.Int32 = 30000
        Private Const ExcelContinuationCharacters As System.Int32 = 1000
        Private Const ExcelContinuationLines As System.Int32 = 18
        Private Const ExcelDisplayCharactersPerLine As System.Int32 = 40
        Private Const MaximumSheetRows As System.Int32 = 1048576

        Private Sub New()
        End Sub

        Public Shared Function Export(data As ReviewReportData, destinationPath As System.String,
                                      overwrite As System.Boolean,
                                      protectedPaths As System.Collections.Generic.IEnumerable(Of System.String),
                                      cancellationToken As System.Threading.CancellationToken) As ReviewReportExportResult
            cancellationToken.ThrowIfCancellationRequested()
            ValidateData(data, cancellationToken)
            Dim fullPath As System.String = System.IO.Path.GetFullPath(destinationPath)
            Dim expectedExtension As System.String = If(IsWordReport(data.Kind), ".docx", ".xlsx")
            If Not System.String.Equals(System.IO.Path.GetExtension(fullPath), expectedExtension, System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.IO.InvalidDataException("Choose " & expectedExtension & " for this report format.")
            End If
            Dim directory As System.String = System.IO.Path.GetDirectoryName(fullPath)
            If Not System.IO.Directory.Exists(directory) Then Throw New System.IO.DirectoryNotFoundException("Choose an existing folder for the report.")
            Dim protectedFiles As New System.Collections.Generic.List(Of System.String)()
            If protectedPaths IsNot Nothing Then protectedFiles.AddRange(protectedPaths)
            For Each document As ReviewReportDocument In data.Documents
                cancellationToken.ThrowIfCancellationRequested()
                If Not System.String.IsNullOrWhiteSpace(document.SourcePath) Then protectedFiles.Add(document.SourcePath)
            Next
            EnsureIndependentDestination(fullPath, protectedFiles)
            Dim existed As System.Boolean = System.IO.File.Exists(fullPath)
            If existed AndAlso Not overwrite Then Throw New System.IO.IOException("This report filename already exists. Choose another filename or confirm replacement.")
            Dim previousHash As System.String = If(existed, HashFile(fullPath), "")
            Dim temporary As System.String = System.IO.Path.Combine(directory, ".redink-report-" & System.Guid.NewGuid().ToString("N") & ".tmp")
            Try
                Dim expectedDigest As System.String
                Using stream As New System.IO.FileStream(temporary, System.IO.FileMode.CreateNew, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
                    If IsWordReport(data.Kind) Then
                        expectedDigest = WriteWordReport(stream, data, fullPath, cancellationToken)
                    Else
                        expectedDigest = WriteExcelReport(stream, data, fullPath, cancellationToken)
                    End If
                    stream.Flush(True)
                End Using
                ValidatePackage(temporary, data.Kind, expectedDigest, cancellationToken)
                cancellationToken.ThrowIfCancellationRequested()
                EnsureIndependentDestination(fullPath, protectedFiles)
                If existed Then
                    If Not System.IO.File.Exists(fullPath) OrElse Not System.String.Equals(previousHash, HashFile(fullPath), System.StringComparison.Ordinal) Then
                        Throw New System.IO.IOException("The existing report changed during export. Choose another filename to keep both versions.")
                    End If
                    ' A failed replace keeps the previous file. Never delete it as
                    ' a fallback on a filesystem that cannot replace atomically.
                    System.IO.File.Replace(temporary, fullPath, Nothing)
                Else
                    ' Move refuses a filename another process created meanwhile.
                    System.IO.File.Move(temporary, fullPath)
                End If
                Dim evidenceCount As System.Int32 = 0
                For Each item As ReviewReportItem In data.Items
                    evidenceCount += item.Evidence.Count
                Next
                System.Diagnostics.Trace.WriteLine("[TabularReview:report_export] kind=" & data.Kind & " | items=" & data.Items.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " | evidence=" & evidenceCount.ToString(System.Globalization.CultureInfo.InvariantCulture))
                Return New ReviewReportExportResult With {
                    .Path = fullPath, .ItemCount = If(IsMatrixReport(data.Kind), data.Rows.Count, data.Items.Count), .EvidenceCount = evidenceCount,
                    .Message = If(IsMatrixReport(data.Kind), "Exported " & data.Rows.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " documents in your selected field order.", "Exported " & data.Items.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " review items and " & evidenceCount.ToString(System.Globalization.CultureInfo.InvariantCulture) & " source quotations.")}
            Finally
                Try
                    If System.IO.File.Exists(temporary) Then System.IO.File.Delete(temporary)
                Catch cleanupFailure As System.Exception
                    System.Diagnostics.Trace.WriteLine("[TabularReview:report_cleanup_failed] " & cleanupFailure.Message)
                End Try
            End Try
        End Function

        Private Shared Sub MarkReport(part As DocumentFormat.OpenXml.Packaging.CustomFilePropertiesPart)
            Dim properties As New DocumentFormat.OpenXml.CustomProperties.Properties()
            Dim marker As New DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty With {
                .FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", .PropertyId = 2,
                .Name = ReviewSourceService.ReportArtifactProperty}
            marker.Append(New DocumentFormat.OpenXml.VariantTypes.VTLPWSTR(ReviewSourceService.ReportArtifactValue))
            properties.Append(marker)
            part.Properties = properties
            part.Properties.Save()
        End Sub

        Private Shared Function IsWordReport(kind As System.String) As System.Boolean
            Return kind = "report" OrElse kind = "matrix_docx" OrElse kind = "document_report"
        End Function

        Private Shared Function IsMatrixReport(kind As System.String) As System.Boolean
            Return kind = "matrix_xlsx" OrElse kind = "matrix_docx" OrElse kind = "document_report"
        End Function

        Private Shared Sub ValidateData(data As ReviewReportData, cancellationToken As System.Threading.CancellationToken)
            If data Is Nothing Then Throw New System.ArgumentNullException(NameOf(data))
            If Not IsMatrixReport(data.Kind) Then Throw New System.IO.InvalidDataException("Choose a configurable Matrix or Documents with cover report.")
            If data.PaperSize <> "A4" AndAlso data.PaperSize <> "A3" Then Throw New System.IO.InvalidDataException("Choose A4 or A3 for the report.")
            If data.Design <> "redink" AndAlso data.Design <> "neutral" Then Throw New System.IO.InvalidDataException("Choose Red Ink or Neutral as the report design.")
            If data.Items Is Nothing OrElse (Not IsMatrixReport(data.Kind) AndAlso data.Items.Count = 0) Then Throw New System.IO.InvalidDataException("Select at least one answer or cross-document finding to export.")
            If data.Documents Is Nothing Then Throw New System.IO.InvalidDataException("The report has no source-document register.")
            If data.ProjectRevision < 0 OrElse System.String.IsNullOrWhiteSpace(data.ProjectId) Then Throw New System.IO.InvalidDataException("The report has an invalid review identity or revision.")
            CheckText(data.Title, data.ProjectId, data.GeneratedUtc)
            Dim documents As New System.Collections.Generic.Dictionary(Of System.String, ReviewReportDocument)(System.StringComparer.Ordinal)
            For Each document As ReviewReportDocument In data.Documents
                cancellationToken.ThrowIfCancellationRequested()
                If document Is Nothing OrElse System.String.IsNullOrWhiteSpace(document.Id) OrElse documents.ContainsKey(document.Id) Then Throw New System.IO.InvalidDataException("The report source list contains a missing or repeated document. Reload the review and select the items again.")
                CheckText(document.Id, document.Number, document.FileName, document.RelativePath, document.SourcePath, document.SourceHash, document.ExtractionStatus, document.ExtractionMessage)
                documents.Add(document.Id, document)
            Next
            If IsMatrixReport(data.Kind) Then
                If data.Orientation <> "portrait" AndAlso data.Orientation <> "landscape" Then Throw New System.IO.InvalidDataException("Choose portrait or landscape.")
                If data.Columns Is Nothing OrElse data.Columns.Count = 0 OrElse data.Columns.Count > 16384 OrElse data.Rows Is Nothing OrElse data.Rows.Count = 0 Then Throw New System.IO.InvalidDataException("Choose report columns and documents within Office limits.")
                Dim keys As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                For Each column As ReviewReportColumn In data.Columns
                    If column Is Nothing OrElse System.String.IsNullOrEmpty(column.Key) OrElse Not keys.Add(column.Key) OrElse
                       System.Double.IsNaN(column.Width) OrElse System.Double.IsInfinity(column.Width) OrElse column.Width < 1 OrElse column.Width > 255 Then Throw New System.IO.InvalidDataException("Invalid report column.")
                    CheckText(column.Title)
                Next
                Dim rowIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                For Each row As ReviewReportRow In data.Rows
                    cancellationToken.ThrowIfCancellationRequested()
                    If row Is Nothing OrElse row.DocumentId Is Nothing OrElse Not documents.ContainsKey(row.DocumentId) OrElse Not rowIds.Add(row.DocumentId) OrElse
                       row.Values Is Nothing OrElse row.Values.Count <> data.Columns.Count Then Throw New System.IO.InvalidDataException("The report has inconsistent document rows.")
                    CheckText(row.Values.ToArray())
                Next
                If rowIds.Count <> documents.Count Then Throw New System.IO.InvalidDataException("The report has inconsistent document coverage.")
            End If
            Dim itemIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each item As ReviewReportItem In data.Items
                cancellationToken.ThrowIfCancellationRequested()
                If item Is Nothing OrElse System.String.IsNullOrWhiteSpace(item.Id) OrElse Not itemIds.Add(item.Id) Then Throw New System.IO.InvalidDataException("The report selection contains a missing or repeated item. Select each item once.")
                If item.Kind <> "answer" AndAlso item.Kind <> "contradiction" AndAlso item.Kind <> "finding" AndAlso item.Kind <> "review_finding" Then Throw New System.IO.InvalidDataException("The report contains an unsupported item type.")
                CheckText(item.Id, item.Title, item.Question, item.Answer, item.Assessment, item.Priority, item.Status, item.Reviewer, item.Notes, item.ModelKey,
                    item.RequestedAction, item.Context, item.Coverage, item.Prompt, item.ModelLabel)
                If item.DocumentIds Is Nothing OrElse item.Evidence Is Nothing Then Throw New System.IO.InvalidDataException("An exported item has an incomplete source-reference list.")
                Dim references As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                For Each id As System.String In item.DocumentIds
                    If id Is Nothing OrElse Not documents.ContainsKey(id) OrElse Not references.Add(id) Then Throw New System.IO.InvalidDataException("An exported item has an unknown or repeated source-document reference.")
                Next
                For Each evidence As ReviewReportEvidence In item.Evidence
                    If evidence Is Nothing OrElse evidence.DocumentId Is Nothing OrElse Not references.Contains(evidence.DocumentId) Then Throw New System.IO.InvalidDataException("A source quotation does not belong to its item's source-document list.")
                    CheckText(evidence.DocumentId, evidence.SegmentId, evidence.Locator, evidence.Quote, evidence.MatchStatus, evidence.SourceHash, evidence.Role)
                    If System.String.IsNullOrEmpty(evidence.Quote) Then Throw New System.IO.InvalidDataException("A report source quotation is empty.")
                    If evidence.MatchStatus <> "exact" AndAlso evidence.MatchStatus <> "whitespace" AndAlso evidence.MatchStatus <> "unverified" Then Throw New System.IO.InvalidDataException("A source quotation has an unknown verification status.")
                    If evidence.MatchStatus <> "unverified" AndAlso (evidence.Start < 0 OrElse evidence.Length <= 0 OrElse System.String.IsNullOrEmpty(evidence.SegmentId) OrElse System.String.IsNullOrEmpty(evidence.SourceHash)) Then
                        Throw New System.IO.InvalidDataException("The report cannot verify a stored source quotation. Refresh and review its source before exporting.")
                    End If
                Next
            Next
        End Sub

        Private Shared Sub CheckText(ParamArray values As System.String())
            For Each value As System.String In values
                Try
                    System.Xml.XmlConvert.VerifyXmlChars(If(value, ""))
                Catch failure As System.Xml.XmlException
                    Throw New System.IO.InvalidDataException("The selected text contains a character that Office cannot store in an XML report. Correct that text before exporting; no output file has changed.", failure)
                End Try
            Next
        End Sub

        Private Shared Sub EnsureIndependentDestination(path As System.String, protectedPaths As System.Collections.Generic.IEnumerable(Of System.String))
            Dim existingTarget As System.Boolean = System.IO.File.Exists(path)
            Dim checkedPaths As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each protectedPath As System.String In protectedPaths
                If System.String.IsNullOrWhiteSpace(protectedPath) OrElse Not checkedPaths.Add(protectedPath) Then Continue For
                If Not System.IO.Path.IsPathRooted(protectedPath) OrElse protectedPath.IndexOf("://", System.StringComparison.Ordinal) >= 0 Then
                    If existingTarget Then Throw New System.IO.IOException("Choose a new report filename. The workbook or a source uses a non-filesystem location, so export cannot verify that an existing target is independent of it.")
                    Continue For
                End If
                If System.String.Equals(path, System.IO.Path.GetFullPath(protectedPath), System.StringComparison.OrdinalIgnoreCase) Then
                    Throw New System.IO.IOException("Choose a new report filename. Export cannot replace the review workbook or an original source document.")
                End If
                If existingTarget Then
                    Dim samePhysicalFile As System.Boolean
                    Try
                        ' The shared physical-path helper resolves mapped drives,
                        ' UNC and short-name aliases without granting source access.
                        samePhysicalFile = Global.SharedLibrary.SharedLibrary.GeneratedOutputRegistry.IsPhysicalPathAtOrBelow(protectedPath, path)
                    Catch failure As System.Exception When TypeOf failure Is System.IO.IOException OrElse TypeOf failure Is System.UnauthorizedAccessException OrElse TypeOf failure Is System.ArgumentException OrElse TypeOf failure Is System.NotSupportedException
                        Throw New System.IO.IOException("Choose a new report filename. Export could not verify that this existing destination is independent of every protected source.", failure)
                    End Try
                    If samePhysicalFile Then Throw New System.IO.IOException("Choose a new report filename. This destination resolves to the review workbook or an original source document.")
                End If
            Next
            If existingTarget AndAlso (System.IO.File.GetAttributes(path) And System.IO.FileAttributes.ReparsePoint) <> 0 Then
                Throw New System.IO.IOException("Choose a new report filename instead of replacing a linked file.")
            End If
        End Sub

        Private Shared Function HashFile(path As System.String) As System.String
            Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                Using hash As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                    Return System.Convert.ToBase64String(hash.ComputeHash(stream))
                End Using
            End Using
        End Function

        Private Shared Function SourceLink(sourcePath As System.String, outputPath As System.String) As System.Uri
            If System.String.IsNullOrWhiteSpace(sourcePath) OrElse Not System.IO.Path.IsPathRooted(sourcePath) Then Return Nothing
            If sourcePath.StartsWith("\\?\", System.StringComparison.Ordinal) OrElse sourcePath.StartsWith("\\.\", System.StringComparison.Ordinal) Then Return Nothing
            Dim fullPath As System.String
            Dim target As System.Uri = Nothing
            Try
                fullPath = System.IO.Path.GetFullPath(sourcePath)
                If Not System.IO.File.Exists(fullPath) OrElse Not System.Uri.TryCreate(fullPath, System.UriKind.Absolute, target) OrElse Not target.IsFile Then Return Nothing
                Dim directory As System.String = System.IO.Path.GetDirectoryName(outputPath).TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar) & System.IO.Path.DirectorySeparatorChar
                If fullPath.StartsWith(directory, System.StringComparison.OrdinalIgnoreCase) Then
                    Return New System.Uri(directory, System.UriKind.Absolute).MakeRelativeUri(target)
                End If
                ' Reports carry portable source links only; external locations stay in the review.
                Return Nothing
            Catch failure As System.Exception When TypeOf failure Is System.ArgumentException OrElse TypeOf failure Is System.IO.IOException OrElse TypeOf failure Is System.NotSupportedException OrElse TypeOf failure Is System.UnauthorizedAccessException OrElse TypeOf failure Is System.UriFormatException
                System.Diagnostics.Trace.WriteLine("[TabularReview:report_source_link_skipped] " & failure.Message)
                Return Nothing
            End Try
        End Function

        Private Shared Sub WritePart(part As DocumentFormat.OpenXml.Packaging.OpenXmlPart, writeContent As System.Action(Of System.Xml.XmlWriter))
            Using stream As System.IO.Stream = part.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write)
                Using writer As System.Xml.XmlWriter = System.Xml.XmlWriter.Create(stream, New System.Xml.XmlWriterSettings With {
                    .Encoding = New System.Text.UTF8Encoding(False, True), .Indent = False, .NewLineHandling = System.Xml.NewLineHandling.Entitize, .CloseOutput = False})
                    writer.WriteStartDocument()
                    writeContent(writer)
                    writer.WriteEndDocument()
                End Using
            End Using
        End Sub

        Private NotInheritable Class ContentDigest
            Implements System.IDisposable
            Private ReadOnly _hash As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
            Public Sub Add(ParamArray values As System.String())
                For Each value As System.String In values
                    Dim bytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(If(value, ""))
                    Dim length As System.Byte() = System.BitConverter.GetBytes(bytes.Length)
                    _hash.TransformBlock(length, 0, length.Length, length, 0)
                    If bytes.Length > 0 Then _hash.TransformBlock(bytes, 0, bytes.Length, bytes, 0)
                Next
            End Sub
            Public Function Finish() As System.String
                _hash.TransformFinalBlock(New System.Byte() {}, 0, 0)
                Return System.Convert.ToBase64String(_hash.Hash)
            End Function
            Public Sub Dispose() Implements System.IDisposable.Dispose
                _hash.Dispose()
            End Sub
        End Class

        Private NotInheritable Class ReportCell
            Public Property Emphasis As System.Int32
            Public Property Text As System.String
            Public Property Link As System.Uri
            Public Property Location As System.String = ""
            Public Property Style As System.Int32 = -1
            Public Sub New(value As System.String, Optional hyperlink As System.Uri = Nothing)
                Text = If(value, "")
                Link = hyperlink
            End Sub
        End Class

        Private NotInheritable Class ReportSheet
            Public Property PaperSize As System.String = "A4"
            Public Property ColumnPixelWidths As System.Double()
            Public Property PixelWidth As System.Double
            Public Property AutoSize As System.Boolean
            Public Property PreserveCells As System.Boolean
            Public Property Name As System.String
            Public Property Title As System.String
            Public Property Description As System.String
            Public Property Headers As System.String()
            Public Property Widths As System.Double()
            Public Property FreezeColumns As System.Int32 = 1
            Public Property Filter As System.Boolean = True
            Public Property MaximumRowHeight As System.Double = 409.0R
            Public Property Rows As New System.Collections.Generic.List(Of System.Collections.Generic.List(Of ReportCell))()
        End Class

        Private NotInheritable Class LongText
            Public Property Key As System.String
            Public Property Origin As System.String
            Public Property OriginLocation As System.String
            Public Property OriginalLink As System.Uri
            Public Property OriginalLocation As System.String = ""
            Public Property FirstRow As System.Int32
            Public Property PartCount As System.Int32
            Public Property Text As System.String
        End Class

        Private NotInheritable Class ReportHyperlink
            Public Property Reference As System.String
            Public Property Relationship As System.String = ""
            Public Property Location As System.String = ""
        End Class

        Private Shared Function Cells(ParamArray values As System.String()) As System.Collections.Generic.List(Of ReportCell)
            Dim result As New System.Collections.Generic.List(Of ReportCell)()
            For Each value As System.String In values
                result.Add(New ReportCell(value))
            Next
            Return result
        End Function


        Private Shared Function ValueOr(value As System.String, fallback As System.String) As System.String
            Return If(System.String.IsNullOrWhiteSpace(value), fallback, value)
        End Function

        Private Shared Function ItemReference(index As System.Int32) As System.String
            Return "Item " & (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Shared Function DocumentLabel(document As ReviewReportDocument, index As System.Int32) As System.String
            Dim number As System.String = If(System.String.IsNullOrWhiteSpace(document.Number), "", document.Number & " · ")
            Return "Document " & (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) & " — " & number & ValueOr(document.FileName, "Unnamed document")
        End Function

        Private Shared Function ReviewStatus(value As System.String) As System.String
            Select Case If(value, "")
                Case "pending" : Return "Not yet analyzed"
                Case "draft" : Return "Needs review"
                Case "reviewed" : Return "Reviewed"
                Case "stale" : Return "Changed — review again"
                Case "error" : Return "Analysis incomplete"
                Case "imported" : Return "Imported — verify"
                Case "dismissed" : Return "Dismissed"
                Case Else : Return ValueOr(value, "Not recorded")
            End Select
        End Function

        Private Shared Function PriorityLabel(value As System.String) As System.String
            Select Case If(value, "")
                Case "unspecified", "" : Return "Not set"
                Case "high" : Return "High"
                Case "medium" : Return "Medium"
                Case "low" : Return "Low"
                Case Else : Return value
            End Select
        End Function

        Private Shared Function PriorityStyle(value As System.String) As System.Int32
            Select Case If(value, "")
                Case "high" : Return 9
                Case "medium" : Return 8
                Case "low" : Return 10
                Case Else : Return -1
            End Select
        End Function

        Private Shared Function SourceStatus(document As ReviewReportDocument) As System.String
            Select Case document.ExtractionStatus
                Case "complete" : Return "Complete saved extraction"
                Case "partial" : Return "Partial saved extraction"
                Case "missing" : Return "Original unavailable"
                Case "changed" : Return "Original changed"
                Case "error" : Return "Extraction failed"
                Case "imported" : Return "Imported record"
                Case Else : Return "Extraction not confirmed"
            End Select
        End Function

        Private Shared Function SourceNotice(document As ReviewReportDocument) As System.String
            Select Case document.ExtractionStatus
                Case "complete" : Return "The review used the saved extracted text. Check the quoted passages against the original when reviewing an item."
                Case "partial" : Return "The review covers only the available extracted text. Extract the remaining content and review it before treating the document as fully covered."
                Case "missing" : Return "The original is unavailable. This report retains the saved quotations; restore or relink the original and verify them before closing the review."
                Case "changed" : Return "The original differs from the saved text. Refresh the document and review the affected answers and quotations again."
                Case "error" : Return "Extraction did not finish. Open the original, resolve the extraction error in Tabular Review and repeat the affected review."
                Case "imported" : Return "The review contains an imported record. Link and extract the original to verify source coverage."
                Case Else : Return "The review has no confirmed extraction for this source. Extract the original and verify the affected answers."
            End Select
        End Function

        Private Shared Function SourceLocation(document As ReviewReportDocument) As System.String
            Dim relative As System.String = If(document.RelativePath, "")
            Dim inspected As System.String = relative.Trim()
            ' A report may show a path inside the source folder, never a private absolute location.
            If inspected.Length = 0 Then Return "Not recorded"
            If inspected.IndexOf(":"c) >= 0 OrElse inspected.StartsWith("\", System.StringComparison.Ordinal) OrElse
               inspected.StartsWith("/", System.StringComparison.Ordinal) Then Return "Outside the source folder"
            For Each part As System.String In inspected.Replace("\", "/").Split("/"c)
                If part = ".." Then Return "Outside the source folder"
            Next
            Return relative
        End Function

        Private Shared Function EvidenceStatus(evidence As ReviewReportEvidence, document As ReviewReportDocument) As System.String
            Dim status As System.String = If(evidence.MatchStatus = "exact", "Verified against saved text", If(evidence.MatchStatus = "whitespace", "Verified with whitespace normalization", "Quotation not verified"))
            If Not System.String.IsNullOrEmpty(evidence.SourceHash) AndAlso Not System.String.Equals(evidence.SourceHash, document.SourceHash, System.StringComparison.Ordinal) Then status &= " — earlier source version; review again"
            If document.ExtractionStatus = "missing" Then status &= " — original unavailable"
            If document.ExtractionStatus = "changed" Then status &= " — original changed"
            If document.ExtractionStatus = "partial" Then status &= " — partial source extraction"
            If document.ExtractionStatus = "error" Then status &= " — current extraction failed"
            Return status
        End Function

        Private Shared Function NextStep(item As ReviewReportItem) As System.String
            If Not System.String.IsNullOrWhiteSpace(item.RequestedAction) Then Return item.RequestedAction
            If item.Kind <> "answer" AndAlso Not System.String.IsNullOrWhiteSpace(item.Question) Then Return "Clarify: " & item.Question
            Return "No follow-up recorded"
        End Function

        Private Shared Function ReviewLine(item As ReviewReportItem) As System.String
            Return ReviewStatus(item.Status) & System.Environment.NewLine & "Reviewer: " & ValueOr(item.Reviewer, "Not assigned")
        End Function

        Private Shared Function GeneratedLabel(value As System.String) As System.String
            Dim parsed As System.DateTimeOffset
            If System.DateTimeOffset.TryParse(value, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AllowWhiteSpaces Or System.Globalization.DateTimeStyles.AssumeUniversal, parsed) Then
                Return parsed.ToLocalTime().ToString("dd.MM.yyyy · HH:mm zzz", System.Globalization.CultureInfo.CurrentCulture)
            End If
            Return "Time not recorded"
        End Function

        Private Shared Function ReportMetadata(data As ReviewReportData) As System.String
            Return "Created: " & GeneratedLabel(data.GeneratedUtc) & " · " & data.Rows.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " documents · " & data.Columns.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " fields"
        End Function

        Private Shared Function StatusSummary(data As ReviewReportData) As System.String
            Dim counts As New System.Collections.Generic.Dictionary(Of System.String, System.Int32)(System.StringComparer.Ordinal)
            Dim order As New System.Collections.Generic.List(Of System.String)()
            For Each item As ReviewReportItem In data.Items
                Dim label As System.String = ReviewStatus(item.Status)
                If Not counts.ContainsKey(label) Then
                    counts.Add(label, 0)
                    order.Add(label)
                End If
                counts(label) += 1
            Next
            Dim labels As New System.Collections.Generic.List(Of System.String)()
            For Each label As System.String In order
                labels.Add(label & ": " & counts(label).ToString(System.Globalization.CultureInfo.InvariantCulture))
            Next
            Return System.String.Join(" · ", labels)
        End Function

        Private Shared Function LinkedCell(value As System.String, sheet As System.String, row As System.Int32) As ReportCell
            Return New ReportCell(value) With {.Location = "'" & sheet.Replace("'", "''") & "'!A" & row.ToString(System.Globalization.CultureInfo.InvariantCulture)}
        End Function

        Private Shared Sub AddDetail(sheet As ReportSheet, reference As System.String, label As System.String, value As System.String,
                                     Optional includeEmpty As System.Boolean = False)
            If includeEmpty OrElse Not System.String.IsNullOrWhiteSpace(value) Then sheet.Rows.Add(Cells(reference, label, value))
        End Sub

        Private Shared Function ReportSections(data As ReviewReportData) As System.Collections.Generic.List(Of ReviewReportData)
            Dim sections As New System.Collections.Generic.List(Of ReviewReportData)()
            If data.TableLayout <> "types" Then
                sections.Add(data)
                Return sections
            End If
            Dim categories As New System.Collections.Generic.List(Of System.String)()
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            If data.Rows.Exists(Function(row As ReviewReportRow) row.IsGroup) Then categories.Add(Nothing)
            For Each row As ReviewReportRow In data.Rows
                If Not row.IsGroup AndAlso seen.Add(row.Category.Trim()) Then categories.Add(row.Category.Trim())
            Next
            For Each category As System.String In categories
                Dim groups As System.Boolean = category Is Nothing
                Dim heading As System.String = If(groups, "Group overview", If(category = "", "Unassigned type", category))
                Dim section As New ReviewReportData With {.Title = data.Title & " — " & heading, .Kind = data.Kind, .Orientation = data.Orientation, .Design = data.Design, .PaperSize = data.PaperSize,
                    .ProjectId = data.ProjectId, .ProjectRevision = data.ProjectRevision, .GeneratedUtc = data.GeneratedUtc}
                Dim indexes As New System.Collections.Generic.List(Of System.Int32)()
                For index As System.Int32 = 0 To data.Columns.Count - 1
                    Dim column As ReviewReportColumn = data.Columns(index)
                    If groups OrElse System.String.IsNullOrWhiteSpace(column.Category) OrElse System.String.Equals(column.Category.Trim(), category, System.StringComparison.OrdinalIgnoreCase) Then
                        section.Columns.Add(column)
                        indexes.Add(index)
                    End If
                Next
                If indexes.Count = 0 Then section.Columns.Add(New ReviewReportColumn With {.Key = "document", .Title = "Document"})
                For index As System.Int32 = 0 To data.Rows.Count - 1
                    Dim row As ReviewReportRow = data.Rows(index)
                    If row.IsGroup <> groups OrElse (Not groups AndAlso Not System.String.Equals(row.Category.Trim(), category, System.StringComparison.OrdinalIgnoreCase)) Then Continue For
                    Dim subset As New ReviewReportRow With {.DocumentId = row.DocumentId, .Category = row.Category, .IsGroup = row.IsGroup}
                    For Each field As System.Int32 In indexes
                        subset.Values.Add(row.Values(field))
                        subset.Styles.Add(row.Styles(field))
                    Next
                    If indexes.Count = 0 Then
                        subset.Values.Add(data.Documents(index).FileName)
                        subset.Styles.Add(0)
                    End If
                    section.Rows.Add(subset)
                    section.Documents.Add(data.Documents(index))
                Next
                sections.Add(section)
            Next
            Return sections
        End Function

        Private Shared Function BuildSheets(data As ReviewReportData, outputPath As System.String, cancellationToken As System.Threading.CancellationToken) As System.Collections.Generic.List(Of ReportSheet)
            If data.Kind = "matrix_xlsx" Then
                Dim headers As New System.Collections.Generic.List(Of System.String)()
                Dim widths As New System.Collections.Generic.List(Of System.Double)()
                For Each column As ReviewReportColumn In data.Columns
                    headers.Add(column.Title)
                    widths.Add(column.Width)
                Next
                Dim matrix As New ReportSheet With {.Name = "Results", .Title = data.Title, .PaperSize = data.PaperSize,
                    .Description = ReportMetadata(data),
                    .Headers = headers.ToArray(), .Widths = widths.ToArray(), .PreserveCells = True, .AutoSize = True}
                For Each row As ReviewReportRow In data.Rows
                    cancellationToken.ThrowIfCancellationRequested()
                    Dim values As System.Collections.Generic.List(Of ReportCell) = Cells(row.Values.ToArray())
                    For index As System.Int32 = 0 To row.Styles.Count - 1
                        values(index).Emphasis = row.Styles(index)
                    Next
                    matrix.Rows.Add(values)
                Next
                Return New System.Collections.Generic.List(Of ReportSheet) From {matrix}
            End If
            Throw New System.IO.InvalidDataException("Choose Matrix — Excel for a spreadsheet report.")
        End Function

        Private Shared Function WriteExcelReport(stream As System.IO.Stream, data As ReviewReportData, outputPath As System.String,
                                                 cancellationToken As System.Threading.CancellationToken) As System.String
            Using package As DocumentFormat.OpenXml.Packaging.SpreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook)
                MarkReport(package.AddCustomFilePropertiesPart())
                Dim workbook As DocumentFormat.OpenXml.Packaging.WorkbookPart = package.AddWorkbookPart()
                WriteExcelStyles(workbook.AddNewPart(Of DocumentFormat.OpenXml.Packaging.WorkbookStylesPart)(), data.Design)
                Dim sheets As New System.Collections.Generic.List(Of ReportSheet)()
                For Each section As ReviewReportData In ReportSections(data)
                    Dim sheet As ReportSheet = BuildSheets(section, outputPath, cancellationToken)(0)
                    If data.TableLayout = "types" Then
                        Dim label As System.String = If(section.Rows(0).IsGroup, "Groups", section.Rows(0).Category.Trim())
                        If label = "" Then label = "Unassigned"
                        For Each invalid As System.Char In "[]:*?/\"
                            label = label.Replace(invalid, "_"c)
                        Next
                        label = label.Trim("'"c)
                        Dim prefix As System.String = (sheets.Count + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) & " "
                        Dim length As System.Int32 = System.Math.Min(label.Length, 31 - prefix.Length)
                        If length > 0 AndAlso System.Char.IsHighSurrogate(label(length - 1)) Then length -= 1
                        sheet.Name = (prefix & label.Substring(0, length)).TrimEnd("'"c, " "c)
                    End If
                    sheets.Add(sheet)
                Next
                Dim relations As New System.Collections.Generic.List(Of System.String)()
                Dim longTexts As New System.Collections.Generic.List(Of LongText)()
                Using digest As New ContentDigest()
                    For Each sheet As ReportSheet In sheets
                        cancellationToken.ThrowIfCancellationRequested()
                        Dim part As DocumentFormat.OpenXml.Packaging.WorksheetPart = workbook.AddNewPart(Of DocumentFormat.OpenXml.Packaging.WorksheetPart)()
                        relations.Add(workbook.GetIdOfPart(part))
                        WriteExcelSheet(part, sheet, longTexts, digest, cancellationToken)
                    Next
                    WritePart(workbook,
                        Sub(writer As System.Xml.XmlWriter)
                            writer.WriteStartElement("workbook", SpreadsheetNamespace)
                            writer.WriteAttributeString("xmlns", "r", Nothing, RelationshipNamespace)
                            writer.WriteStartElement("bookViews", SpreadsheetNamespace)
                            writer.WriteStartElement("workbookView", SpreadsheetNamespace)
                            writer.WriteAttributeString("activeTab", If(data.Kind = "matrix_xlsx", "0", "1"))
                            writer.WriteEndElement()
                            writer.WriteEndElement()
                            writer.WriteStartElement("sheets", SpreadsheetNamespace)
                            For index As System.Int32 = 0 To sheets.Count - 1
                                writer.WriteStartElement("sheet", SpreadsheetNamespace)
                                writer.WriteAttributeString("name", sheets(index).Name)
                                writer.WriteAttributeString("sheetId", (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture))
                                writer.WriteAttributeString("r", "id", RelationshipNamespace, relations(index))
                                writer.WriteEndElement()
                            Next
                            writer.WriteEndElement()
                            writer.WriteStartElement("definedNames", SpreadsheetNamespace)
                            For index As System.Int32 = 0 To sheets.Count - 1
                                Dim sheetName As System.String = "'" & sheets(index).Name.Replace("'", "''") & "'!"
                                writer.WriteStartElement("definedName", SpreadsheetNamespace)
                                writer.WriteAttributeString("name", "_xlnm.Print_Titles")
                                writer.WriteAttributeString("localSheetId", index.ToString(System.Globalization.CultureInfo.InvariantCulture))
                                writer.WriteString(sheetName & "$1:$4")
                                writer.WriteEndElement()
                                writer.WriteStartElement("definedName", SpreadsheetNamespace)
                                writer.WriteAttributeString("name", "_xlnm.Print_Area")
                                writer.WriteAttributeString("localSheetId", index.ToString(System.Globalization.CultureInfo.InvariantCulture))
                                writer.WriteString(sheetName & "$A$1:$" & ColumnName(sheets(index).Headers.Length) & "$" & (sheets(index).Rows.Count + 4).ToString(System.Globalization.CultureInfo.InvariantCulture))
                                writer.WriteEndElement()
                            Next
                            writer.WriteEndElement()
                            writer.WriteEndElement()
                        End Sub)
                    Return digest.Finish()
                End Using
            End Using
        End Function


        Private Shared Sub WriteExcelSheet(part As DocumentFormat.OpenXml.Packaging.WorksheetPart, sheet As ReportSheet,
                                          longTexts As System.Collections.Generic.List(Of LongText), digest As ContentDigest,
                                          cancellationToken As System.Threading.CancellationToken)
            If sheet.Rows.Count > MaximumSheetRows - 4 Then Throw New System.IO.InvalidDataException("This report exceeds Excel's row limit. Export fewer items or choose a Word report; no text has been discarded.")
            If sheet.Headers.Length = 0 OrElse sheet.Headers.Length <> sheet.Widths.Length OrElse sheet.FreezeColumns < 1 OrElse sheet.FreezeColumns > sheet.Headers.Length Then
                Throw New System.IO.InvalidDataException("The report layout has inconsistent columns.")
            End If
            If sheet.AutoSize Then AutoSizeExcelColumns(sheet, cancellationToken)
            Dim lastColumn As System.String = ColumnName(sheet.Headers.Length)
            Dim lastRow As System.Int32 = sheet.Rows.Count + 4
            Dim mergedWidth As System.Double = 0.0R
            Dim headerHeight As System.Double = 32.0R
            For index As System.Int32 = 0 To sheet.Widths.Length - 1
                mergedWidth += sheet.Widths(index)
                headerHeight = System.Math.Max(headerHeight, System.Math.Min(96.0R, ExcelTextHeight(sheet.Headers(index), sheet.Widths(index))))
            Next
            ' The title uses a larger font across a merged row. Keep the complete
            ' merged width for headings, rather than treating them as column A.
            Dim titleWidth As System.Double = mergedWidth * 11.0R / 20.0R
            Dim titleHeight As System.Double = System.Math.Max(32.0R, System.Math.Min(120.0R, ExcelTextHeight(sheet.Title, titleWidth, 28.0R, System.Double.MaxValue)))
            Dim descriptionHeight As System.Double = System.Math.Max(46.0R, System.Math.Min(96.0R, ExcelTextHeight(sheet.Description, mergedWidth, 15.0R, System.Double.MaxValue)))
            If sheet.AutoSize Then
                titleHeight = MeasureExcelHeadingHeight(sheet.Title, sheet.PixelWidth, 20.0F, 30.0R, True)
                descriptionHeight = MeasureExcelHeadingHeight(sheet.Description, sheet.PixelWidth, 11.0F, 18.0R)
                headerHeight = 22.0R
                For index As System.Int32 = 0 To sheet.Headers.Length - 1
                    headerHeight = System.Math.Max(headerHeight, MeasureExcelHeadingHeight(sheet.Headers(index), sheet.ColumnPixelWidths(index), 11.0F, 22.0R, True))
                Next
            End If
            Dim links As New System.Collections.Generic.List(Of ReportHyperlink)()
            WritePart(part,
                Sub(writer As System.Xml.XmlWriter)
                    writer.WriteStartElement("worksheet", SpreadsheetNamespace)
                    writer.WriteAttributeString("xmlns", "r", Nothing, RelationshipNamespace)
                    writer.WriteRaw("<sheetPr><pageSetUpPr fitToPage='1'/></sheetPr>")
                    writer.WriteStartElement("dimension", SpreadsheetNamespace)
                    writer.WriteAttributeString("ref", "A1:" & lastColumn & lastRow.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    writer.WriteEndElement()
                    writer.WriteRaw("<sheetViews><sheetView showGridLines='0' workbookViewId='0'><selection activeCell='A1' sqref='A1'/></sheetView></sheetViews><sheetFormatPr defaultRowHeight='18'/>")
                    writer.WriteStartElement("cols", SpreadsheetNamespace)
                    For index As System.Int32 = 0 To sheet.Widths.Length - 1
                        writer.WriteStartElement("col", SpreadsheetNamespace)
                        writer.WriteAttributeString("min", (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture))
                        writer.WriteAttributeString("max", (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture))
                        writer.WriteAttributeString("width", sheet.Widths(index).ToString(System.Globalization.CultureInfo.InvariantCulture))
                        writer.WriteAttributeString("customWidth", "1")
                        writer.WriteEndElement()
                    Next
                    writer.WriteEndElement()
                    writer.WriteStartElement("sheetData", SpreadsheetNamespace)
                    StartExcelRow(writer, 1, titleHeight)
                    Dim titleLocation As System.String = WriteExcelCell(writer, sheet.Name, "A1", sheet.Title, 1, longTexts, digest, titleWidth, titleHeight,
                        lineHeight:=28.0R, maximumLineCharacters:=System.Double.MaxValue, preserveCell:=sheet.AutoSize)
                    If titleLocation.Length > 0 Then links.Add(New ReportHyperlink With {.Reference = "A1", .Location = titleLocation})
                    writer.WriteEndElement()
                    StartExcelRow(writer, 2, descriptionHeight)
                    Dim descriptionLocation As System.String = WriteExcelCell(writer, sheet.Name, "A2", sheet.Description, 2, longTexts, digest, mergedWidth, descriptionHeight,
                        maximumLineCharacters:=System.Double.MaxValue, preserveCell:=sheet.AutoSize)
                    If descriptionLocation.Length > 0 Then links.Add(New ReportHyperlink With {.Reference = "A2", .Location = descriptionLocation})
                    writer.WriteEndElement()
                    StartExcelRow(writer, 3, 5.0R)
                    writer.WriteEndElement()
                    StartExcelRow(writer, 4, headerHeight, sheet.AutoSize)
                    For index As System.Int32 = 0 To sheet.Headers.Length - 1
                        Dim headerReference As System.String = ColumnName(index + 1) & "4"
                        Dim headerLocation As System.String = WriteExcelCell(writer, sheet.Name, headerReference, sheet.Headers(index), 3, longTexts, digest, sheet.Widths(index), headerHeight)
                        If headerLocation.Length > 0 Then links.Add(New ReportHyperlink With {.Reference = headerReference, .Location = headerLocation})
                    Next
                    writer.WriteEndElement()
                    For rowIndex As System.Int32 = 0 To sheet.Rows.Count - 1
                        cancellationToken.ThrowIfCancellationRequested()
                        Dim row As System.Collections.Generic.List(Of ReportCell) = sheet.Rows(rowIndex)
                        If row.Count <> sheet.Headers.Length Then Throw New System.IO.InvalidDataException("A report row has inconsistent columns.")
                        Dim height As System.Double = 30.0R
                        For index As System.Int32 = 0 To row.Count - 1
                            height = System.Math.Max(height, System.Math.Min(sheet.MaximumRowHeight, ExcelRowHeight(row(index).Text, sheet.Widths(index))))
                        Next
                        StartExcelRow(writer, rowIndex + 5, height, sheet.AutoSize)
                        For index As System.Int32 = 0 To row.Count - 1
                            Dim reference As System.String = ColumnName(index + 1) & (rowIndex + 5).ToString(System.Globalization.CultureInfo.InvariantCulture)
                            Dim style As System.Int32 = If(rowIndex Mod 2 = 0, 4, 5)
                            Dim link As New ReportHyperlink With {.Reference = reference}
                            If row(index).Link IsNot Nothing AndAlso row(index).Location.Length > 0 Then Throw New System.IO.InvalidDataException("A report cell has two competing links.")
                            If row(index).Link IsNot Nothing OrElse row(index).Location.Length > 0 Then style = If(rowIndex Mod 2 = 0, 6, 7)
                            If row(index).Style >= 0 Then style = row(index).Style
                            Dim continuationLocation As System.String = WriteExcelCell(writer, sheet.Name, reference, row(index).Text, style, longTexts, digest,
                                sheet.Widths(index), height, row(index).Link, row(index).Location, preserveCell:=sheet.PreserveCells, emphasis:=row(index).Emphasis)
                            If continuationLocation.Length > 0 Then
                                link.Location = continuationLocation
                            ElseIf row(index).Link IsNot Nothing Then
                                Dim relation As DocumentFormat.OpenXml.Packaging.HyperlinkRelationship = part.AddHyperlinkRelationship(row(index).Link, True)
                                link.Relationship = relation.Id
                            Else
                                link.Location = row(index).Location
                            End If
                            If link.Relationship.Length > 0 OrElse link.Location.Length > 0 Then links.Add(link)
                        Next
                        writer.WriteEndElement()
                    Next
                    writer.WriteEndElement()
                    If sheet.Filter Then
                        writer.WriteStartElement("autoFilter", SpreadsheetNamespace)
                        writer.WriteAttributeString("ref", "A4:" & lastColumn & lastRow.ToString(System.Globalization.CultureInfo.InvariantCulture))
                        writer.WriteEndElement()
                    End If
                    If sheet.Headers.Length > 1 Then
                        writer.WriteStartElement("mergeCells", SpreadsheetNamespace)
                        writer.WriteAttributeString("count", "2")
                        For index As System.Int32 = 1 To 2
                            writer.WriteStartElement("mergeCell", SpreadsheetNamespace)
                            writer.WriteAttributeString("ref", "A" & index.ToString(System.Globalization.CultureInfo.InvariantCulture) & ":" & lastColumn & index.ToString(System.Globalization.CultureInfo.InvariantCulture))
                            writer.WriteEndElement()
                        Next
                        writer.WriteEndElement()
                    End If
                    If links.Count > 0 Then
                        writer.WriteStartElement("hyperlinks", SpreadsheetNamespace)
                        For Each link As ReportHyperlink In links
                            writer.WriteStartElement("hyperlink", SpreadsheetNamespace)
                            writer.WriteAttributeString("ref", link.Reference)
                            If link.Relationship.Length > 0 Then writer.WriteAttributeString("r", "id", RelationshipNamespace, link.Relationship)
                            If link.Location.Length > 0 Then writer.WriteAttributeString("location", link.Location)
                            writer.WriteEndElement()
                        Next
                        writer.WriteEndElement()
                    End If
                    writer.WriteRaw("<printOptions headings='0' gridLines='0'/>")
                    writer.WriteStartElement("pageMargins", SpreadsheetNamespace)
                    writer.WriteAttributeString("left", ExcelPrintSideMarginInches.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    writer.WriteAttributeString("right", ExcelPrintSideMarginInches.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    writer.WriteAttributeString("top", "0.5")
                    writer.WriteAttributeString("bottom", "0.5")
                    writer.WriteAttributeString("header", "0.2")
                    writer.WriteAttributeString("footer", "0.2")
                    writer.WriteEndElement()
                    writer.WriteStartElement("pageSetup", SpreadsheetNamespace)
                    writer.WriteAttributeString("paperSize", If(sheet.PaperSize = "A3", "8", "9"))
                    writer.WriteAttributeString("orientation", "landscape")
                    writer.WriteAttributeString("fitToWidth", "1")
                    writer.WriteAttributeString("fitToHeight", "0")
                    writer.WriteEndElement()
                    writer.WriteStartElement("headerFooter", SpreadsheetNamespace)
                    writer.WriteElementString("oddFooter", SpreadsheetNamespace, "&LTabular Review&C" & sheet.Name.Replace("&", "&&") & "&RPage &P of &N")
                    writer.WriteEndElement()
                    writer.WriteEndElement()
                End Sub)
        End Sub

        Private Shared Sub AutoSizeExcelColumns(sheet As ReportSheet, cancellationToken As System.Threading.CancellationToken)
            Using bitmap As New System.Drawing.Bitmap(1, 1)
                Using graphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(bitmap)
                    graphics.PageUnit = System.Drawing.GraphicsUnit.Pixel
                    Using font As New System.Drawing.Font("Calibri", 11.0F), headingFont As New System.Drawing.Font("Calibri", 11.0F, System.Drawing.FontStyle.Bold)
                        Using format As System.Drawing.StringFormat = DirectCast(System.Drawing.StringFormat.GenericTypographic.Clone(), System.Drawing.StringFormat)
                            format.FormatFlags = format.FormatFlags Or System.Drawing.StringFormatFlags.MeasureTrailingSpaces
                            Dim digit As System.Double = 1.0R
                            For number As System.Int32 = 0 To 9
                                digit = System.Math.Max(digit, System.Windows.Forms.TextRenderer.MeasureText(graphics,
                                    number.ToString(System.Globalization.CultureInfo.InvariantCulture), font,
                                    New System.Drawing.Size(System.Int32.MaxValue, System.Int32.MaxValue),
                                    System.Windows.Forms.TextFormatFlags.NoPadding Or System.Windows.Forms.TextFormatFlags.NoPrefix Or System.Windows.Forms.TextFormatFlags.SingleLine).Width)
                            Next
                            For column As System.Int32 = 0 To sheet.Headers.Length - 1
                                cancellationToken.ThrowIfCancellationRequested()
                                Dim pixels As System.Double = MeasureExcelTextWidth(graphics, headingFont, format, sheet.Headers(column), digit * 255)
                                For Each row As System.Collections.Generic.List(Of ReportCell) In sheet.Rows
                                    cancellationToken.ThrowIfCancellationRequested()
                                    pixels = System.Math.Max(pixels, MeasureExcelTextWidth(graphics, font, format, row(column).Text, digit * 255))
                                    If pixels >= digit * 255 Then Exit For
                                Next
                                sheet.Widths(column) = System.Math.Max(8.0R, System.Math.Min(255.0R, System.Math.Ceiling((pixels + 10.0R) / digit)))
                            Next
                            Dim pageInches As System.Double = If(sheet.PaperSize = "A3", 420.0R, 297.0R) / 25.4R
                            FitExcelColumnWidths(sheet.Widths, (pageInches - 2 * ExcelPrintSideMarginInches) * graphics.DpiX, digit)
                            sheet.PixelWidth = 0
                            sheet.ColumnPixelWidths = New System.Double(sheet.Widths.Length - 1) {}
                            For index As System.Int32 = 0 To sheet.Widths.Length - 1
                                sheet.ColumnPixelWidths(index) = ExcelColumnPixelWidth(sheet.Widths(index), digit)
                                sheet.PixelWidth += sheet.ColumnPixelWidths(index)
                            Next
                        End Using
                    End Using
                End Using
            End Using
        End Sub

        Private Shared Function ExcelColumnPixelWidth(width As System.Double, digitPixels As System.Double) As System.Double
            ' ISO/IEC 29500 column-width-to-pixel conversion, also documented
            ' by Microsoft Learn's Open XML Column class.
            Return System.Math.Floor(((256.0R * width + System.Math.Floor(128.0R / digitPixels)) / 256.0R) * digitPixels)
        End Function

        Private Shared Sub FitExcelColumnWidths(widths As System.Double(), availablePixels As System.Double, digitPixels As System.Double)
            ' Reserve one pixel per column for Open XML's pixel rounding.
            Dim availableUnits As System.Double = (availablePixels - widths.Length) / digitPixels
            If widths.Length = 0 OrElse availableUnits < widths.Length Then Throw New System.IO.InvalidDataException("The selected fields cannot fit this paper width. Choose A3 or fewer fields; no report file has changed.")
            Dim naturalUnits As System.Double = 0
            For Each width As System.Double In widths
                naturalUnits += width
            Next
            If naturalUnits <= availableUnits Then Return
            Dim minimum As System.Double = System.Math.Min(6.0R, availableUnits / widths.Length)
            Dim weight As System.Double = 0
            For Each width As System.Double In widths
                weight += System.Math.Max(0.0R, width - minimum)
            Next
            Dim flexible As System.Double = System.Math.Max(0.0R, availableUnits - minimum * widths.Length)
            For index As System.Int32 = 0 To widths.Length - 1
                Dim extra As System.Double = If(weight > 0, flexible * System.Math.Max(0.0R, widths(index) - minimum) / weight, 0.0R)
                ' Excel stores widths in 1/256-character units. Round down so
                ' rounding never pushes the last column beyond the print width.
                widths(index) = System.Math.Floor((minimum + extra) * 256.0R) / 256.0R
            Next
        End Sub

        Private Shared Function MeasureExcelHeadingHeight(value As System.String, widthPixels As System.Double,
                                                          fontSize As System.Single, minimum As System.Double, Optional bold As System.Boolean = False) As System.Double
            Using bitmap As New System.Drawing.Bitmap(1, 1)
                Using graphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(bitmap)
                    graphics.PageUnit = System.Drawing.GraphicsUnit.Pixel
                    Using font As New System.Drawing.Font("Calibri", fontSize, If(bold, System.Drawing.FontStyle.Bold, System.Drawing.FontStyle.Regular))
                        Dim layout As New System.Drawing.SizeF(CSng(System.Math.Max(1.0R, widthPixels - 12.0R)), 1000000.0F)
                        Dim size As System.Drawing.SizeF = graphics.MeasureString(If(value, ""), font, layout)
                        Return System.Math.Min(409.0R, System.Math.Max(minimum, System.Math.Ceiling(size.Height * 72.0R / graphics.DpiY + 6.0R)))
                    End Using
                End Using
            End Using
        End Function

        Private Shared Function MeasureExcelTextWidth(graphics As System.Drawing.Graphics, font As System.Drawing.Font,
                                                     format As System.Drawing.StringFormat, value As System.String, maximum As System.Double) As System.Double
            Dim result As System.Double = 0
            For Each line As System.String In If(value, "").Replace(Microsoft.VisualBasic.Constants.vbCrLf, Microsoft.VisualBasic.Constants.vbLf).Replace(Microsoft.VisualBasic.Constants.vbCr, Microsoft.VisualBasic.Constants.vbLf).Split(System.Convert.ToChar(10))
                Dim width As System.Double = 0
                For Each chunk As System.String In SplitText(line.Replace(Microsoft.VisualBasic.Constants.vbTab, "    "), ContinuationCharacters)
                    width += graphics.MeasureString(chunk, font, System.Drawing.PointF.Empty, format).Width
                    If width >= maximum Then Return maximum
                Next
                result = System.Math.Max(result, width)
            Next
            Return result
        End Function

        Private Shared Function ExcelTextHeight(value As System.String, width As System.Double,
                                                Optional lineHeight As System.Double = 15.0R,
                                                Optional maximumLineCharacters As System.Double = ExcelDisplayCharactersPerLine) As System.Double
            Dim lines As System.String() = If(value, "").Replace(Microsoft.VisualBasic.Constants.vbCrLf, Microsoft.VisualBasic.Constants.vbLf).Replace(Microsoft.VisualBasic.Constants.vbCr, Microsoft.VisualBasic.Constants.vbLf).Split(System.Convert.ToChar(10))
            Dim count As System.Double = 0
            Dim readableWidth As System.Double = System.Math.Min(maximumLineCharacters, System.Math.Max(4.0R, (width - 2.0R) / 2.0R))
            For Each line As System.String In lines
                Dim units As System.Int32 = line.Replace(Microsoft.VisualBasic.Constants.vbTab, "    ").Length
                count += System.Math.Max(1.0R, System.Math.Ceiling(units / readableWidth))
            Next
            Return System.Math.Max(30.0R, count * lineHeight + 8.0R)
        End Function

        Private Shared Function ExcelRowHeight(value As System.String, width As System.Double) As System.Double
            Return System.Math.Min(409.0R, ExcelTextHeight(value, width))
        End Function

        Private Shared Sub StartExcelRow(writer As System.Xml.XmlWriter, index As System.Int32, height As System.Double, Optional automatic As System.Boolean = False)
            writer.WriteStartElement("row", SpreadsheetNamespace)
            writer.WriteAttributeString("r", index.ToString(System.Globalization.CultureInfo.InvariantCulture))
            writer.WriteAttributeString("ht", height.ToString(System.Globalization.CultureInfo.InvariantCulture))
            writer.WriteAttributeString("customHeight", If(automatic, "0", "1"))
        End Sub


        Private Shared Function WriteExcelCell(writer As System.Xml.XmlWriter, sheet As System.String, reference As System.String,
                                               value As System.String, style As System.Int32,
                                               longTexts As System.Collections.Generic.List(Of LongText), digest As ContentDigest,
                                               width As System.Double, maximumHeight As System.Double,
                                               Optional originalLink As System.Uri = Nothing, Optional originalLocation As System.String = "",
                                               Optional lineHeight As System.Double = 15.0R,
                                               Optional maximumLineCharacters As System.Double = ExcelDisplayCharactersPerLine,
                                               Optional preserveCell As System.Boolean = False, Optional emphasis As System.Int32 = 0) As System.String
            Dim text As System.String = If(value, "")
            If text.Length > MaximumCellCharacters Then Throw New System.IO.InvalidDataException("A selected value exceeds Excel's cell limit. Choose Word to preserve its complete text; the Excel report structure and existing file remain unchanged.")
            Dim location As System.String = ""
            digest.Add(sheet, reference, text)
            writer.WriteStartElement("c", SpreadsheetNamespace)
            writer.WriteAttributeString("r", reference)
            writer.WriteAttributeString("s", style.ToString(System.Globalization.CultureInfo.InvariantCulture))
            writer.WriteAttributeString("t", "inlineStr")
            writer.WriteStartElement("is", SpreadsheetNamespace)
            If emphasis <> 0 Then
                writer.WriteStartElement("r", SpreadsheetNamespace)
                writer.WriteStartElement("rPr", SpreadsheetNamespace)
                If (emphasis And 1) <> 0 Then writer.WriteElementString("b", SpreadsheetNamespace, "")
                If (emphasis And 2) <> 0 Then writer.WriteElementString("i", SpreadsheetNamespace, "")
                writer.WriteEndElement()
            End If
            writer.WriteStartElement("t", SpreadsheetNamespace)
            writer.WriteAttributeString("xml", "space", XmlNamespace, "preserve")
            ' Excel interprets _xNNNN_ tokens even inside inline strings. Escape
            ' every token-start underscore, including overlapping literal tokens,
            ' before representing carriage returns explicitly.
            text = System.Text.RegularExpressions.Regex.Replace(text, "_(?=x[0-9A-Fa-f]{4}_)", "_x005F_")
            writer.WriteString(text.Replace(System.Convert.ToChar(13).ToString(), "_x000D_"))
            writer.WriteEndElement()
            writer.WriteEndElement()
            writer.WriteEndElement()
            If emphasis <> 0 Then writer.WriteEndElement()
            Return location
        End Function

        Private Shared Function SplitExcelText(value As System.String) As System.Collections.Generic.List(Of System.String)
            ' Bound presentation only. Concatenating these fragments restores the
            ' exact value, including every newline, tab and Unicode scalar.
            Dim parts As New System.Collections.Generic.List(Of System.String)()
            Dim start As System.Int32 = 0
            Dim position As System.Int32 = 0
            Dim lines As System.Int32 = 1
            Dim lineUnits As System.Int32 = 0
            While position < value.Length
                Dim current As System.Char = value(position)
                Dim tokenLength As System.Int32 = 1
                Dim lineBreak As System.Boolean = current = System.Convert.ToChar(10) OrElse current = System.Convert.ToChar(13)
                If current = System.Convert.ToChar(13) AndAlso position + 1 < value.Length AndAlso value(position + 1) = System.Convert.ToChar(10) Then
                    tokenLength = 2
                ElseIf System.Char.IsHighSurrogate(current) AndAlso position + 1 < value.Length AndAlso System.Char.IsLowSurrogate(value(position + 1)) Then
                    tokenLength = 2
                End If
                Dim nextLines As System.Int32 = lines
                Dim nextUnits As System.Int32 = lineUnits
                If lineBreak Then
                    nextLines += 1
                    nextUnits = 0
                Else
                    Dim units As System.Int32 = If(current = System.Convert.ToChar(9), 4, tokenLength)
                    If lineUnits > 0 AndAlso lineUnits + units > ExcelDisplayCharactersPerLine Then
                        nextLines += 1
                        nextUnits = units
                    Else
                        nextUnits += units
                    End If
                End If
                If position > start AndAlso (position - start + tokenLength > ExcelContinuationCharacters OrElse nextLines > ExcelContinuationLines) Then
                    parts.Add(value.Substring(start, position - start))
                    start = position
                    lines = 1
                    lineUnits = 0
                    Continue While
                End If
                lines = nextLines
                lineUnits = nextUnits
                position += tokenLength
            End While
            If start < value.Length Then parts.Add(value.Substring(start))
            If parts.Count = 0 Then parts.Add("")
            Return parts
        End Function

        Private Shared Function SplitText(value As System.String, maximum As System.Int32) As System.Collections.Generic.List(Of System.String)
            Dim parts As New System.Collections.Generic.List(Of System.String)()
            Dim start As System.Int32 = 0
            While start < value.Length
                Dim length As System.Int32 = System.Math.Min(maximum, value.Length - start)
                If start + length < value.Length AndAlso System.Char.IsHighSurrogate(value(start + length - 1)) Then length -= 1
                parts.Add(value.Substring(start, length))
                start += length
            End While
            If parts.Count = 0 Then parts.Add("")
            Return parts
        End Function

        Private Shared Function ColumnName(index As System.Int32) As System.String
            Dim result As System.String = ""
            While index > 0
                index -= 1
                result = System.Convert.ToChar(65 + index Mod 26).ToString() & result
                index \= 26
            End While
            Return result
        End Function

        Private Shared Sub WriteExcelStyles(part As DocumentFormat.OpenXml.Packaging.WorkbookStylesPart, design As System.String)
            WritePart(part,
                Sub(writer As System.Xml.XmlWriter)
                    writer.WriteStartElement("styleSheet", SpreadsheetNamespace)
                    writer.WriteRaw(
                        "<fonts count='8'><font><sz val='11'/><color rgb='FF332A2C'/><name val='Calibri'/></font><font><b/><sz val='20'/><color rgb='FF961E2D'/><name val='Calibri'/></font><font><sz val='11'/><color rgb='FF6F5F63'/><name val='Calibri'/></font><font><b/><sz val='11'/><color rgb='FFFFFFFF'/><name val='Calibri'/></font><font><u/><sz val='11'/><color rgb='FF961E2D'/><name val='Calibri'/></font><font><sz val='11'/><color rgb='FF704600'/><name val='Calibri'/></font><font><b/><sz val='11'/><color rgb='FF9F2135'/><name val='Calibri'/></font><font><sz val='11'/><color rgb='FF1B5868'/><name val='Calibri'/></font></fonts>" &
                        "<fills count='7'><fill><patternFill patternType='none'/></fill><fill><patternFill patternType='gray125'/></fill><fill><patternFill patternType='solid'><fgColor rgb='FF961E2D'/><bgColor indexed='64'/></patternFill></fill><fill><patternFill patternType='solid'><fgColor rgb='FFF7F0F1'/><bgColor indexed='64'/></patternFill></fill><fill><patternFill patternType='solid'><fgColor rgb='FFFFF4D6'/><bgColor indexed='64'/></patternFill></fill><fill><patternFill patternType='solid'><fgColor rgb='FFFCE8EC'/><bgColor indexed='64'/></patternFill></fill><fill><patternFill patternType='solid'><fgColor rgb='FFE5F2F7'/><bgColor indexed='64'/></patternFill></fill></fills>" &
                        "<borders count='2'><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style='thin'><color rgb='FFE4CDD1'/></bottom><diagonal/></border></borders>" &
                        "<cellStyleXfs count='1'><xf numFmtId='0' fontId='0' fillId='0' borderId='0'/></cellStyleXfs><cellXfs count='11'>" &
                        "<xf numFmtId='49' fontId='0' fillId='0' borderId='0' xfId='0'/>" &
                        "<xf numFmtId='49' fontId='1' fillId='0' borderId='0' xfId='0' applyAlignment='1'><alignment horizontal='left' vertical='top' wrapText='1'/></xf>" &
                        "<xf numFmtId='49' fontId='2' fillId='0' borderId='0' xfId='0' applyAlignment='1'><alignment horizontal='left' vertical='top' wrapText='1'/></xf>" &
                        "<xf numFmtId='49' fontId='3' fillId='2' borderId='1' xfId='0' applyFont='1' applyFill='1' applyBorder='1' applyAlignment='1'><alignment horizontal='left' vertical='top' wrapText='1'/></xf>" &
                        "<xf numFmtId='49' fontId='0' fillId='0' borderId='1' xfId='0' applyBorder='1' applyAlignment='1'><alignment horizontal='left' vertical='top' wrapText='1'/></xf>" &
                        "<xf numFmtId='49' fontId='0' fillId='3' borderId='1' xfId='0' applyFill='1' applyBorder='1' applyAlignment='1'><alignment horizontal='left' vertical='top' wrapText='1'/></xf>" &
                        "<xf numFmtId='49' fontId='4' fillId='0' borderId='1' xfId='0' applyFont='1' applyBorder='1' applyAlignment='1'><alignment horizontal='left' vertical='top' wrapText='1'/></xf>" &
                        "<xf numFmtId='49' fontId='4' fillId='3' borderId='1' xfId='0' applyFont='1' applyFill='1' applyBorder='1' applyAlignment='1'><alignment horizontal='left' vertical='top' wrapText='1'/></xf>" &
                        "<xf numFmtId='49' fontId='5' fillId='4' borderId='1' xfId='0' applyFont='1' applyFill='1' applyBorder='1' applyAlignment='1'><alignment horizontal='left' vertical='top' wrapText='1'/></xf>" &
                        "<xf numFmtId='49' fontId='6' fillId='5' borderId='1' xfId='0' applyFont='1' applyFill='1' applyBorder='1' applyAlignment='1'><alignment horizontal='left' vertical='top' wrapText='1'/></xf>" &
                        "<xf numFmtId='49' fontId='7' fillId='6' borderId='1' xfId='0' applyFont='1' applyFill='1' applyBorder='1' applyAlignment='1'><alignment horizontal='left' vertical='top' wrapText='1'/></xf>" &
                        "</cellXfs><cellStyles count='1'><cellStyle name='Normal' xfId='0' builtinId='0'/></cellStyles>")
                    writer.WriteEndElement()
                End Sub)
            ApplyReportDesign(part, design)
        End Sub


        Private Shared Function WriteWordReport(stream As System.IO.Stream, data As ReviewReportData, outputPath As System.String,
                                                cancellationToken As System.Threading.CancellationToken) As System.String
            Using package As DocumentFormat.OpenXml.Packaging.WordprocessingDocument = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document)
                MarkReport(package.AddCustomFilePropertiesPart())
                Dim main As DocumentFormat.OpenXml.Packaging.MainDocumentPart = package.AddMainDocumentPart()
                WriteWordStyles(main.AddNewPart(Of DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart)(), data.Design)
                Dim footer As DocumentFormat.OpenXml.Packaging.FooterPart = main.AddNewPart(Of DocumentFormat.OpenXml.Packaging.FooterPart)()
                WritePart(footer,
                    Sub(writer As System.Xml.XmlWriter)
                        writer.WriteStartElement("w", "ftr", WordNamespace)
                        writer.WriteRaw("<w:p><w:pPr><w:pStyle w:val='Metadata'/><w:jc w:val='right'/></w:pPr><w:r><w:t xml:space='preserve'>Tabular Review · Page </w:t></w:r><w:fldSimple w:instr='PAGE'><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space='preserve'> of </w:t></w:r><w:fldSimple w:instr='NUMPAGES'><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>")
                        writer.WriteEndElement()
                    End Sub)
                Dim documents As New System.Collections.Generic.Dictionary(Of System.String, ReviewReportDocument)(System.StringComparer.Ordinal)
                Dim labels As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
                Dim links As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
                For index As System.Int32 = 0 To data.Documents.Count - 1
                    cancellationToken.ThrowIfCancellationRequested()
                    Dim document As ReviewReportDocument = data.Documents(index)
                    documents.Add(document.Id, document)
                    labels.Add(document.Id, DocumentLabel(document, index))
                    Dim target As System.Uri = SourceLink(document.SourcePath, outputPath)
                    If target IsNot Nothing Then links.Add(document.Id, main.AddHyperlinkRelationship(target, True).Id)
                Next
                Using digest As New ContentDigest()
                    WritePart(main,
                        Sub(writer As System.Xml.XmlWriter)
                            writer.WriteStartElement("w", "document", WordNamespace)
                            writer.WriteAttributeString("xmlns", "r", Nothing, RelationshipNamespace)
                            writer.WriteStartElement("w", "body", WordNamespace)
                            For Each section As ReviewReportData In ReportSections(data)
                                WriteMatrixWordBody(writer, section, digest, cancellationToken)
                            Next
                            writer.WriteStartElement("w", "sectPr", WordNamespace)
                            writer.WriteStartElement("w", "footerReference", WordNamespace)
                            writer.WriteAttributeString("w", "type", WordNamespace, "default")
                            writer.WriteAttributeString("r", "id", RelationshipNamespace, main.GetIdOfPart(footer))
                            writer.WriteEndElement()
                            writer.WriteStartElement("w", "pgSz", WordNamespace)
                            writer.WriteAttributeString("w", "w", WordNamespace, WordPageWidth(data).ToString(System.Globalization.CultureInfo.InvariantCulture))
                            writer.WriteAttributeString("w", "h", WordNamespace, WordPageHeight(data).ToString(System.Globalization.CultureInfo.InvariantCulture))
                            If data.Orientation = "landscape" Then writer.WriteAttributeString("w", "orient", WordNamespace, "landscape")
                            writer.WriteEndElement()
                            writer.WriteRaw("<w:pgMar w:top='1134' w:right='1134' w:bottom='1134' w:left='1134' w:header='567' w:footer='567' w:gutter='0'/>")
                            writer.WriteEndElement()
                            writer.WriteEndElement()
                            writer.WriteEndElement()
                        End Sub)
                    ApplyReportDesign(main, data.Design)
                    Return digest.Finish()
                End Using
            End Using
        End Function

        Private Shared Function WordPageWidth(data As ReviewReportData) As System.Int32
            If data.PaperSize = "A3" Then Return If(data.Orientation = "landscape", 23811, 16838)
            Return If(data.Orientation = "landscape", 16838, 11906)
        End Function

        Private Shared Function WordPageHeight(data As ReviewReportData) As System.Int32
            If data.PaperSize = "A3" Then Return If(data.Orientation = "landscape", 16838, 23811)
            Return If(data.Orientation = "landscape", 11906, 16838)
        End Function

        Private Shared Sub WriteMatrixWordBody(writer As System.Xml.XmlWriter, data As ReviewReportData, digest As ContentDigest,
                                              cancellationToken As System.Threading.CancellationToken)
            WriteWordParagraph(writer, "Title", data.Title, digest)
            WriteWordParagraph(writer, "Subtitle", ReportMetadata(data), digest)
            If data.Kind = "document_report" Then
                WriteWordParagraph(writer, "Heading1", "Document review", digest)
                For index As System.Int32 = 0 To data.Rows.Count - 1
                    cancellationToken.ThrowIfCancellationRequested()
                    Dim row As ReviewReportRow = data.Rows(index)
                    Dim document As ReviewReportDocument = data.Documents.Find(Function(item As ReviewReportDocument) item.Id = row.DocumentId)
                    Dim heading As System.String = document.FileName
                    Dim nameColumn As System.Int32 = data.Columns.FindIndex(Function(column As ReviewReportColumn) column.Key = "document")
                    If Not row.IsGroup AndAlso nameColumn >= 0 Then heading = row.Values(nameColumn)
                    WriteWordParagraph(writer, "SectionHeading", heading, digest)
                    If Not row.IsGroup Then WriteWordParagraph(writer, "Metadata", SourceStatus(document) & " · " & SourceLocation(document), digest)
                    If Not row.IsGroup AndAlso document.ExtractionStatus <> "complete" Then WriteWordParagraph(writer, "Metadata", SourceNotice(document), digest)
                    For column As System.Int32 = 0 To data.Columns.Count - 1
                        WriteWordParagraph(writer, "Heading3", data.Columns(column).Title, digest)
                        WriteWordParagraph(writer, "Normal", row.Values(column), digest, emphasis:=row.Styles(column))
                    Next
                Next
                Return
            End If
            Dim pageWidth As System.Int32 = WordPageWidth(data) - 2 * 1134 - 40
            Dim total As System.Double = 0
            For Each column As ReviewReportColumn In data.Columns
                total += column.Width
            Next
            If data.Columns.Count > 63 Then Throw New System.IO.InvalidDataException("Word supports at most 63 table columns. Choose fewer fields, Excel or the document report.")
            Dim widths As New System.Collections.Generic.List(Of System.Int32)()
            Const minimumWidth As System.Int32 = 400
            If data.Columns.Count * minimumWidth > pageWidth Then Throw New System.IO.InvalidDataException("These fields cannot fit this Word page width. Choose A3 landscape, fewer fields or the document report.")
            Dim flexible As System.Int32 = pageWidth - data.Columns.Count * minimumWidth
            Dim remaining As System.Int32 = pageWidth
            For index As System.Int32 = 0 To data.Columns.Count - 1
                Dim width As System.Int32 = If(index = data.Columns.Count - 1, remaining, minimumWidth + CInt(System.Math.Floor(flexible * data.Columns(index).Width / total)))
                widths.Add(width)
                remaining -= width
            Next
            writer.WriteStartElement("w", "tbl", WordNamespace)
            writer.WriteRaw("<w:tblPr><w:tblW w:w='" & pageWidth.ToString(System.Globalization.CultureInfo.InvariantCulture) & "' w:type='dxa'/><w:jc w:val='left'/><w:tblInd w:w='0' w:type='dxa'/><w:tblBorders><w:top w:val='single' w:sz='4' w:color='E4CDD1'/><w:left w:val='single' w:sz='4' w:color='E4CDD1'/><w:bottom w:val='single' w:sz='4' w:color='E4CDD1'/><w:right w:val='single' w:sz='4' w:color='E4CDD1'/><w:insideH w:val='single' w:sz='4' w:color='E4CDD1'/><w:insideV w:val='single' w:sz='4' w:color='E4CDD1'/></w:tblBorders><w:tblLayout w:type='fixed'/><w:tblCellMar><w:top w:w='90' w:type='dxa'/><w:left w:w='100' w:type='dxa'/><w:bottom w:w='90' w:type='dxa'/><w:right w:w='100' w:type='dxa'/></w:tblCellMar></w:tblPr>")
            writer.WriteStartElement("w", "tblGrid", WordNamespace)
            For Each width As System.Int32 In widths
                writer.WriteRaw("<w:gridCol w:w='" & width.ToString(System.Globalization.CultureInfo.InvariantCulture) & "'/>")
            Next
            writer.WriteEndElement()
            For rowIndex As System.Int32 = -1 To data.Rows.Count - 1
                cancellationToken.ThrowIfCancellationRequested()
                writer.WriteStartElement("w", "tr", WordNamespace)
                If rowIndex < 0 Then writer.WriteRaw("<w:trPr><w:tblHeader/></w:trPr>")
                For column As System.Int32 = 0 To data.Columns.Count - 1
                    writer.WriteStartElement("w", "tc", WordNamespace)
                    writer.WriteRaw("<w:tcPr><w:tcW w:w='" & widths(column).ToString(System.Globalization.CultureInfo.InvariantCulture) & "' w:type='dxa'/>" &
                        If(rowIndex < 0, "<w:shd w:val='clear' w:fill='961E2D'/>", If(rowIndex Mod 2 = 1, "<w:shd w:val='clear' w:fill='F7F0F1'/>", "")) & "<w:vAlign w:val='top'/></w:tcPr>")
                    WriteWordParagraph(writer, If(rowIndex < 0, "RegisterHeader", "RegisterText"), If(rowIndex < 0, data.Columns(column).Title, data.Rows(rowIndex).Values(column)), digest, emphasis:=If(rowIndex < 0, 0, data.Rows(rowIndex).Styles(column)))
                    writer.WriteEndElement()
                Next
                writer.WriteEndElement()
            Next
            writer.WriteEndElement()
        End Sub

        Private Shared Sub WriteWordRegister(writer As System.Xml.XmlWriter, data As ReviewReportData, digest As ContentDigest,
                                            cancellationToken As System.Threading.CancellationToken)
            Dim widths As System.Int32() = {780, 900, 2780, 3378, 1800}
            Dim headings As System.String() = {"Item", "Priority", "Issue / title", "Next step", "Review"}
            writer.WriteStartElement("w", "tbl", WordNamespace)
            writer.WriteRaw("<w:tblPr><w:tblW w:w='9638' w:type='dxa'/><w:tblBorders><w:top w:val='single' w:sz='4' w:color='E4CDD1'/><w:left w:val='single' w:sz='4' w:color='E4CDD1'/><w:bottom w:val='single' w:sz='4' w:color='E4CDD1'/><w:right w:val='single' w:sz='4' w:color='E4CDD1'/><w:insideH w:val='single' w:sz='4' w:color='E4CDD1'/><w:insideV w:val='single' w:sz='4' w:color='E4CDD1'/></w:tblBorders><w:tblLayout w:type='fixed'/><w:tblCellMar><w:top w:w='90' w:type='dxa'/><w:left w:w='100' w:type='dxa'/><w:bottom w:w='90' w:type='dxa'/><w:right w:w='100' w:type='dxa'/></w:tblCellMar></w:tblPr>")
            writer.WriteStartElement("w", "tblGrid", WordNamespace)
            For Each width As System.Int32 In widths
                writer.WriteStartElement("w", "gridCol", WordNamespace)
                writer.WriteAttributeString("w", "w", WordNamespace, width.ToString(System.Globalization.CultureInfo.InvariantCulture))
                writer.WriteEndElement()
            Next
            writer.WriteEndElement()
            For rowIndex As System.Int32 = -1 To data.Items.Count - 1
                cancellationToken.ThrowIfCancellationRequested()
                Dim values As System.String() = headings
                If rowIndex >= 0 Then
                    Dim item As ReviewReportItem = data.Items(rowIndex)
                    values = New System.String() {ItemReference(rowIndex), PriorityLabel(item.Priority), ValueOr(item.Title, "Review item"), NextStep(item), ReviewLine(item)}
                End If
                writer.WriteStartElement("w", "tr", WordNamespace)
                If rowIndex < 0 Then writer.WriteRaw("<w:trPr><w:tblHeader/></w:trPr>")
                For column As System.Int32 = 0 To widths.Length - 1
                    writer.WriteStartElement("w", "tc", WordNamespace)
                    writer.WriteStartElement("w", "tcPr", WordNamespace)
                    writer.WriteStartElement("w", "tcW", WordNamespace)
                    writer.WriteAttributeString("w", "w", WordNamespace, widths(column).ToString(System.Globalization.CultureInfo.InvariantCulture))
                    writer.WriteAttributeString("w", "type", WordNamespace, "dxa")
                    writer.WriteEndElement()
                    If rowIndex < 0 Then
                        writer.WriteRaw("<w:shd w:val='clear' w:color='auto' w:fill='961E2D'/>")
                    ElseIf rowIndex Mod 2 = 1 Then
                        writer.WriteRaw("<w:shd w:val='clear' w:color='auto' w:fill='F7F0F1'/>")
                    End If
                    writer.WriteEndElement()
                    WriteWordParagraph(writer, If(rowIndex < 0, "RegisterHeader", "RegisterText"), values(column), digest)
                    writer.WriteEndElement()
                Next
                writer.WriteEndElement()
            Next
            writer.WriteEndElement()
        End Sub

        Private Shared Sub WriteWordFields(writer As System.Xml.XmlWriter, digest As ContentDigest, ParamArray fields As System.String())
            If fields.Length Mod 2 <> 0 Then Throw New System.ArgumentException("Report metadata needs complete label/value pairs.", NameOf(fields))
            writer.WriteStartElement("w", "tbl", WordNamespace)
            writer.WriteRaw("<w:tblPr><w:tblW w:w='9638' w:type='dxa'/><w:tblBorders><w:top w:val='single' w:sz='4' w:color='E4CDD1'/><w:left w:val='single' w:sz='4' w:color='E4CDD1'/><w:bottom w:val='single' w:sz='4' w:color='E4CDD1'/><w:right w:val='single' w:sz='4' w:color='E4CDD1'/><w:insideH w:val='single' w:sz='4' w:color='E4CDD1'/><w:insideV w:val='single' w:sz='4' w:color='E4CDD1'/></w:tblBorders><w:tblCellMar><w:top w:w='85' w:type='dxa'/><w:left w:w='110' w:type='dxa'/><w:bottom w:w='85' w:type='dxa'/><w:right w:w='110' w:type='dxa'/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w='2100'/><w:gridCol w:w='7538'/></w:tblGrid>")
            For index As System.Int32 = 0 To fields.Length - 1 Step 2
                writer.WriteStartElement("w", "tr", WordNamespace)
                For column As System.Int32 = 0 To 1
                    writer.WriteStartElement("w", "tc", WordNamespace)
                    writer.WriteStartElement("w", "tcPr", WordNamespace)
                    writer.WriteStartElement("w", "tcW", WordNamespace)
                    writer.WriteAttributeString("w", "w", WordNamespace, If(column = 0, "2100", "7538"))
                    writer.WriteAttributeString("w", "type", WordNamespace, "dxa")
                    writer.WriteEndElement()
                    If column = 0 Then writer.WriteRaw("<w:shd w:val='clear' w:color='auto' w:fill='F7F0F1'/>")
                    writer.WriteEndElement()
                    WriteWordParagraph(writer, If(column = 0, "FieldLabel", "Normal"), fields(index + column), digest)
                    writer.WriteEndElement()
                Next
                writer.WriteEndElement()
            Next
            writer.WriteEndElement()
        End Sub

        Private Shared Sub WriteWordParagraph(writer As System.Xml.XmlWriter, style As System.String, value As System.String,
                                              digest As ContentDigest, Optional linkId As System.String = "", Optional emphasis As System.Int32 = 0)
            digest.Add("p", "")
            writer.WriteStartElement("w", "p", WordNamespace)
            writer.WriteStartElement("w", "pPr", WordNamespace)
            writer.WriteStartElement("w", "pStyle", WordNamespace)
            writer.WriteAttributeString("w", "val", WordNamespace, style)
            writer.WriteEndElement()
            writer.WriteRaw("<w:jc w:val='left'/>")
            writer.WriteEndElement()
            If linkId.Length > 0 Then
                writer.WriteStartElement("w", "hyperlink", WordNamespace)
                writer.WriteAttributeString("r", "id", RelationshipNamespace, linkId)
            End If
            WriteWordText(writer, If(value, ""), digest, linkId.Length > 0, emphasis)
            If linkId.Length > 0 Then writer.WriteEndElement()
            writer.WriteEndElement()
        End Sub

        Private Shared Sub WriteWordText(writer As System.Xml.XmlWriter, value As System.String, digest As ContentDigest, hyperlink As System.Boolean, Optional emphasis As System.Int32 = 0)
            Dim start As System.Int32 = 0
            Dim index As System.Int32 = 0
            While index < value.Length
                Dim code As System.Int32 = System.Convert.ToInt32(value(index))
                If code = 9 OrElse code = 10 OrElse code = 13 Then
                    If index > start Then WriteWordRuns(writer, value.Substring(start, index - start), digest, hyperlink, emphasis)
                    writer.WriteStartElement("w", "r", WordNamespace)
                    Dim element As System.String = If(code = 9, "tab", "br")
                    writer.WriteStartElement("w", element, WordNamespace)
                    writer.WriteEndElement()
                    writer.WriteEndElement()
                    digest.Add(element, "")
                    If code = 13 AndAlso index + 1 < value.Length AndAlso System.Convert.ToInt32(value(index + 1)) = 10 Then index += 1
                    start = index + 1
                End If
                index += 1
            End While
            If start < value.Length OrElse value.Length = 0 Then WriteWordRuns(writer, value.Substring(start), digest, hyperlink, emphasis)
        End Sub

        Private Shared Sub WriteWordRuns(writer As System.Xml.XmlWriter, value As System.String, digest As ContentDigest, hyperlink As System.Boolean, Optional emphasis As System.Int32 = 0)
            For Each chunk As System.String In SplitText(value, ContinuationCharacters)
                writer.WriteStartElement("w", "r", WordNamespace)
                If hyperlink OrElse emphasis <> 0 Then
                    writer.WriteStartElement("w", "rPr", WordNamespace)
                    If hyperlink Then writer.WriteRaw("<w:rStyle w:val='Hyperlink'/>")
                    If (emphasis And 1) <> 0 Then writer.WriteRaw("<w:b/>")
                    If (emphasis And 2) <> 0 Then writer.WriteRaw("<w:i/>")
                    writer.WriteEndElement()
                End If
                writer.WriteStartElement("w", "t", WordNamespace)
                writer.WriteAttributeString("xml", "space", XmlNamespace, "preserve")
                writer.WriteString(chunk)
                writer.WriteEndElement()
                writer.WriteEndElement()
                digest.Add("t", chunk)
            Next
        End Sub

        Private Shared Sub WriteWordStyles(part As DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart, design As System.String)
            WritePart(part,
                Sub(writer As System.Xml.XmlWriter)
                    writer.WriteStartElement("w", "styles", WordNamespace)
                    writer.WriteRaw(
                        "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii='Calibri' w:hAnsi='Calibri' w:eastAsia='Calibri' w:cs='Calibri'/><w:color w:val='332A2C'/><w:sz w:val='22'/><w:szCs w:val='22'/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after='120' w:line='276' w:lineRule='auto'/></w:pPr></w:pPrDefault></w:docDefaults>" &
                        "<w:style w:type='paragraph' w:default='1' w:styleId='Normal'><w:name w:val='Normal'/><w:qFormat/></w:style>" &
                        "<w:style w:type='paragraph' w:styleId='Title'><w:name w:val='Title'/><w:basedOn w:val='Normal'/><w:next w:val='Subtitle'/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:after='120'/></w:pPr><w:rPr><w:b/><w:color w:val='961E2D'/><w:sz w:val='44'/></w:rPr></w:style>" &
                        "<w:style w:type='paragraph' w:styleId='Subtitle'><w:name w:val='Subtitle'/><w:basedOn w:val='Normal'/><w:next w:val='Normal'/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:after='280'/></w:pPr><w:rPr><w:color w:val='6F5F63'/><w:sz w:val='22'/></w:rPr></w:style>" &
                        "<w:style w:type='paragraph' w:styleId='Heading1'><w:name w:val='heading 1'/><w:basedOn w:val='Normal'/><w:next w:val='Normal'/><w:uiPriority w:val='9'/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before='360' w:after='160'/><w:outlineLvl w:val='0'/></w:pPr><w:rPr><w:b/><w:color w:val='961E2D'/><w:sz w:val='32'/></w:rPr></w:style>" &
                        "<w:style w:type='paragraph' w:styleId='Heading2'><w:name w:val='heading 2'/><w:basedOn w:val='Normal'/><w:next w:val='Normal'/><w:uiPriority w:val='9'/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before='280' w:after='120'/><w:outlineLvl w:val='1'/></w:pPr><w:rPr><w:b/><w:color w:val='961E2D'/><w:sz w:val='26'/></w:rPr></w:style>" &
                        "<w:style w:type='paragraph' w:styleId='Heading3'><w:name w:val='heading 3'/><w:basedOn w:val='Normal'/><w:next w:val='Normal'/><w:uiPriority w:val='9'/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before='180' w:after='90'/><w:outlineLvl w:val='2'/></w:pPr><w:rPr><w:b/><w:color w:val='961E2D'/><w:sz w:val='22'/></w:rPr></w:style>" &
                        "<w:style w:type='paragraph' w:styleId='SectionHeading'><w:name w:val='Report section'/><w:basedOn w:val='Heading1'/><w:next w:val='Normal'/><w:qFormat/><w:pPr><w:pageBreakBefore/></w:pPr></w:style>" &
                        "<w:style w:type='paragraph' w:styleId='RegisterHeader'><w:name w:val='Register heading'/><w:basedOn w:val='Normal'/><w:pPr><w:wordWrap w:val='0'/><w:spacing w:after='30' w:line='240' w:lineRule='auto'/></w:pPr><w:rPr><w:b/><w:color w:val='FFFFFF'/><w:sz w:val='19'/></w:rPr></w:style>" &
                        "<w:style w:type='paragraph' w:styleId='RegisterText'><w:name w:val='Register text'/><w:basedOn w:val='Normal'/><w:pPr><w:wordWrap w:val='0'/><w:spacing w:after='40' w:line='240' w:lineRule='auto'/></w:pPr><w:rPr><w:sz w:val='19'/></w:rPr></w:style>" &
                        "<w:style w:type='paragraph' w:styleId='Quote'><w:name w:val='Source quotation'/><w:basedOn w:val='Normal'/><w:pPr><w:pBdr><w:left w:val='single' w:sz='12' w:space='8' w:color='961E2D'/></w:pBdr><w:shd w:val='clear' w:color='auto' w:fill='F7F0F1'/><w:spacing w:after='100'/><w:ind w:left='240' w:right='240'/></w:pPr></w:style>" &
                        "<w:style w:type='paragraph' w:styleId='Citation'><w:name w:val='Source reference'/><w:basedOn w:val='Normal'/><w:pPr><w:keepNext/><w:spacing w:before='120' w:after='75'/></w:pPr><w:rPr><w:b/><w:sz w:val='20'/></w:rPr></w:style>" &
                        "<w:style w:type='paragraph' w:styleId='Metadata'><w:name w:val='Review metadata'/><w:basedOn w:val='Normal'/><w:pPr><w:spacing w:after='90'/></w:pPr><w:rPr><w:color w:val='6F5F63'/><w:sz w:val='18'/></w:rPr></w:style>" &
                        "<w:style w:type='paragraph' w:styleId='FieldLabel'><w:name w:val='Field label'/><w:basedOn w:val='Normal'/><w:rPr><w:b/><w:sz w:val='20'/></w:rPr></w:style>" &
                        "<w:style w:type='character' w:styleId='Hyperlink'><w:name w:val='Hyperlink'/><w:rPr><w:color w:val='961E2D'/><w:u w:val='single'/></w:rPr></w:style>")
                    writer.WriteEndElement()
                End Sub)
            ApplyReportDesign(part, design)
        End Sub

        ' Only presentation attributes participate in a design. Never replace text,
        ' field keys, hyperlinks, source identities or user-authored color-like strings.
        Private Shared Sub ApplyReportDesign(part As DocumentFormat.OpenXml.Packaging.OpenXmlPart, design As System.String)
            If design = "redink" Then Return
            If design <> "neutral" Then Throw New System.IO.InvalidDataException("Choose a supported report design.")
            Dim colors As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal) From {
                {"961E2D", "333333"}, {"F7F0F1", "F2F2F2"}, {"E4CDD1", "D6D6D6"}, {"332A2C", "222222"}, {"6F5F63", "666666"}}
            Dim document As New System.Xml.XmlDocument With {.PreserveWhitespace = True, .XmlResolver = Nothing}
            Using reader As System.Xml.XmlReader = ReadPart(part)
                document.Load(reader)
            End Using
            Dim namespaces As New System.Xml.XmlNamespaceManager(document.NameTable)
            namespaces.AddNamespace("w", WordNamespace)
            namespaces.AddNamespace("s", SpreadsheetNamespace)
            For Each attribute As System.Xml.XmlNode In document.SelectNodes("//w:color/@w:val | //@w:color | //w:shd/@w:fill | //s:color/@rgb | //s:fgColor/@rgb | //s:bgColor/@rgb", namespaces)
                Dim value As System.String = attribute.Value
                Dim prefix As System.String = If(value.Length = 8, value.Substring(0, 2), "")
                Dim replacement As System.String = Nothing
                If colors.TryGetValue(If(prefix.Length > 0, value.Substring(2), value), replacement) Then attribute.Value = prefix & replacement
            Next
            Using stream As System.IO.Stream = part.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write)
                document.Save(stream)
            End Using
        End Sub

        Private Shared Sub ValidatePackage(path As System.String, kind As System.String, expectedDigest As System.String,
                                           cancellationToken As System.Threading.CancellationToken)
            Dim actualDigest As System.String
            If IsWordReport(kind) Then
                Using package As DocumentFormat.OpenXml.Packaging.WordprocessingDocument = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, False)
                    ValidateOpenXml(package, cancellationToken)
                    Using digest As New ContentDigest()
                        Using reader As System.Xml.XmlReader = ReadPart(package.MainDocumentPart)
                            While reader.Read()
                                cancellationToken.ThrowIfCancellationRequested()
                                If reader.NodeType <> System.Xml.XmlNodeType.Element OrElse reader.NamespaceURI <> WordNamespace Then Continue While
                                Select Case reader.LocalName
                                    Case "p", "br", "tab"
                                        digest.Add(reader.LocalName, "")
                                    Case "t"
                                        digest.Add("t", reader.ReadElementContentAsString())
                                End Select
                            End While
                        End Using
                        actualDigest = digest.Finish()
                    End Using
                End Using
            Else
                Using package As DocumentFormat.OpenXml.Packaging.SpreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, False)
                    ValidateOpenXml(package, cancellationToken)
                    Dim sheets As New System.Collections.Generic.List(Of System.Collections.Generic.KeyValuePair(Of System.String, System.String))()
                    Using reader As System.Xml.XmlReader = ReadPart(package.WorkbookPart)
                        While reader.Read()
                            If reader.NodeType = System.Xml.XmlNodeType.Element AndAlso reader.NamespaceURI = SpreadsheetNamespace AndAlso reader.LocalName = "sheet" Then
                                sheets.Add(New System.Collections.Generic.KeyValuePair(Of System.String, System.String)(reader.GetAttribute("name"), reader.GetAttribute("id", RelationshipNamespace)))
                            End If
                        End While
                    End Using
                    Using digest As New ContentDigest()
                        For Each sheet As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In sheets
                            Dim part As DocumentFormat.OpenXml.Packaging.WorksheetPart = TryCast(package.WorkbookPart.GetPartById(sheet.Value), DocumentFormat.OpenXml.Packaging.WorksheetPart)
                            If part Is Nothing Then Throw New System.IO.InvalidDataException("The generated workbook has an invalid worksheet relationship.")
                            Using reader As System.Xml.XmlReader = ReadPart(part)
                                While reader.Read()
                                    cancellationToken.ThrowIfCancellationRequested()
                                    If reader.NodeType <> System.Xml.XmlNodeType.Element OrElse reader.NamespaceURI <> SpreadsheetNamespace OrElse reader.LocalName <> "c" Then Continue While
                                    Dim reference As System.String = reader.GetAttribute("r")
                                    If reader.GetAttribute("t") <> "inlineStr" Then Throw New System.IO.InvalidDataException("An exported value is not an explicit text cell.")
                                    Dim text As New System.Text.StringBuilder()
                                    Dim textElements As System.Int32 = 0
                                    Using cell As System.Xml.XmlReader = reader.ReadSubtree()
                                        While cell.Read()
                                            If cell.NodeType <> System.Xml.XmlNodeType.Element OrElse cell.NamespaceURI <> SpreadsheetNamespace Then Continue While
                                            If cell.LocalName = "f" Then Throw New System.IO.InvalidDataException("An exported text cell unexpectedly contains a formula.")
                                            If cell.LocalName = "t" Then
                                                text.Append(cell.ReadElementContentAsString())
                                                textElements += 1
                                            End If
                                        End While
                                    End Using
                                    If textElements <> 1 Then Throw New System.IO.InvalidDataException("An exported text cell has an unexpected content structure.")
                                    Dim decoded As System.String = System.Text.RegularExpressions.Regex.Replace(text.ToString(), "_x([0-9A-Fa-f]{4})_",
                                        Function(match As System.Text.RegularExpressions.Match) System.Convert.ToChar(System.Int32.Parse(match.Groups(1).Value, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture)).ToString())
                                    If decoded.Length > MaximumCellCharacters Then Throw New System.IO.InvalidDataException("An exported cell exceeds Excel's text limit.")
                                    digest.Add(sheet.Key, reference, decoded)
                                End While
                            End Using
                        Next
                        actualDigest = digest.Finish()
                    End Using
                End Using
            End If
            If Not System.String.Equals(expectedDigest, actualDigest, System.StringComparison.Ordinal) Then Throw New System.IO.InvalidDataException("The report failed its complete-text readback check. No existing output file has changed.")
        End Sub

        Private Shared Function ReadPart(part As DocumentFormat.OpenXml.Packaging.OpenXmlPart) As System.Xml.XmlReader
            If part Is Nothing Then Throw New System.IO.InvalidDataException("The generated report is missing a required document part.")
            Return System.Xml.XmlReader.Create(part.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read), New System.Xml.XmlReaderSettings With {
                .DtdProcessing = System.Xml.DtdProcessing.Prohibit, .XmlResolver = Nothing, .CloseInput = True, .IgnoreComments = True})
        End Function

        Private Shared Sub ValidateOpenXml(package As DocumentFormat.OpenXml.Packaging.OpenXmlPackage, cancellationToken As System.Threading.CancellationToken)
            Dim validator As New DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2013) With {.MaxNumberOfErrors = 20}
            Dim invalid As System.Boolean = False
            For Each failure As DocumentFormat.OpenXml.Validation.ValidationErrorInfo In validator.Validate(package, cancellationToken)
                invalid = True
                System.Diagnostics.Trace.WriteLine("[TabularReview:report_validation_failed] " & failure.Id & " | " & failure.Description)
            Next
            If invalid Then Throw New System.IO.InvalidDataException("Office package validation rejected the generated report. No existing output file has changed. Retry the export after correcting the reported export error.")
        End Sub
    End Class
End Namespace
