' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. For license to use see https://redink.ai.
'
' Source adapter for Tabular Review. Reuses the Data Collector's shared readers;
' the additional layer only pins source bytes, preserves provenance and creates
' deterministic review anchors. Source files are never changed.
Option Explicit On
Option Strict On
Option Infer On

Namespace Global.SharedLibrary.TabularReview
    Public NotInheritable Class ReviewSourceService
        Public Const ReportArtifactProperty As System.String = "RedInk.TabularReview.Artifact"
        Public Const ReportArtifactValue As System.String = "report.v1"
        Private Const WorkbookStoreMarker As System.String = "RedInk.TabularReview.Store.v1"
        Private Const SpreadsheetNamespace As System.String = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        Private Const CustomPropertiesNamespace As System.String = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
        Private Const VariantNamespace As System.String = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

        Private Sub New()
        End Sub

        Private Shared ReadOnly Extensions As New System.Collections.Generic.HashSet(Of System.String)(
            New System.String() {".pdf", ".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm",
                ".txt", ".rtf", ".ini", ".csv", ".tsv", ".log", ".json", ".xml", ".html", ".htm",
                ".md", ".yaml", ".yml", ".eml", ".msg"}, System.StringComparer.OrdinalIgnoreCase)

        Public Shared Function IsSupportedFile(path As System.String,
                                               Optional allowLegacyDoc As System.Boolean = False) As System.Boolean
            Dim extension As System.String = System.IO.Path.GetExtension(path)
            Return Extensions.Contains(extension) OrElse
                (allowLegacyDoc AndAlso System.String.Equals(extension, ".doc", System.StringComparison.OrdinalIgnoreCase))
        End Function

        ''' <summary>Returns a stable file order. Access failures abort the enumeration:
        ''' a missing subtree must not silently look like a complete document population.</summary>
        Public Shared Function EnumerateFiles(root As System.String, recursive As System.Boolean,
                                             Optional allowLegacyDoc As System.Boolean = False,
                                             Optional excludedPaths As System.Collections.Generic.IEnumerable(Of System.String) = Nothing,
                                             Optional cancellationToken As System.Threading.CancellationToken = Nothing) As System.Collections.Generic.List(Of System.String)
            Dim normalizedRoot As System.String = System.IO.Path.GetFullPath(root)
            If Not System.IO.Directory.Exists(normalizedRoot) Then Throw New System.IO.DirectoryNotFoundException("The source folder is unavailable: " & normalizedRoot)
            Dim excluded As New System.Collections.Generic.List(Of System.String)()
            If excludedPaths IsNot Nothing Then
                For Each path As System.String In excludedPaths
                    If Not System.String.IsNullOrWhiteSpace(path) AndAlso System.IO.Path.IsPathRooted(path) AndAlso path.IndexOf("://", System.StringComparison.Ordinal) < 0 Then excluded.Add(System.IO.Path.GetFullPath(path))
                Next
            End If
            Dim result As New System.Collections.Generic.List(Of System.String)()
            Dim pending As New System.Collections.Generic.Stack(Of System.String)()
            pending.Push(normalizedRoot)
            While pending.Count > 0
                cancellationToken.ThrowIfCancellationRequested()
                Dim directory As System.String = pending.Pop()
                Try
                    For Each file As System.String In System.IO.Directory.GetFiles(directory)
                        If System.IO.Path.GetFileName(file).StartsWith("~$", System.StringComparison.Ordinal) Then Continue For
                        If (System.IO.File.GetAttributes(file) And System.IO.FileAttributes.ReparsePoint) <> 0 Then Continue For
                        If Not IsSupportedFile(file, allowLegacyDoc) Then Continue For
                        cancellationToken.ThrowIfCancellationRequested()
                        Dim generated As System.Boolean = False
                        For Each protectedPath As System.String In excluded
                            If System.String.Equals(System.IO.Path.GetFullPath(file), protectedPath, System.StringComparison.OrdinalIgnoreCase) OrElse
                               (System.IO.File.Exists(protectedPath) AndAlso Global.SharedLibrary.SharedLibrary.GeneratedOutputRegistry.IsPhysicalPathAtOrBelow(protectedPath, file)) Then
                                generated = True
                                Exit For
                            End If
                        Next
                        If generated OrElse ReviewArtifactReason(file, cancellationToken).Length > 0 Then
                            System.Diagnostics.Trace.WriteLine("[TabularReview:source_excluded] " & System.IO.Path.GetFileName(file))
                            Continue For
                        End If
                        result.Add(file)
                    Next
                    If recursive Then
                        For Each child As System.String In System.IO.Directory.GetDirectories(directory)
                            If (System.IO.File.GetAttributes(child) And System.IO.FileAttributes.ReparsePoint) = 0 Then pending.Push(child)
                        Next
                    End If
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    Throw New System.IO.IOException("Cannot enumerate source folder '" & directory & "'. The folder import was stopped so that documents are not silently omitted. " & ex.Message, ex)
                End Try
            End While
            result.Sort(System.StringComparer.OrdinalIgnoreCase)
            Return result
        End Function

        ' Exact internal tags and the persisted store marker identify artifacts;
        ' report titles, filenames, colors and business text never decide eligibility.
        Public Shared Function ReviewArtifactReason(path As System.String, cancellationToken As System.Threading.CancellationToken) As System.String
            Dim extension As System.String = System.IO.Path.GetExtension(path).ToLowerInvariant()
            If extension <> ".xlsx" AndAlso extension <> ".xlsm" AndAlso extension <> ".docx" AndAlso extension <> ".docm" Then Return ""
            cancellationToken.ThrowIfCancellationRequested()
            Try
                Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite Or System.IO.FileShare.Delete)
                    Using archive As New System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Read, False)
                        Dim properties As System.IO.Compression.ZipArchiveEntry = archive.GetEntry("docProps/custom.xml")
                        If properties IsNot Nothing Then
                            Using reader As System.Xml.XmlReader = OpenArtifactXml(properties)
                                While reader.Read()
                                    cancellationToken.ThrowIfCancellationRequested()
                                    If reader.NodeType <> System.Xml.XmlNodeType.Element OrElse reader.NamespaceURI <> CustomPropertiesNamespace OrElse reader.LocalName <> "property" OrElse reader.GetAttribute("name") <> ReportArtifactProperty Then Continue While
                                    Using value As System.Xml.XmlReader = reader.ReadSubtree()
                                        While value.Read()
                                            cancellationToken.ThrowIfCancellationRequested()
                                            If value.NodeType = System.Xml.XmlNodeType.Element AndAlso value.NamespaceURI = VariantNamespace AndAlso value.LocalName = "lpwstr" AndAlso value.ReadString() = ReportArtifactValue Then Return "report"
                                        End While
                                    End Using
                                End While
                            End Using
                        End If
                        If extension = ".xlsx" OrElse extension = ".xlsm" Then
                            Dim indices As New System.Collections.Generic.HashSet(Of System.Int32)()
                            For Each entry As System.IO.Compression.ZipArchiveEntry In archive.Entries
                                cancellationToken.ThrowIfCancellationRequested()
                                If Not entry.FullName.StartsWith("xl/worksheets/", System.StringComparison.Ordinal) OrElse Not entry.FullName.EndsWith(".xml", System.StringComparison.Ordinal) Then Continue For
                                Using reader As System.Xml.XmlReader = OpenArtifactXml(entry)
                                    While reader.Read()
                                        cancellationToken.ThrowIfCancellationRequested()
                                        If reader.NodeType <> System.Xml.XmlNodeType.Element OrElse reader.NamespaceURI <> SpreadsheetNamespace Then Continue While
                                        If reader.LocalName = "row" Then
                                            Dim row As System.Int32
                                            If System.Int32.TryParse(reader.GetAttribute("r"), row) AndAlso row > 1 Then Exit While
                                        End If
                                        If reader.LocalName <> "c" OrElse reader.GetAttribute("r") <> "A1" Then Continue While
                                        Dim cellType As System.String = reader.GetAttribute("t")
                                        Dim text As New System.Text.StringBuilder()
                                        Dim hasFormula As System.Boolean = False
                                        Dim sharedIndex As System.Int32 = -1
                                        Using cell As System.Xml.XmlReader = reader.ReadSubtree()
                                            While cell.Read()
                                                cancellationToken.ThrowIfCancellationRequested()
                                                If cell.NodeType <> System.Xml.XmlNodeType.Element OrElse cell.NamespaceURI <> SpreadsheetNamespace Then Continue While
                                                If cell.LocalName = "f" Then hasFormula = True
                                                If cellType = "s" AndAlso cell.LocalName = "v" Then
                                                    Dim index As System.Int32
                                                    If System.Int32.TryParse(cell.ReadString(), index) AndAlso index >= 0 Then sharedIndex = index
                                                ElseIf (cellType = "inlineStr" AndAlso cell.LocalName = "t") OrElse (cellType = "str" AndAlso cell.LocalName = "v") Then
                                                    text.Append(cell.ReadString())
                                                End If
                                            End While
                                        End Using
                                        If Not hasFormula Then
                                            If sharedIndex >= 0 Then indices.Add(sharedIndex)
                                            If text.ToString() = WorkbookStoreMarker Then Return "review workbook"
                                        End If
                                        Exit While
                                    End While
                                End Using
                            Next
                            Dim stringsEntry As System.IO.Compression.ZipArchiveEntry = archive.GetEntry("xl/sharedStrings.xml")
                            If indices.Count > 0 AndAlso stringsEntry IsNot Nothing Then
                                Dim maximumIndex As System.Int32 = 0
                                For Each candidateIndex As System.Int32 In indices
                                    maximumIndex = System.Math.Max(maximumIndex, candidateIndex)
                                Next
                                Using reader As System.Xml.XmlReader = OpenArtifactXml(stringsEntry)
                                    Dim index As System.Int32 = -1
                                    While reader.Read()
                                        cancellationToken.ThrowIfCancellationRequested()
                                        If reader.NodeType <> System.Xml.XmlNodeType.Element OrElse reader.NamespaceURI <> SpreadsheetNamespace OrElse reader.LocalName <> "si" Then Continue While
                                        index += 1
                                        If index > maximumIndex Then Exit While
                                        If Not indices.Contains(index) Then Continue While
                                        Dim text As New System.Text.StringBuilder()
                                        Using value As System.Xml.XmlReader = reader.ReadSubtree()
                                            While value.Read()
                                                cancellationToken.ThrowIfCancellationRequested()
                                                If value.NodeType = System.Xml.XmlNodeType.Element AndAlso value.NamespaceURI = SpreadsheetNamespace AndAlso value.LocalName = "t" Then text.Append(value.ReadString())
                                            End While
                                        End Using
                                        If text.ToString() = WorkbookStoreMarker Then Return "review workbook"
                                    End While
                                End Using
                            End If
                        End If
                    End Using
                End Using
            Catch failure As System.OperationCanceledException
                Throw
            Catch failure As System.IO.InvalidDataException
                ' A non-package/corrupt source remains an explicit extraction error;
                ' it does not become an invented positive artifact match.
                System.Diagnostics.Trace.WriteLine("[TabularReview:artifact_inspection_invalid_package] " & failure.Message)
            End Try
            Return ""
        End Function

        Private Shared Function OpenArtifactXml(entry As System.IO.Compression.ZipArchiveEntry) As System.Xml.XmlReader
            Return System.Xml.XmlReader.Create(entry.Open(), New System.Xml.XmlReaderSettings With {
                .DtdProcessing = System.Xml.DtdProcessing.Prohibit, .XmlResolver = Nothing, .CloseInput = True})
        End Function

        Public Shared Function FileHash(path As System.String, Optional cancellationToken As System.Threading.CancellationToken = Nothing) As System.String
            cancellationToken.ThrowIfCancellationRequested()
            Using input As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                    Dim buffer(81919) As System.Byte
                    Do
                        cancellationToken.ThrowIfCancellationRequested()
                        Dim count As System.Int32 = input.Read(buffer, 0, buffer.Length)
                        If count = 0 Then Exit Do
                        algorithm.TransformBlock(buffer, 0, count, buffer, 0)
                    Loop
                    cancellationToken.ThrowIfCancellationRequested()
                    algorithm.TransformFinalBlock(New System.Byte() {}, 0, 0)
                    Return ToHashString(algorithm.Hash)
                End Using
            End Using
        End Function

        Public Shared Async Function ExtractAsync(path As System.String, root As System.String,
                context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext,
                doOcr As System.Boolean, cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of ReviewDocument)
            Dim fullPath As System.String = System.IO.Path.GetFullPath(path)
            Dim relative As System.String = RelativeSourcePath(fullPath, root)
            Dim document As New ReviewDocument With {
                .SourcePath = fullPath,
                .RelativePath = relative,
                .FileName = System.IO.Path.GetFileName(fullPath),
                .Number = DeriveDocumentNumber(relative)
            }
            Dim staging As System.String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "RedInkReviewSource-" & System.Guid.NewGuid().ToString("N"))
            Try
                cancellationToken.ThrowIfCancellationRequested()
                If Not System.IO.File.Exists(fullPath) Then
                    document.ExtractionStatus = "missing"
                    document.ExtractionMessage = "The source file is unavailable."
                    Return document
                End If
                Dim extension As System.String = System.IO.Path.GetExtension(fullPath).ToLowerInvariant()
                If Not IsSupportedFile(fullPath, context IsNot Nothing AndAlso context.INI_AllowLegacyDocFiles) Then
                    document.ExtractionStatus = "error"
                    document.ExtractionMessage = "This file type is not enabled by the document text readers."
                    Return document
                End If
                System.IO.Directory.CreateDirectory(staging)
                Dim snapshot As System.String = System.IO.Path.Combine(staging, "source" & extension)
                Await System.Threading.Tasks.Task.Run(
                    Sub() CaptureSource(fullPath, snapshot, document, cancellationToken), cancellationToken).ConfigureAwait(False)

                Dim artifact As System.String = Await System.Threading.Tasks.Task.Run(Function() ReviewArtifactReason(snapshot, cancellationToken), cancellationToken).ConfigureAwait(False)
                If artifact.Length > 0 Then
                    document.ExtractionStatus = "error"
                    document.ExtractionMessage = "Tabular Review excludes its own " & artifact & " from source documents. Choose an original source file."
                    Return document
                End If
                Dim warnings As New System.Collections.Generic.List(Of System.String)()
                Dim complete As System.Nullable(Of System.Boolean) = Nothing
                Dim text As System.String = System.String.Empty
                Dim readError As System.String = Nothing
                If extension = ".pdf" Then
                    Dim isolatedContext As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext = context
                    If context IsNot Nothing Then isolatedContext = Global.SharedLibrary.SharedLibrary.SharedMethods.CreateIsolatedModelCallContext(context)
                    ' This is exactly the PDF reader used by Data Collector. A second OCR
                    ' pass is unnecessary because its page-aligned text is retained here.
                    Dim pdf As Global.SharedLibrary.SharedLibrary.SharedMethods.PdfReadResult =
                        Await Global.SharedLibrary.SharedLibrary.SharedMethods.ReadPdfAsTextEx(
                            pdfPath:=snapshot, ReturnErrorInsteadOfEmpty:=False, DoOCR:=doOcr,
                            AskUser:=False, context:=isolatedContext, ShowOcrProgressWindow:=False,
                            CancellationToken:=cancellationToken).ConfigureAwait(False)
                    text = If(pdf.Content, System.String.Empty)
                    complete = pdf.ExtractionComplete
                    If pdf.ExtractionWarnings IsNot Nothing Then warnings.AddRange(pdf.ExtractionWarnings)
                    If Not System.String.IsNullOrEmpty(pdf.ErrorCode) Then
                        readError = pdf.ErrorCode & ": " & pdf.ErrorMessage
                        complete = False
                    Else
                        readError = System.String.Empty
                    End If
                    If pdf.OcrUsed Then warnings.Add("OCR text is an interpretation of the source image. Native page highlighting is only used when the quote also matches a text layer.")
                    If pdf.PageTexts IsNot Nothing AndAlso pdf.PageTexts.Count > 0 Then
                        For pageIndex As System.Int32 = 0 To pdf.PageTexts.Count - 1
                            AppendSegments(document, If(pdf.PageTexts(pageIndex), System.String.Empty),
                                "Page " & (pageIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), "pdf", pageIndex + 1)
                        Next
                        If pdf.PageCount.HasValue AndAlso pdf.PageTexts.Count <> pdf.PageCount.Value Then
                            complete = False
                            warnings.Add("The number of page text records differs from the PDF page count.")
                        End If
                        ' Existing reader versions may retain recovered text outside the
                        ' page array. Preserve it as an unlocated segment instead of losing it.
                        Dim joinedPages As System.String = System.String.Join(System.Environment.NewLine, pdf.PageTexts)
                        If text.Length > 0 AndAlso NormalizeWhitespace(joinedPages) <> NormalizeWhitespace(text) Then
                            AppendSegments(document, text, "Reader text (page association unavailable)", "text", 0)
                            complete = False
                            warnings.Add("The reader returned additional or differently assembled text; the complete returned text was retained with no invented page association.")
                        End If
                    Else
                        AppendSegments(document, text, "PDF text (page association unavailable)", "text", 0)
                        If text.Length > 0 Then
                            complete = False
                            warnings.Add("The text reader did not provide stable page associations.")
                        End If
                    End If
                    warnings.Add("Extraction basis: " & If(pdf.ExtractionCoverageBasis, "unverified") & ".")
                ElseIf extension = ".doc" Then
                    If Not Global.SharedLibrary.SharedLibrary.SharedMethods.HasHostUiDispatcher Then Throw New System.InvalidOperationException("Legacy Word text extraction requires the Office UI dispatcher.")
                    text = Await Global.SharedLibrary.SharedLibrary.SharedMethods.PostToHostUiAsync(Of System.String)(
                        Function() Global.SharedLibrary.SharedLibrary.SharedMethods.ReadWordDocument(snapshot, False, readError)).ConfigureAwait(False)
                    If readError Is Nothing Then readError = "The legacy Word reader did not report a structured completion status."
                    If readError.Length > 0 Then text = System.String.Empty
                    AppendSegments(document, text, "Word text", "word", 0)
                    warnings.Add("Legacy Word extraction reads the main document text; headers, notes, images and embedded objects require source review.")
                Else
                    text = Await System.Threading.Tasks.Task.Run(
                        Function() ReadExistingTextReader(snapshot, extension, complete, readError), cancellationToken).ConfigureAwait(False)
                    If readError Is Nothing Then readError = "The text reader did not report a structured completion status."
                    If readError.Length > 0 Then
                        ' Several legacy-compatible readers return their error string
                        ' as well as readError. That string is not source evidence.
                        text = System.String.Empty
                        complete = False
                    End If
                    If extension = ".xlsx" OrElse extension = ".xlsm" Then
                        AppendStructuredSegments(document, text, "excel")
                        warnings.Add("The workbook reader supplies cell values and formulas. Charts, images, embedded objects and full layout coverage are not verified.")
                    ElseIf extension = ".pptx" OrElse extension = ".pptm" Then
                        AppendStructuredSegments(document, text, "powerpoint")
                        warnings.Add("The presentation reader supplies slide text and notes. Images, charts and embedded objects require source review.")
                    Else
                        AppendSegments(document, text, If(extension = ".docx" OrElse extension = ".docm", "Word text", "Text"),
                            If(extension = ".docx" OrElse extension = ".docm", "word", "text"), 0)
                    End If
                End If
                cancellationToken.ThrowIfCancellationRequested()
                Dim hasContent As System.Boolean = False
                For Each segment As ReviewSourceSegment In document.Segments
                    If Not System.String.IsNullOrWhiteSpace(segment.Text) Then
                        hasContent = True
                        Exit For
                    End If
                Next
                If Not System.String.IsNullOrEmpty(readError) Then warnings.Insert(0, readError)
                If Not hasContent Then
                    document.Segments.Clear()
                    document.ExtractionStatus = "error"
                    warnings.Insert(0, "No readable text was extracted. This document has not been assessed.")
                Else
                    document.ExtractionStatus = If(complete.GetValueOrDefault() AndAlso System.String.IsNullOrEmpty(readError), "complete", "partial")
                    If Not complete.HasValue Then warnings.Add("The existing reader does not establish complete source coverage. Findings must be checked against the original document.")
                End If
                document.ExtractionMessage = System.String.Join(System.Environment.NewLine, DistinctWarnings(warnings))
                If document.ExtractionStatus = "complete" OrElse document.ExtractionStatus = "partial" Then
                    document.StoredExtractionStatus = document.ExtractionStatus
                    document.StoredExtractionMessage = document.ExtractionMessage
                End If
                document.TextHash = ReviewContract.Digest(ReviewContract.Serialize(document.Segments))
                System.Diagnostics.Trace.WriteLine("[TabularReview] extraction " & document.Id & "; status=" & document.ExtractionStatus & "; " & document.ExtractionMessage)
                Return document
            Catch ex As System.OperationCanceledException
                Throw
            Catch ex As System.Exception
                document.ExtractionStatus = "error"
                document.ExtractionMessage = ex.Message
                System.Diagnostics.Trace.WriteLine("[TabularReview] source extraction failed: " & ex.ToString())
                Return document
            Finally
                DeleteTemporaryDirectory(staging)
            End Try
        End Function

        Private Shared Function ReadExistingTextReader(path As System.String, extension As System.String,
                ByRef complete As System.Nullable(Of System.Boolean), ByRef readError As System.String) As System.String
            Select Case extension
                Case ".docx", ".docm"
                    Return Global.SharedLibrary.SharedLibrary.SharedMethods.ReadDocxSandboxed(path, False, readError, complete)
                Case ".xlsx", ".xlsm"
                    Return Global.SharedLibrary.SharedLibrary.SharedMethods.ReadXlsxSandboxed(path, silent:=True, askWorksheetSelection:=False, readError:=readError)
                Case ".pptx", ".pptm"
                    Return Global.SharedLibrary.SharedLibrary.SharedMethods.ReadPptxSandboxed(path, readError)
                Case ".rtf"
                    Dim localError As System.String = Nothing
                    Dim result As System.String = Global.SharedLibrary.SharedLibrary.SharedMethods.RunOnStaThread(Of System.String)(
                        Function() Global.SharedLibrary.SharedLibrary.SharedMethods.ReadRtfAsText(path, False, localError), "Red Ink review RTF reader", alwaysUseWorker:=True)
                    readError = localError
                    complete = System.String.IsNullOrEmpty(readError)
                    Return result
                Case ".eml"
                    Return Global.SharedLibrary.SharedLibrary.SharedMethods.ReadEmlSandboxed(path, readError)
                Case ".msg"
                    Return Global.SharedLibrary.SharedLibrary.SharedMethods.ReadMsgSandboxed(path, readError:=readError)
                Case Else
                    Dim result As System.String = Global.SharedLibrary.SharedLibrary.SharedMethods.ReadTextFile(path, False, readError)
                    complete = System.String.IsNullOrEmpty(readError)
                    Return result
            End Select
        End Function

        Private Shared Sub CaptureSource(path As System.String, snapshot As System.String, document As ReviewDocument,
                                         cancellationToken As System.Threading.CancellationToken)
            Using input As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                Using output As New System.IO.FileStream(snapshot, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None)
                    Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                        Dim buffer(81919) As System.Byte
                        Dim read As System.Int32
                        Do
                            cancellationToken.ThrowIfCancellationRequested()
                            read = input.Read(buffer, 0, buffer.Length)
                            If read = 0 Then Exit Do
                            output.Write(buffer, 0, read)
                            algorithm.TransformBlock(buffer, 0, read, buffer, 0)
                        Loop
                        algorithm.TransformFinalBlock(New System.Byte() {}, 0, 0)
                        document.SourceHash = ToHashString(algorithm.Hash)
                    End Using
                    document.SourceLength = input.Length
                    document.SourceLastWriteUtc = System.IO.File.GetLastWriteTimeUtc(path).ToString("o", System.Globalization.CultureInfo.InvariantCulture)
                End Using
            End Using
        End Sub

        Private Shared Function ToHashString(bytes As System.Byte()) As System.String
            Return System.BitConverter.ToString(bytes).Replace("-", System.String.Empty).ToLowerInvariant()
        End Function

        Private Shared Function RelativeSourcePath(fullPath As System.String, root As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(root) Then Return System.IO.Path.GetFileName(fullPath)
            Dim prefix As System.String = System.IO.Path.GetFullPath(root).TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar) & System.IO.Path.DirectorySeparatorChar
            ' A host-owned file picker may explicitly relink a moved source outside
            ' the collection root. Preserve that full path; do not invent a relative
            ' path that could later resolve to a different document.
            If Not fullPath.StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase) Then Return fullPath
            Return fullPath.Substring(prefix.Length)
        End Function

        Private Shared Function DeriveDocumentNumber(relativePath As System.String) As System.String
            Dim withoutExtension As System.String = relativePath.Substring(0, relativePath.Length - System.IO.Path.GetExtension(relativePath).Length)
            Dim components As System.String() = withoutExtension.Replace("\", "/").Split("/"c)
            Dim number As System.String = System.String.Empty
            For Each component As System.String In components
                Dim match As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(component, "^\s*([0-9]+(?:[._-][0-9]+)*)", System.Text.RegularExpressions.RegexOptions.CultureInvariant)
                If Not match.Success Then Continue For
                Dim part As System.String = match.Groups(1).Value.Replace("_", ".").Replace("-", ".")
                If number.Length = 0 OrElse part.StartsWith(number & ".", System.StringComparison.Ordinal) Then
                    number = part
                ElseIf part <> number AndAlso Not number.EndsWith("." & part, System.StringComparison.Ordinal) Then
                    number &= "." & part
                End If
            Next
            Return ReviewContract.ShortDocumentReference(If(number.Length > 0, number, withoutExtension.Replace("\", "/")), ReviewContract.Digest(relativePath))
        End Function

        ''' <summary>Segmentation is lossless. Long units are split without overlap;
        ''' their order and every source character are retained in the workbook.</summary>
        Private Shared Sub AppendSegments(document As ReviewDocument, text As System.String,
                                           locator As System.String, kind As System.String, page As System.Int32)
            Dim source As System.String = If(text, System.String.Empty)
            Const targetLength As System.Int32 = 8000
            If source.Length = 0 Then
                AddSegment(document, source, locator, kind, page)
                Return
            End If
            Dim offset As System.Int32 = 0
            Dim part As System.Int32 = 0
            While offset < source.Length
                Dim length As System.Int32 = System.Math.Min(targetLength, source.Length - offset)
                If offset + length < source.Length Then
                    Dim split As System.Int32 = source.LastIndexOf(Microsoft.VisualBasic.ChrW(10), offset + length - 1, System.Math.Min(length, 3000))
                    If split >= offset + (length \ 2) Then length = split - offset + 1
                    If length > 1 AndAlso System.Char.IsHighSurrogate(source(offset + length - 1)) Then length -= 1
                End If
                part += 1
                AddSegment(document, source.Substring(offset, length), locator & If(source.Length > targetLength, " / part " & part.ToString(System.Globalization.CultureInfo.InvariantCulture), System.String.Empty), kind, page)
                offset += length
            End While
        End Sub

        Private Shared Sub AddSegment(document As ReviewDocument, text As System.String,
                                       locator As System.String, kind As System.String, page As System.Int32)
            document.Segments.Add(New ReviewSourceSegment With {
                .Id = "s" & (document.Segments.Count + 1).ToString("D6", System.Globalization.CultureInfo.InvariantCulture),
                .Text = text, .Locator = locator, .Kind = kind, .PageNumber = page
            })
        End Sub

        Private Shared Sub AppendStructuredSegments(document As ReviewDocument, text As System.String, kind As System.String)
            Dim source As System.String = If(text, System.String.Empty)
            Dim pattern As System.String = If(kind = "excel", "(?m)^=== Sheet: (.*?) ===\r?$", "(?m)^=== Slide ([0-9]+) ===\r?$")
            Dim matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(source, pattern, System.Text.RegularExpressions.RegexOptions.CultureInvariant)
            If matches.Count = 0 Then
                AppendSegments(document, source, If(kind = "excel", "Workbook text", "Presentation text"), kind, 0)
                Return
            End If
            If matches(0).Index > 0 Then AppendSegments(document, source.Substring(0, matches(0).Index), "Reader preamble", kind, 0)
            For index As System.Int32 = 0 To matches.Count - 1
                Dim current As System.Text.RegularExpressions.Match = matches(index)
                Dim ending As System.Int32 = If(index + 1 < matches.Count, matches(index + 1).Index, source.Length)
                Dim section As System.String = source.Substring(current.Index, ending - current.Index)
                If kind = "powerpoint" Then
                    Dim slide As System.Int32
                    System.Int32.TryParse(current.Groups(1).Value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, slide)
                    AppendSegments(document, section, "Slide " & current.Groups(1).Value & " (including notes)", kind, slide)
                Else
                    Dim sheet As System.String = current.Groups(1).Value
                    ' Retain the reader's sheet header and use one addressable unit per
                    ' cell row. Embedded newlines in a value remain in that same unit.
                    Dim cells As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(section, "(?m)^[A-Z]{1,3}[1-9][0-9]*\t", System.Text.RegularExpressions.RegexOptions.CultureInvariant)
                    If cells.Count = 0 Then
                        AppendSegments(document, section, "Sheet " & sheet, kind, 0)
                    Else
                        If cells(0).Index > 0 Then AppendSegments(document, section.Substring(0, cells(0).Index), "Sheet " & sheet, kind, 0)
                        For cellIndex As System.Int32 = 0 To cells.Count - 1
                            Dim cell As System.Text.RegularExpressions.Match = cells(cellIndex)
                            Dim cellEnd As System.Int32 = If(cellIndex + 1 < cells.Count, cells(cellIndex + 1).Index, section.Length)
                            AppendSegments(document, section.Substring(cell.Index, cellEnd - cell.Index),
                                "Sheet " & sheet & " / " & cell.Value.TrimEnd(Microsoft.VisualBasic.ChrW(9)), kind, 0)
                        Next
                    End If
                End If
            Next
        End Sub

        Private Shared Function NormalizeWhitespace(text As System.String) As System.String
            Return System.Text.RegularExpressions.Regex.Replace(If(text, System.String.Empty), "\s+", " ").Trim()
        End Function

        Private Shared Function DistinctWarnings(values As System.Collections.Generic.IEnumerable(Of System.String)) As System.Collections.Generic.List(Of System.String)
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim result As New System.Collections.Generic.List(Of System.String)()
            For Each value As System.String In values
                If Not System.String.IsNullOrWhiteSpace(value) AndAlso seen.Add(value) Then result.Add(value)
            Next
            Return result
        End Function

        Private Shared Sub DeleteTemporaryDirectory(path As System.String)
            Try
                If System.IO.Directory.Exists(path) Then System.IO.Directory.Delete(path, True)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine("[TabularReview] source snapshot cleanup failed: " & ex.Message)
            End Try
        End Sub
    End Class
End Namespace
