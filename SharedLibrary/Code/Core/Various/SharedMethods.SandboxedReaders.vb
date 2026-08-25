' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.SandboxedReaders.vb
' Purpose: Centralised, dependency-free ("sandboxed") text extraction routines for
'          common document and mail formats. Implements ZIP+XML parsing for OpenXML
'          packages (.docx, .xlsx, .pptx), RFC‑compliant parsing for .eml, and a
'          minimal OLE/MAPI reader for .msg files — all without Office interop or
'          third‑party binary libraries. Outputs are formatted to remain compatible
'          with the project's historical `Extract*` helpers so downstream callers
'          and LLM pipelines receive the same plain‑text layout.
'
' Key responsibilities:
'  - DOCX: paragraph‑per‑line plain text, with optional headers/footers/notes and
'          descriptive table sections for robust LL model consumption; old version is
'          contained below, a new richer version with heading numbering is separate
'  - XLSX: produces per‑cell lines of the form
'          "{addr}\tFORMULA:={formula}\tVALUE: {value}" and "=== Sheet: {name} ==="
'          headers to match legacy extractor output.
'  - PPTX: slide headers "=== Slide {n} ===", shape text lines and "--- Notes ---".
'  - EML: RFC 2822 header parsing, MIME boundary handling, inline attachment extraction.
'  - MSG: lightweight OLE Compound File + MAPI property reader with safe fallbacks.
'
' Design notes:
'  - Security: intentionally avoids COM and unmanaged dependencies to reduce attack
'    surface when extracting content from untrusted files. Legacy formats (e.g.
'    binary .doc) are disabled by default and gated by configuration.
'  - Compatibility: strives to preserve the exact output shapes expected by other
'    components (format strings, section headers) to enable drop‑in replacements.
'  - Resilience: defensive parsing, graceful fallbacks, bounded recursion for nested
'    attachments, and temporary work directories that are cleaned up on completion.
'  - Performance: stream / XmlDocument usage tuned for typical document sizes; large
'    binary streams are processed via temporary files to limit memory pressure.
'
' External Dependencies:
'  - System.IO.Compression (ZipArchive/ZipFile)
'  - System.Xml (XmlDocument, XmlNamespaceManager)
'  - System.Text / System.IO
'
' Output compatibility and conventions:
'  - Keep output strings stable (headers, separators, cell formatting) to match
'    `ExtractWordText`, `ExtractExcelText`, and `ExtractPowerPointText`.
'  - Error messages are returned as plain text beginning with "Error:" to allow
'    callers to detect failures without throwing exceptions in common flows.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Xml

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Cached value for allowing legacy document files. Set during LoadConfig.
        ''' </summary>
        Public Shared INI_AllowLegacyDocFiles_Cached As Boolean = False

        ' ═══════════════════════════════════════════════════════════════════════
        '  DOCX — Sandboxed
        ' ═══════════════════════════════════════════════════════════════════════

        Public Shared Function ReadDocxSandboxed(docxPath As String,
                                                 Optional returnMarkdown As System.Boolean = False) As String
            Return DocxTextExtractor.ReadDocxSandboxed(docxPath, returnMarkdown)
        End Function

        Public Shared Function ReadPdfMarkdownSandboxed(pdfPath As String) As String
            Return PdfMarkdownExtractor.ReadPdfAsMarkdown(pdfPath)
        End Function


        ' The following is a previous version

        Private Const SB_WordNs As String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        Public Shared Function oldReadDocxSandboxed(docxPath As String) As String
            If System.String.IsNullOrWhiteSpace(docxPath) OrElse Not System.IO.File.Exists(docxPath) Then
                Return "Error: File not found."
            End If

            Dim tempDir As String =
        System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            "ri_docx_" & System.Guid.NewGuid().ToString("N")
        )

            Try
                System.IO.Compression.ZipFile.ExtractToDirectory(docxPath, tempDir)

                Dim wordDir As String = System.IO.Path.Combine(tempDir, "word")
                Dim documentXmlPath As String = System.IO.Path.Combine(wordDir, "document.xml")

                If Not System.IO.File.Exists(documentXmlPath) Then
                    Return "Error: Not a valid .docx file (missing word/document.xml)."
                End If

                Dim xmlDoc As New System.Xml.XmlDocument()
                xmlDoc.PreserveWhitespace = True
                xmlDoc.Load(documentXmlPath)

                Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
                nsMgr.AddNamespace("w", SB_WordNs)

                Dim sb As New System.Text.StringBuilder(4096)

                ' ── Main body content in original order: paragraphs and tables ──
                Dim bodyNode As System.Xml.XmlNode = xmlDoc.SelectSingleNode("//w:body", nsMgr)
                Dim tableIndex As Integer = 0

                If bodyNode IsNot Nothing Then
                    For Each childNode As System.Xml.XmlNode In bodyNode.ChildNodes

                        If childNode.NamespaceURI <> SB_WordNs Then
                            Continue For
                        End If

                        Select Case childNode.LocalName

                            Case "p"
                                Dim paraText As String = ExtractDocxParagraphText(childNode, nsMgr)

                                If paraText.Length > 0 Then
                                    sb.AppendLine(paraText)
                                Else
                                    ' Empty paragraph = blank line
                                    sb.AppendLine()
                                End If

                            Case "tbl"
                                tableIndex += 1
                                AppendDocxTableForLlm(childNode, nsMgr, sb, tableIndex, 0)

                        End Select
                    Next
                End If

                ' ── Optional: headers, footers, footnotes, endnotes ──
                If DocxIncludeHeaderFooterFootnotes AndAlso System.IO.Directory.Exists(wordDir) Then

                    ' Headers
                    ExtractDocxSubParts(wordDir, "header*.xml", "Header", sb)

                    ' Footers
                    ExtractDocxSubParts(wordDir, "footer*.xml", "Footer", sb)

                    ' Footnotes, skip id="0" = separator, id="-1" = continuation separator
                    ExtractDocxNotes(wordDir, "footnotes.xml", "w:footnote", "Footnote", sb)

                    ' Endnotes, skip id="0" and id="-1"
                    ExtractDocxNotes(wordDir, "endnotes.xml", "w:endnote", "Endnote", sb)
                End If

                Dim result As String = sb.ToString().TrimEnd()

                Return If(
            System.String.IsNullOrWhiteSpace(result),
            "Error: No text content found in .docx.",
            result
        )

            Catch ex As System.Exception
                Return "Error reading .docx: " & ex.Message

            Finally
                Try
                    If System.IO.Directory.Exists(tempDir) Then
                        System.IO.Directory.Delete(tempDir, True)
                    End If
                Catch
                End Try
            End Try
        End Function


        Private Shared Function ExtractDocxParagraphText(
    paraNode As System.Xml.XmlNode,
    nsMgr As System.Xml.XmlNamespaceManager
) As String

            Dim paraText As New System.Text.StringBuilder()

            Dim runs As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:r", nsMgr)

            If runs IsNot Nothing Then
                For Each runNode As System.Xml.XmlNode In runs

                    If DocxIncludeHeaderFooterFootnotes Then
                        Dim fnRef As System.Xml.XmlNode = runNode.SelectSingleNode("w:footnoteReference", nsMgr)

                        If fnRef IsNot Nothing Then
                            Dim fnId As String = GetWordAttributeValue(fnRef, "id")

                            If Not System.String.IsNullOrWhiteSpace(fnId) AndAlso fnId <> "0" Then
                                paraText.Append(" [Footnote " & fnId & "]")
                            End If
                        End If

                        Dim enRef As System.Xml.XmlNode = runNode.SelectSingleNode("w:endnoteReference", nsMgr)

                        If enRef IsNot Nothing Then
                            Dim enId As String = GetWordAttributeValue(enRef, "id")

                            If Not System.String.IsNullOrWhiteSpace(enId) AndAlso enId <> "0" Then
                                paraText.Append(" [Endnote " & enId & "]")
                            End If
                        End If
                    End If

                    For Each runChild As System.Xml.XmlNode In runNode.ChildNodes

                        If runChild.NamespaceURI <> SB_WordNs Then
                            Continue For
                        End If

                        Select Case runChild.LocalName

                            Case "t"
                                paraText.Append(runChild.InnerText)

                            Case "tab"
                                paraText.Append(vbTab)

                            Case "br", "cr"
                                paraText.AppendLine()

                            Case "noBreakHyphen"
                                paraText.Append(ChrW(&H2011))

                            Case "softHyphen"
                                paraText.Append(ChrW(&HAD))

                        End Select
                    Next
                Next
            End If

            Return paraText.ToString().Trim()
        End Function


        Private Shared Sub AppendDocxTableForLlm(
    tableNode As System.Xml.XmlNode,
    nsMgr As System.Xml.XmlNamespaceManager,
    sb As System.Text.StringBuilder,
    tableIndex As Integer,
    nestingLevel As Integer
)

            Dim indent As String = New System.String(" "c, nestingLevel * 2)
            Dim tableNumberText As String = tableIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)

            sb.AppendLine()
            sb.AppendLine(indent & "[Table " & tableNumberText & "]")

            Dim rows As System.Xml.XmlNodeList = tableNode.SelectNodes("w:tr", nsMgr)

            If rows Is Nothing OrElse rows.Count = 0 Then
                sb.AppendLine(indent & "[Empty table]")
                sb.AppendLine(indent & "[/Table " & tableNumberText & "]")
                sb.AppendLine()
                Return
            End If

            Dim rowIndex As Integer = 0

            For Each rowNode As System.Xml.XmlNode In rows
                rowIndex += 1

                Dim cells As System.Xml.XmlNodeList = rowNode.SelectNodes("w:tc", nsMgr)

                If cells Is Nothing OrElse cells.Count = 0 Then
                    sb.AppendLine(
                indent &
                "Row " &
                rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                ": [empty row]"
            )

                    Continue For
                End If

                Dim visualColumnIndex As Integer = 1
                Dim physicalCellIndex As Integer = 0

                For Each cellNode As System.Xml.XmlNode In cells
                    physicalCellIndex += 1

                    Dim gridSpan As Integer = GetDocxGridSpan(cellNode, nsMgr)
                    Dim verticalMerge As String = GetDocxVerticalMerge(cellNode, nsMgr)
                    Dim cellText As String = ExtractDocxCellText(cellNode, nsMgr, tableIndex, nestingLevel + 1)

                    Dim rowText As String = rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    Dim physicalCellText As String = physicalCellIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    Dim startColumnText As String = visualColumnIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    Dim endColumnText As String = (visualColumnIndex + gridSpan - 1).ToString(System.Globalization.CultureInfo.InvariantCulture)

                    Dim cellLabel As String =
                indent &
                "Row " &
                rowText &
                ", Cell " &
                physicalCellText &
                ", Column " &
                startColumnText

                    If gridSpan > 1 Then
                        cellLabel &= "-" & endColumnText & " spanning " & gridSpan.ToString(System.Globalization.CultureInfo.InvariantCulture) & " columns"
                    End If

                    If Not System.String.IsNullOrWhiteSpace(verticalMerge) Then
                        If verticalMerge = "restart" Then
                            cellLabel &= ", starts vertical merge"
                        Else
                            cellLabel &= ", continues vertical merge from row above"
                        End If
                    End If

                    sb.AppendLine(cellLabel & ": " & cellText)

                    visualColumnIndex += gridSpan
                Next
            Next

            sb.AppendLine(indent & "[/Table " & tableNumberText & "]")
            sb.AppendLine()
        End Sub


        Private Shared Function ExtractDocxCellText(
    cellNode As System.Xml.XmlNode,
    nsMgr As System.Xml.XmlNamespaceManager,
    tableIndex As Integer,
    nestingLevel As Integer
) As String

            Dim parts As New System.Collections.Generic.List(Of String)()
            Dim nestedTableIndex As Integer = 0

            For Each childNode As System.Xml.XmlNode In cellNode.ChildNodes

                If childNode.NamespaceURI <> SB_WordNs Then
                    Continue For
                End If

                Select Case childNode.LocalName

                    Case "p"
                        Dim paragraphText As String = ExtractDocxParagraphText(childNode, nsMgr)

                        If Not System.String.IsNullOrWhiteSpace(paragraphText) Then
                            parts.Add(paragraphText)
                        End If

                    Case "tbl"
                        nestedTableIndex += 1

                        Dim nestedBuilder As New System.Text.StringBuilder()

                        AppendDocxTableForLlm(
                    childNode,
                    nsMgr,
                    nestedBuilder,
                    tableIndex * 1000 + nestedTableIndex,
                    nestingLevel
                )

                        parts.Add(nestedBuilder.ToString().Trim())

                End Select
            Next

            If parts.Count = 0 Then
                Return "[empty]"
            End If

            Return System.String.Join(" | ", parts)
        End Function


        Private Shared Function GetDocxGridSpan(
    cellNode As System.Xml.XmlNode,
    nsMgr As System.Xml.XmlNamespaceManager
) As Integer

            Dim gridSpanNode As System.Xml.XmlNode = cellNode.SelectSingleNode("w:tcPr/w:gridSpan", nsMgr)

            If gridSpanNode Is Nothing Then
                Return 1
            End If

            Dim valueText As String = GetWordAttributeValue(gridSpanNode, "val")
            Dim result As Integer

            If System.Int32.TryParse(valueText, result) AndAlso result > 1 Then
                Return result
            End If

            Return 1
        End Function


        Private Shared Function GetDocxVerticalMerge(
    cellNode As System.Xml.XmlNode,
    nsMgr As System.Xml.XmlNamespaceManager
) As String

            Dim vMergeNode As System.Xml.XmlNode = cellNode.SelectSingleNode("w:tcPr/w:vMerge", nsMgr)

            If vMergeNode Is Nothing Then
                Return System.String.Empty
            End If

            Dim valueText As String = GetWordAttributeValue(vMergeNode, "val")

            If System.String.IsNullOrWhiteSpace(valueText) Then
                Return "continue"
            End If

            Return valueText
        End Function


        Private Shared Function GetWordAttributeValue(
    node As System.Xml.XmlNode,
    localName As String
) As String

            If node Is Nothing OrElse node.Attributes Is Nothing Then
                Return System.String.Empty
            End If

            Dim attr As System.Xml.XmlNode = node.Attributes.GetNamedItem(localName, SB_WordNs)

            If attr IsNot Nothing Then
                Return attr.Value
            End If

            attr = node.Attributes.GetNamedItem("w:" & localName)

            If attr IsNot Nothing Then
                Return attr.Value
            End If

            attr = node.Attributes.GetNamedItem(localName)

            If attr IsNot Nothing Then
                Return attr.Value
            End If

            Return System.String.Empty
        End Function


        ''' <summary>
        ''' Extracts text from DOCX sub-parts matching a file pattern (e.g., <c>header*.xml</c>, <c>footer*.xml</c>).
        ''' Each file produces a labeled section with its paragraphs.
        ''' </summary>
        Private Shared Sub ExtractDocxSubParts(wordDir As String, filePattern As String,
                                               sectionLabel As String, sb As StringBuilder)
            Try
                Dim files = Directory.GetFiles(wordDir, filePattern).OrderBy(Function(f) f).ToArray()
                If files.Length = 0 Then Return

                For Each filePath In files
                    Dim partDoc As New XmlDocument()
                    partDoc.PreserveWhitespace = True
                    partDoc.Load(filePath)

                    Dim partNs As New XmlNamespaceManager(partDoc.NameTable)
                    partNs.AddNamespace("w", SB_WordNs)

                    Dim partParagraphs = partDoc.SelectNodes("//w:p", partNs)
                    If partParagraphs Is Nothing OrElse partParagraphs.Count = 0 Then Continue For

                    ' Collect text first to check if there's any real content
                    Dim partText As New StringBuilder()
                    For Each para As XmlNode In partParagraphs
                        Dim tNodes = para.SelectNodes(".//w:t", partNs)
                        If tNodes IsNot Nothing AndAlso tNodes.Count > 0 Then
                            Dim paraLine As New StringBuilder()
                            For Each tNode As XmlNode In tNodes
                                paraLine.Append(tNode.InnerText)
                            Next
                            Dim lineText = paraLine.ToString().Trim()
                            If lineText.Length > 0 Then
                                partText.AppendLine(lineText)
                            End If
                        End If
                    Next

                    If partText.Length > 0 Then
                        ' Derive a display label: "Header 1", "Footer 2", etc.
                        Dim fileLabel = Path.GetFileNameWithoutExtension(filePath)
                        ' Extract trailing number from e.g. "header1" → "1"
                        Dim numPart = ""
                        For i = fileLabel.Length - 1 To 0 Step -1
                            If Char.IsDigit(fileLabel(i)) Then
                                numPart = fileLabel(i) & numPart
                            Else
                                Exit For
                            End If
                        Next

                        sb.AppendLine()
                        sb.AppendLine($"--- {sectionLabel}{If(numPart.Length > 0, " " & numPart, "")} ---")
                        sb.Append(partText.ToString().TrimEnd())
                        sb.AppendLine()
                    End If
                Next
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Extracts text from DOCX footnotes or endnotes XML. Skips system notes (id 0 and -1)
        ''' which are separator/continuation placeholders. Each note is output with its ID
        ''' so it can be cross-referenced with <c>[Footnote n]</c> / <c>[Endnote n]</c> markers
        ''' inserted into the body text.
        ''' </summary>
        Private Shared Sub ExtractDocxNotes(wordDir As String, fileName As String,
                                             noteElementName As String, sectionLabel As String,
                                             sb As StringBuilder)
            Dim filePath = Path.Combine(wordDir, fileName)
            If Not File.Exists(filePath) Then Return

            Try
                Dim notesDoc As New XmlDocument()
                notesDoc.PreserveWhitespace = True
                notesDoc.Load(filePath)

                Dim notesNs As New XmlNamespaceManager(notesDoc.NameTable)
                notesNs.AddNamespace("w", SB_WordNs)

                Dim noteNodes = notesDoc.SelectNodes($"//{ noteElementName}", notesNs)
                If noteNodes Is Nothing OrElse noteNodes.Count = 0 Then Return

                Dim notesCollected As New StringBuilder()
                Dim noteCount As Integer = 0

                For Each noteNode As XmlNode In noteNodes
                    ' Skip system separator/continuation notes (id="0" or id="-1")
                    Dim noteId = noteNode.Attributes?("w:id")?.Value
                    If noteId = "0" OrElse noteId = "-1" Then Continue For

                    Dim noteParagraphs = noteNode.SelectNodes("w:p", notesNs)
                    If noteParagraphs Is Nothing OrElse noteParagraphs.Count = 0 Then Continue For

                    Dim noteText As New StringBuilder()
                    For Each para As XmlNode In noteParagraphs
                        ' Skip the footnoteRef/endnoteRef marker runs (the auto-number "1", "2" etc.)
                        Dim tNodes = para.SelectNodes(".//w:r[not(w:footnoteRef) and not(w:endnoteRef)]/w:t", notesNs)
                        If tNodes IsNot Nothing AndAlso tNodes.Count > 0 Then
                            For Each tNode As XmlNode In tNodes
                                noteText.Append(tNode.InnerText)
                            Next
                        End If
                    Next

                    Dim trimmedNote = noteText.ToString().Trim()
                    If trimmedNote.Length > 0 Then
                        notesCollected.AppendLine($"  [{sectionLabel} {If(noteId, "?")}] {trimmedNote}")
                        noteCount += 1
                    End If
                Next

                If noteCount > 0 Then
                    sb.AppendLine()
                    sb.AppendLine($"--- {sectionLabel}s ---")
                    sb.Append(notesCollected.ToString().TrimEnd())
                    sb.AppendLine()
                End If
            Catch
            End Try
        End Sub


        ' ═══════════════════════════════════════════════════════════════════════
        '  XLSX — Sandboxed
        ' ═══════════════════════════════════════════════════════════════════════

        Private Const SB_XlsxNs As String = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        Private Const SB_XlsxRelNs As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        ''' <summary>
        ''' Represents a readable worksheet discovered in an .xlsx workbook.
        ''' </summary>
        Private Structure XlsxSheetEntry
            Public Name As String
            Public SheetXmlPath As String
        End Structure

        ''' <summary>
        ''' Represents a worksheet declared in workbook.xml, whether or not its XML part
        ''' could be resolved immediately.
        ''' </summary>
        Private Structure XlsxDeclaredSheetEntry
            Public Name As String
            Public SheetXmlPath As String
        End Structure

        Public Const XlsxSelectionCancelledMarker As String = "__RI_XLSX_SELECTION_CANCELLED__"

        ''' <summary>
        ''' Extracts text from an .xlsx file without COM interop.
        ''' Output format matches <c>ExtractExcelText</c>: <c>{addr}\tFORMULA:={formula}\tVALUE: {value}</c> per cell,
        ''' with <c>=== Sheet: {name} ===</c> headers.
        ''' </summary>
        ''' <param name="xlsxPath">Absolute path to the .xlsx file.</param>
        ''' <param name="silent">
        ''' When <c>True</c>, suppresses worksheet-selection UI and always loads all readable worksheets.
        ''' </param>
        ''' <param name="askWorksheetSelection">
        ''' When <c>True</c> and <paramref name="silent"/> is <c>False</c>, prompts the user via
        ''' <see cref="SelectValue(IEnumerable(Of SelectionItem), Integer, String, String)"/> to choose
        ''' either all readable worksheets or one specific worksheet when the workbook contains multiple sheets.
        ''' </param>
        ''' <returns>
        ''' Extracted text representation of the selected worksheet set, or an error string on failure.
        ''' Returns an empty string when worksheet selection is canceled.
        ''' </returns>
        Public Shared Function ReadXlsxSandboxed(xlsxPath As String,
                                                 Optional silent As Boolean = True,
                                                 Optional askWorksheetSelection As Boolean = False) As String
            If String.IsNullOrWhiteSpace(xlsxPath) OrElse Not File.Exists(xlsxPath) Then
                Return "Error: File not found."
            End If

            If Not EnsureClosedWorkbookForSandboxedRead(xlsxPath, silent) Then
                Return "Error: The workbook is open in Excel."
            End If

            Dim tempDir As String = Path.Combine(Path.GetTempPath(), "ri_xlsx_" & Guid.NewGuid().ToString("N"))

            Try
                ZipFile.ExtractToDirectory(xlsxPath, tempDir)

                ' ── Load shared strings ──
                Dim sharedStrings As New List(Of String)()
                Dim sstPath = Path.Combine(tempDir, "xl", "sharedStrings.xml")
                If File.Exists(sstPath) Then
                    Dim sstDoc As New XmlDocument()
                    sstDoc.Load(sstPath)
                    Dim sstNs As New XmlNamespaceManager(sstDoc.NameTable)
                    sstNs.AddNamespace("x", SB_XlsxNs)
                    Dim siNodes = sstDoc.SelectNodes("//x:si", sstNs)
                    If siNodes IsNot Nothing Then
                        For Each si As XmlNode In siNodes
                            Dim tNodes = si.SelectNodes(".//x:t", sstNs)
                            Dim cellText As New StringBuilder()
                            If tNodes IsNot Nothing Then
                                For Each tNode As XmlNode In tNodes
                                    cellText.Append(tNode.InnerText)
                                Next
                            End If
                            sharedStrings.Add(cellText.ToString())
                        Next
                    End If
                End If

                ' ── Discover sheet names from workbook.xml ──
                Dim wbPath = Path.Combine(tempDir, "xl", "workbook.xml")
                If Not File.Exists(wbPath) Then Return "Error: Not a valid .xlsx file (missing xl/workbook.xml)."

                Dim wbDoc As New XmlDocument()
                wbDoc.Load(wbPath)
                Dim wbNs As New XmlNamespaceManager(wbDoc.NameTable)
                wbNs.AddNamespace("x", SB_XlsxNs)
                wbNs.AddNamespace("r", SB_XlsxRelNs)

                Dim sheetNodes = wbDoc.SelectNodes("//x:sheets/x:sheet", wbNs)
                If sheetNodes Is Nothing OrElse sheetNodes.Count = 0 Then Return "Error: No sheets found in workbook."

                ' ── Map rId → file path via workbook.xml.rels ──
                Dim relsPath = Path.Combine(tempDir, "xl", "_rels", "workbook.xml.rels")
                Dim ridMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                If File.Exists(relsPath) Then
                    Dim relsDoc As New XmlDocument()
                    relsDoc.Load(relsPath)
                    If relsDoc.DocumentElement IsNot Nothing Then
                        For Each relNode As XmlNode In relsDoc.DocumentElement.ChildNodes
                            If relNode.Attributes Is Nothing Then Continue For
                            Dim id = relNode.Attributes("Id")
                            Dim target = relNode.Attributes("Target")
                            If id IsNot Nothing AndAlso target IsNot Nothing Then
                                ridMap(id.Value) = target.Value
                            End If
                        Next
                    End If
                End If

                Dim declaredSheets As New List(Of XlsxDeclaredSheetEntry)()
                Dim availableSheets As New List(Of XlsxSheetEntry)()
                Dim sheetIdx As Integer = 0

                For Each sheetNode As XmlNode In sheetNodes
                    sheetIdx += 1
                    Dim sheetName = If(sheetNode.Attributes("name")?.Value, "Sheet")
                    Dim rId = If(sheetNode.Attributes("r:id")?.Value, "")

                    Dim sheetXmlPath = ResolveSheetPath(tempDir, rId, ridMap, sheetIdx)

                    declaredSheets.Add(New XlsxDeclaredSheetEntry With {
                        .Name = sheetName,
                        .SheetXmlPath = sheetXmlPath
                    })

                    If Not String.IsNullOrWhiteSpace(sheetXmlPath) AndAlso File.Exists(sheetXmlPath) Then
                        availableSheets.Add(New XlsxSheetEntry With {
                            .Name = sheetName,
                            .SheetXmlPath = sheetXmlPath
                        })
                    End If
                Next

                If availableSheets.Count = 0 Then
                    Return "Error: No readable sheets found in workbook."
                End If

                Dim sheetsToRead As New List(Of XlsxSheetEntry)()
                For Each s In availableSheets
                    sheetsToRead.Add(s)
                Next

                If askWorksheetSelection AndAlso Not silent AndAlso declaredSheets.Count > 1 Then
                    Dim items As New List(Of SelectionItem) From {
                        New SelectionItem("All worksheets", -1)
                    }

                    For i As Integer = 0 To declaredSheets.Count - 1
                        items.Add(New SelectionItem(
                            declaredSheets(i).Name & " (worksheet " & (i + 1).ToString(Globalization.CultureInfo.InvariantCulture) & ")",
                            i + 1))
                    Next

                    Dim selectedValue As Integer = SelectValue(
                        items,
                        -1,
                        "This workbook contains multiple worksheets. Load all worksheets or only one worksheet?",
                        AN & " - Select Worksheet")

                    If selectedValue = 0 Then
                        Return XlsxSelectionCancelledMarker
                    End If

                    If selectedValue > 0 AndAlso selectedValue <= declaredSheets.Count Then
                        Dim selectedSheet = declaredSheets(selectedValue - 1)

                        If String.IsNullOrWhiteSpace(selectedSheet.SheetXmlPath) OrElse
                           Not File.Exists(selectedSheet.SheetXmlPath) Then
                            Return "Error: The selected worksheet could not be read."
                        End If

                        sheetsToRead.Clear()
                        sheetsToRead.Add(New XlsxSheetEntry With {
                            .Name = selectedSheet.Name,
                            .SheetXmlPath = selectedSheet.SheetXmlPath
                        })
                    End If
                End If

                Dim sb As New StringBuilder(4096)

                For Each sheet In sheetsToRead
                    sb.AppendLine("=== Sheet: " & sheet.Name & " ===")

                    Dim sheetDoc As New XmlDocument()
                    sheetDoc.Load(sheet.SheetXmlPath)
                    Dim sheetNs As New XmlNamespaceManager(sheetDoc.NameTable)
                    sheetNs.AddNamespace("x", SB_XlsxNs)

                    Dim cellNodes = sheetDoc.SelectNodes("//x:sheetData/x:row/x:c", sheetNs)
                    If cellNodes IsNot Nothing Then
                        For Each cellNode As XmlElement In cellNodes
                            Dim cellRef = cellNode.GetAttribute("r")
                            If String.IsNullOrEmpty(cellRef) Then Continue For

                            Dim cellType = cellNode.GetAttribute("t")
                            Dim vNode = cellNode.SelectSingleNode("x:v", sheetNs)
                            Dim fNode = cellNode.SelectSingleNode("x:f", sheetNs)

                            Dim displayValue As String = ""
                            Dim formulaStr As String = ""

                            If fNode IsNot Nothing Then
                                formulaStr = fNode.InnerText
                                displayValue = If(vNode?.InnerText, "")
                            ElseIf cellType = "s" Then
                                If vNode IsNot Nothing Then
                                    Dim ssIndex As Integer
                                    If Integer.TryParse(vNode.InnerText, ssIndex) AndAlso
                                       ssIndex >= 0 AndAlso ssIndex < sharedStrings.Count Then
                                        displayValue = sharedStrings(ssIndex)
                                    End If
                                End If
                            ElseIf cellType = "b" Then
                                displayValue = If(vNode?.InnerText = "1", "TRUE", "FALSE")
                            Else
                                displayValue = If(vNode?.InnerText, "")
                            End If

                            If String.IsNullOrEmpty(formulaStr) AndAlso String.IsNullOrEmpty(displayValue) Then Continue For

                            sb.Append(cellRef)
                            sb.Append(vbTab)
                            sb.Append("FORMULA:")
                            If Not String.IsNullOrEmpty(formulaStr) Then
                                sb.Append("=")
                                sb.Append(formulaStr)
                            End If
                            sb.Append(vbTab)
                            sb.Append("VALUE: ")
                            sb.AppendLine(displayValue)
                        Next
                    End If

                    sb.AppendLine()
                Next

                Dim result = sb.ToString().Trim()
                Return If(String.IsNullOrWhiteSpace(result), "Error: No data found in .xlsx.", result)

            Catch ex As IOException
                If Not silent AndAlso IsWorkbookOpenInExcel(xlsxPath) Then
                    If EnsureClosedWorkbookForSandboxedRead(xlsxPath, silent) Then
                        Return ReadXlsxSandboxed(xlsxPath, silent, askWorksheetSelection)
                    End If

                    Return "Error: The workbook is open in Excel."
                End If

                Return $"Error reading .xlsx: {ex.Message}"

            Catch ex As Exception
                Return $"Error reading .xlsx: {ex.Message}"

            Finally
                Try : If Directory.Exists(tempDir) Then Directory.Delete(tempDir, True)
                Catch : End Try
            End Try
        End Function


        ''' <summary>
        ''' Returns <c>True</c> when the workbook cannot currently be opened with exclusive read access,
        ''' which usually means that Excel still has the file open.
        ''' </summary>
        Private Shared Function IsWorkbookOpenInExcel(xlsxPath As String) As Boolean
            Try
                Using fs As New FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.None)
                End Using

                Return False

            Catch ex As IOException
                Return True
            End Try
        End Function

        ''' <summary>
        ''' In interactive mode, prompts the user to close an open workbook and retry.
        ''' </summary>
        Private Shared Function EnsureClosedWorkbookForSandboxedRead(xlsxPath As String,
                                                                     silent As Boolean) As Boolean
            If silent Then
                Return True
            End If

            Do While IsWorkbookOpenInExcel(xlsxPath)
                Dim answer As Integer = ShowCustomYesNoBox(
                    "The Excel workbook '" & Path.GetFileName(xlsxPath) & "' appears to be open. " &
                    "Please close it, then click Retry to try reading it again.",
                    "Retry",
                    "Cancel",
                    AN & " - Workbook Open",
                    nonModal:=True)

                If answer <> 1 Then
                    Return False
                End If

                System.Threading.Thread.Sleep(250)
            Loop

            Return True
        End Function


        ''' <summary>
        ''' Resolves the full path to a sheet XML file from a relationship ID or positional fallback.
        ''' </summary>
        Private Shared Function ResolveSheetPath(tempDir As String, rId As String,
                                                  ridMap As Dictionary(Of String, String),
                                                  positionalIndex As Integer) As String
            If Not String.IsNullOrWhiteSpace(rId) AndAlso ridMap.ContainsKey(rId) Then
                Dim rel = ridMap(rId)
                Dim resolved = Path.Combine(tempDir, "xl", rel.Replace("/"c, Path.DirectorySeparatorChar))
                If File.Exists(resolved) Then Return resolved
            End If
            ' Positional fallback
            Dim fallback = Path.Combine(tempDir, "xl", "worksheets", $"sheet{positionalIndex}.xml")
            Return If(File.Exists(fallback), fallback, Nothing)
        End Function


        ' ═══════════════════════════════════════════════════════════════════════
        '  PPTX — Sandboxed
        ' ═══════════════════════════════════════════════════════════════════════

        Private Const SB_PptxNs As String = "http://schemas.openxmlformats.org/presentationml/2006/main"
        Private Const SB_DrawNs As String = "http://schemas.openxmlformats.org/drawingml/2006/main"
        Private Const SB_PptxRelNs As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        ''' <summary>
        ''' Extracts plain text from a .pptx file without COM interop.
        ''' Output format matches <c>ExtractPowerPointText</c>: <c>=== Slide {n} ===</c> + shape text +
        ''' <c>--- Notes ---</c> per slide.
        ''' </summary>
        ''' <param name="pptxPath">Absolute path to the .pptx file.</param>
        ''' <returns>Extracted text content, or an error string on failure.</returns>
        Public Shared Function ReadPptxSandboxed(pptxPath As String) As String
            If String.IsNullOrWhiteSpace(pptxPath) OrElse Not File.Exists(pptxPath) Then
                Return "Error: File not found."
            End If

            Dim tempDir As String = Path.Combine(Path.GetTempPath(), "ri_pptx_" & Guid.NewGuid().ToString("N"))
            Try
                ZipFile.ExtractToDirectory(pptxPath, tempDir)

                Dim presPath = Path.Combine(tempDir, "ppt", "presentation.xml")
                If Not File.Exists(presPath) Then Return "Error: Not a valid .pptx file (missing ppt/presentation.xml)."

                Dim presDoc As New XmlDocument()
                presDoc.Load(presPath)
                Dim presNs As New XmlNamespaceManager(presDoc.NameTable)
                presNs.AddNamespace("p", SB_PptxNs)
                presNs.AddNamespace("r", SB_PptxRelNs)

                ' Map rId → file path
                Dim relsPath = Path.Combine(tempDir, "ppt", "_rels", "presentation.xml.rels")
                Dim ridMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                If File.Exists(relsPath) Then
                    Dim relsDoc As New XmlDocument()
                    relsDoc.Load(relsPath)
                    If relsDoc.DocumentElement IsNot Nothing Then
                        For Each relNode As XmlNode In relsDoc.DocumentElement.ChildNodes
                            If relNode.Attributes Is Nothing Then Continue For
                            Dim id = relNode.Attributes("Id")
                            Dim target = relNode.Attributes("Target")
                            If id IsNot Nothing AndAlso target IsNot Nothing Then
                                ridMap(id.Value) = target.Value
                            End If
                        Next
                    End If
                End If

                Dim slideIdNodes = presDoc.SelectNodes("//p:sldIdLst/p:sldId", presNs)
                Dim sb As New StringBuilder(2048)
                Dim slideIndex As Integer = 0

                If slideIdNodes IsNot Nothing Then
                    For Each slideIdNode As XmlNode In slideIdNodes
                        slideIndex += 1
                        Dim rId = slideIdNode.Attributes("r:id")?.Value

                        Dim slidePath = ResolveSlidePath(tempDir, rId, ridMap, slideIndex)
                        If slidePath Is Nothing OrElse Not File.Exists(slidePath) Then Continue For

                        sb.AppendLine("=== Slide " & slideIndex.ToString(Globalization.CultureInfo.InvariantCulture) & " ===")

                        ' Extract shape text via DrawingML <a:t> nodes, grouped by <a:p>
                        ExtractDrawingText(slidePath, sb)

                        ' Extract notes
                        ExtractSlideNotes(slidePath, sb)

                        sb.AppendLine()
                    Next
                End If

                Dim result = sb.ToString().Trim()
                Return If(String.IsNullOrWhiteSpace(result), "Error: No text content found in .pptx.", result)

            Catch ex As Exception
                Return $"Error reading .pptx: {ex.Message}"
            Finally
                Try : If Directory.Exists(tempDir) Then Directory.Delete(tempDir, True)
                Catch : End Try
            End Try
        End Function

        Private Shared Function ResolveSlidePath(tempDir As String, rId As String,
                                                  ridMap As Dictionary(Of String, String),
                                                  positionalIndex As Integer) As String
            If Not String.IsNullOrWhiteSpace(rId) AndAlso ridMap.ContainsKey(rId) Then
                Dim rel = ridMap(rId)
                Dim resolved = Path.Combine(tempDir, "ppt", rel.Replace("/"c, Path.DirectorySeparatorChar))
                If File.Exists(resolved) Then Return resolved
            End If
            Dim fallback = Path.Combine(tempDir, "ppt", "slides", $"slide{positionalIndex}.xml")
            Return If(File.Exists(fallback), fallback, Nothing)
        End Function

        ''' <summary>
        ''' Extracts text from DrawingML shapes in a slide XML, one line per text shape.
        ''' Matches <c>ExtractPowerPointText</c> output: each shape's text on its own line.
        ''' </summary>
        Private Shared Sub ExtractDrawingText(slideXmlPath As String, sb As StringBuilder)
            Dim slideDoc As New XmlDocument()
            slideDoc.Load(slideXmlPath)
            Dim slideNs As New XmlNamespaceManager(slideDoc.NameTable)
            slideNs.AddNamespace("a", SB_DrawNs)
            slideNs.AddNamespace("p", SB_PptxNs)

            ' Group by <p:sp> or <p:txBody> shapes — each shape gets one output line
            Dim spNodes = slideDoc.SelectNodes("//p:sp", slideNs)
            If spNodes IsNot Nothing Then
                For Each sp As XmlNode In spNodes
                    Dim txBody = sp.SelectSingleNode(".//p:txBody", slideNs)
                    If txBody Is Nothing Then Continue For

                    Dim shapeSb As New StringBuilder()
                    Dim paragraphs = txBody.SelectNodes("a:p", slideNs)
                    If paragraphs IsNot Nothing Then
                        Dim first As Boolean = True
                        For Each para As XmlNode In paragraphs
                            Dim tNodes = para.SelectNodes(".//a:t", slideNs)
                            If tNodes Is Nothing OrElse tNodes.Count = 0 Then Continue For
                            Dim paraText As New StringBuilder()
                            For Each tNode As XmlNode In tNodes
                                paraText.Append(tNode.InnerText)
                            Next
                            Dim text = paraText.ToString()
                            If Not String.IsNullOrWhiteSpace(text) Then
                                If Not first Then shapeSb.AppendLine()
                                shapeSb.Append(text)
                                first = False
                            End If
                        Next
                    End If

                    Dim shapeText = shapeSb.ToString().Trim()
                    If Not String.IsNullOrWhiteSpace(shapeText) Then
                        sb.AppendLine(shapeText)
                    End If
                Next
            End If
        End Sub

        ''' <summary>
        ''' Extracts notes from a slide's associated notesSlide, if present.
        ''' Matches <c>ExtractPowerPointText</c> output: <c>--- Notes ---</c> header.
        ''' </summary>
        Private Shared Sub ExtractSlideNotes(slideXmlPath As String, sb As StringBuilder)
            Dim notesRelsPath = Path.Combine(Path.GetDirectoryName(slideXmlPath), "_rels",
                                             Path.GetFileName(slideXmlPath) & ".rels")
            If Not File.Exists(notesRelsPath) Then Return

            Try
                Dim relsDoc As New XmlDocument()
                relsDoc.Load(notesRelsPath)
                If relsDoc.DocumentElement Is Nothing Then Return

                For Each rel As XmlNode In relsDoc.DocumentElement.ChildNodes
                    If rel.Attributes Is Nothing Then Continue For
                    Dim relType = rel.Attributes("Type")?.Value
                    If relType Is Nothing OrElse Not relType.Contains("notesSlide") Then Continue For

                    Dim notesTarget = rel.Attributes("Target")?.Value
                    If String.IsNullOrWhiteSpace(notesTarget) Then Continue For

                    Dim notesPath = Path.Combine(Path.GetDirectoryName(slideXmlPath),
                                                 notesTarget.Replace("/"c, Path.DirectorySeparatorChar))
                    If Not File.Exists(notesPath) Then Continue For

                    Dim notesDoc As New XmlDocument()
                    notesDoc.Load(notesPath)
                    Dim notesNs As New XmlNamespaceManager(notesDoc.NameTable)
                    notesNs.AddNamespace("a", SB_DrawNs)

                    Dim noteTexts = notesDoc.SelectNodes("//a:t", notesNs)
                    If noteTexts Is Nothing OrElse noteTexts.Count = 0 Then Continue For

                    Dim hasNotes = False
                    For Each nt As XmlNode In noteTexts
                        If Not String.IsNullOrWhiteSpace(nt.InnerText) Then
                            If Not hasNotes Then
                                sb.AppendLine("--- Notes ---")
                                hasNotes = True
                            End If
                            sb.AppendLine(nt.InnerText.Trim())
                        End If
                    Next
                Next
            Catch
            End Try
        End Sub


        ' ═══════════════════════════════════════════════════════════════════════
        '  EML — Sandboxed
        ' ═══════════════════════════════════════════════════════════════════════

        ''' <summary>
        ''' Extracts plain text from a .eml (RFC 2822) file by parsing headers and body.
        ''' Output matches <c>ParseEmlAsText</c> format in AutoPilot.
        ''' Attachments are extracted inline when they can be decoded safely.
        ''' </summary>
        Public Shared Function ReadEmlSandboxed(emlPath As String) As String
            Return ReadEmlSandboxedInternal(emlPath, 0)
        End Function


        Private Shared Function NormalizeHeaderBlockLineEndings(value As String) As String
            If value Is Nothing Then Return ""
            Return value.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf)
        End Function

        Private Shared Function CleanFriendlyHeaderValue(value As String) As String
            value = If(value, "")
            If value = "" Then Return ""

            value = DecodeMimeEncodedWords(value)

            value = value.Replace(ChrW(0), "")
            value = value.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            value = System.Text.RegularExpressions.Regex.Replace(value, "\n[ \t]+", " ")
            value = System.Text.RegularExpressions.Regex.Replace(value, "\s+", " ")
            value = value.Trim()

            If value = "" Then Return ""
            If value.IndexOf("__substg1.0_", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return ""
            If value.IndexOf("__recip_version1.0", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return ""
            If IsLikelyMsgNoiseLine(value) Then Return ""

            Return value
        End Function

        Private Shared Function DecodeMimeEncodedWords(value As String) As String
            If System.String.IsNullOrWhiteSpace(value) Then Return ""

            Try
                Return System.Text.RegularExpressions.Regex.Replace(
            value,
            "=\?([^?]+)\?([bBqQ])\?([^?]*)\?=",
            Function(m As System.Text.RegularExpressions.Match) As String
                Try
                    Dim charsetName As String = m.Groups(1).Value.Trim()
                    Dim mode As String = m.Groups(2).Value.ToUpperInvariant()
                    Dim encodedText As String = m.Groups(3).Value

                    Dim data As Byte()

                    If mode = "B" Then
                        data = System.Convert.FromBase64String(encodedText)
                    Else
                        encodedText = encodedText.Replace("_"c, " "c)
                        data = DecodeQuotedPrintableHeaderBytes(encodedText)
                    End If

                    Dim enc As System.Text.Encoding = GetEncodingByNameSafe(charsetName)
                    Return enc.GetString(data)
                Catch
                    Return m.Value
                End Try
            End Function,
            System.Text.RegularExpressions.RegexOptions.Singleline)
            Catch
                Return value
            End Try
        End Function

        Private Shared Function DecodeQuotedPrintableHeaderBytes(input As String) As Byte()
            If input Is Nothing Then Return System.Array.Empty(Of Byte)()

            Using ms As New System.IO.MemoryStream()
                Dim i As Integer = 0

                While i < input.Length
                    If input(i) = "="c AndAlso i + 2 < input.Length Then
                        Dim hexText As String = input.Substring(i + 1, 2)
                        Dim b As Byte

                        If System.Byte.TryParse(
                    hexText,
                    System.Globalization.NumberStyles.HexNumber,
                    System.Globalization.CultureInfo.InvariantCulture,
                    b
                ) Then
                            ms.WriteByte(b)
                            i += 3
                            Continue While
                        End If
                    End If

                    ms.WriteByte(System.Convert.ToByte(AscW(input(i)) And &HFF))
                    i += 1
                End While

                Return ms.ToArray()
            End Using
        End Function

        Private Shared Function ExtractMimeHeaderValue(headerSection As String, headerName As String) As String
            If String.IsNullOrWhiteSpace(headerSection) OrElse String.IsNullOrWhiteSpace(headerName) Then Return ""

            Try
                Dim normalizedHeaders As String =
                    System.Text.RegularExpressions.Regex.Replace(
                        headerSection,
                        "(\r?\n)[ \t]+",
                        " ",
                        System.Text.RegularExpressions.RegexOptions.Singleline)

                Dim pattern As String =
                    "(?im)^" &
                    System.Text.RegularExpressions.Regex.Escape(headerName) &
                    ":\s*(.+?)(?=\r?\n[!-9;-~]+:|\z)"

                Dim matchResult =
                    System.Text.RegularExpressions.Regex.Match(
                        normalizedHeaders,
                        pattern,
                        System.Text.RegularExpressions.RegexOptions.Singleline Or
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase)

                If Not matchResult.Success Then
                    Return ""
                End If

                Return CleanFriendlyHeaderValue(matchResult.Groups(1).Value)
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function ExtractMimeHeaderParameter(headerSection As String,
                                                           headerName As String,
                                                           parameterName As String) As String
            If String.IsNullOrWhiteSpace(headerSection) OrElse
               String.IsNullOrWhiteSpace(headerName) OrElse
               String.IsNullOrWhiteSpace(parameterName) Then
                Return ""
            End If

            Try
                Dim headerValue As String = ExtractMimeHeaderValue(headerSection, headerName)
                If headerValue = "" Then Return ""

                Dim pattern As String =
                    "(?i)(?:^|;)\s*" &
                    System.Text.RegularExpressions.Regex.Escape(parameterName) &
                    "\s*=\s*(""([^""]*)""|([^;]*))"

                Dim matchResult =
                    System.Text.RegularExpressions.Regex.Match(
                        headerValue,
                        pattern,
                        System.Text.RegularExpressions.RegexOptions.Singleline Or
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase)

                If Not matchResult.Success Then
                    Return ""
                End If

                Dim value As String = matchResult.Groups(2).Value
                If value = "" Then
                    value = matchResult.Groups(3).Value
                End If

                Return CleanFriendlyHeaderValue(value.Trim())
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function GetEncodingByNameSafe(charsetName As String) As Encoding
            If String.IsNullOrWhiteSpace(charsetName) Then
                Return Encoding.UTF8
            End If

            Try
                Return Encoding.GetEncoding(charsetName.Trim())
            Catch
                Return Encoding.UTF8
            End Try
        End Function

        Private Shared Function DecodeQuotedPrintableBytes(input As String) As Byte()
            If input Is Nothing Then Return Array.Empty(Of Byte)()

            Dim text As String = input.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            text = text.Replace("=" & vbLf, "")

            Using ms As New MemoryStream()
                Dim i As Integer = 0

                While i < text.Length
                    Dim ch As Char = text(i)

                    If ch = "="c AndAlso i + 2 < text.Length Then
                        Dim hex = text.Substring(i + 1, 2)
                        Dim b As Byte

                        If Byte.TryParse(hex, Globalization.NumberStyles.HexNumber,
                                         Globalization.CultureInfo.InvariantCulture, b) Then
                            ms.WriteByte(b)
                            i += 3
                            Continue While
                        End If
                    End If

                    Dim bytes = Encoding.ASCII.GetBytes(New Char() {ch})
                    ms.Write(bytes, 0, bytes.Length)
                    i += 1
                End While

                Return ms.ToArray()
            End Using
        End Function

        Private Shared Function DecodeMimePartBytes(partBody As String,
                                                    transferEncoding As String) As Byte()
            transferEncoding = If(transferEncoding, "").Trim().ToLowerInvariant()
            partBody = If(partBody, "")

            Select Case transferEncoding
                Case "base64"
                    Dim compact As String =
                        System.Text.RegularExpressions.Regex.Replace(partBody, "\s+", "")
                    Try
                        Return System.Convert.FromBase64String(compact)
                    Catch
                        Return Array.Empty(Of Byte)()
                    End Try

                Case "quoted-printable"
                    Return DecodeQuotedPrintableBytes(partBody)

                Case Else
                    Return Encoding.UTF8.GetBytes(partBody)
            End Select
        End Function

        Private Shared Function DecodeMimePartText(partBody As String,
                                                   transferEncoding As String,
                                                   charsetName As String) As String
            Dim data As Byte() = DecodeMimePartBytes(partBody, transferEncoding)
            If data Is Nothing OrElse data.Length = 0 Then Return ""

            Try
                Dim enc = GetEncodingByNameSafe(charsetName)
                Dim text As String = enc.GetString(data)
                text = text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf)
                Return text.Trim()
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function SanitizeAttachmentFileName(fileName As String,
                                                           fallbackName As String) As String
            Dim result As String = If(fileName, "").Trim()
            If result = "" Then
                result = fallbackName
            End If

            For Each ch As Char In Path.GetInvalidFileNameChars()
                result = result.Replace(ch, "_"c)
            Next

            If result = "" Then
                result = fallbackName
            End If

            Return result
        End Function

        Private Shared Function ExtractInlineAttachmentTextFromSavedFile(filePath As String,
                                                                        depth As Integer) As String
            If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
                Return "[Skipped: attachment file could not be created for extraction]"
            End If

            If depth > 5 Then
                Return $"[Skipped: max nesting depth reached for {Path.GetFileName(filePath)}]"
            End If

            Dim ext As String = Path.GetExtension(filePath).ToLowerInvariant()

            Try
                Select Case ext
                    Case ".txt", ".csv", ".tsv", ".log", ".json", ".xml", ".html", ".htm",
                         ".md", ".yaml", ".yml", ".ini", ".ics", ".vcf",
                         ".vb", ".cs", ".js", ".ts", ".py", ".java", ".cpp", ".c", ".h", ".sql"
                        Dim text = ReadTextFile(filePath, False).Trim()
                        If text <> "" Then Return text
                        Return $"[Skipped: no readable text extracted from '{Path.GetFileName(filePath)}']"

                    Case ".rtf"
                        Dim text = ReadRtfAsText(filePath, False).Trim()
                        If text <> "" Then Return text
                        Return $"[Skipped: no readable text extracted from '{Path.GetFileName(filePath)}']"

                    Case ".doc"
                        If INI_AllowLegacyDocFiles_Cached Then
                            Dim text = ReadWordDocument(filePath, False).Trim()
                            If text <> "" Then Return text
                            Return $"[Skipped: no readable text extracted from '{Path.GetFileName(filePath)}']"
                        End If
                        Return "[Skipped: .doc format disabled for security]"

                    Case ".docx", ".docm"
                        Dim text = ReadDocxSandboxed(filePath).Trim()
                        If text <> "" AndAlso Not text.StartsWith("Error", StringComparison.OrdinalIgnoreCase) Then Return text
                        Return $"[Skipped: no readable text extracted from '{Path.GetFileName(filePath)}']"

                    Case ".xlsx", ".xlsm"
                        Dim text = ReadXlsxSandboxed(filePath, silent:=True, askWorksheetSelection:=False).Trim()
                        If text <> "" AndAlso Not text.StartsWith("Error", StringComparison.OrdinalIgnoreCase) Then Return text
                        Return $"[Skipped: no readable text extracted from '{Path.GetFileName(filePath)}']"

                    Case ".pptx", ".pptm"
                        Dim text = ReadPptxSandboxed(filePath).Trim()
                        If text <> "" AndAlso Not text.StartsWith("Error", StringComparison.OrdinalIgnoreCase) Then Return text
                        Return $"[Skipped: no readable text extracted from '{Path.GetFileName(filePath)}']"

                    Case ".pdf"
                        Try
                            Dim text = ReadPdfAsText(filePath, True, False, False, Nothing).Result.Trim()
                            If text <> "" AndAlso Not text.StartsWith("Error", StringComparison.OrdinalIgnoreCase) Then Return text
                            Return $"[Skipped: no readable text extracted from '{Path.GetFileName(filePath)}']"
                        Catch ex As Exception
                            Return $"[Skipped: PDF extraction failed for '{Path.GetFileName(filePath)}': {ex.Message}]"
                        End Try

                    Case ".eml"
                        Return ReadEmlSandboxedInternal(filePath, depth + 1).Trim()

                    Case ".msg"
                        Return ReadMsgSandboxed(filePath, Nothing, depth + 1).Trim()

                    Case Else
                        Return $"[Skipped: unsupported attachment type '{ext}' for '{Path.GetFileName(filePath)}']"
                End Select
            Catch ex As Exception
                Return $"[Skipped: failed to extract '{Path.GetFileName(filePath)}': {ex.Message}]"
            End Try
        End Function

        Private Shared Function BuildInlineAttachmentSection(fileName As String,
                                                             attachmentText As String) As String
            If String.IsNullOrWhiteSpace(attachmentText) Then Return ""

            Dim sb As New StringBuilder()
            sb.AppendLine($"═══ Attachment: {fileName} ═══")
            sb.AppendLine()
            sb.Append(attachmentText.Trim())
            Return sb.ToString().TrimEnd()
        End Function

        Private Shared Function ReadEmlInlineAttachmentSections(headerSection As String,
                                                                bodySection As String,
                                                                sourceFileName As String,
                                                                depth As Integer) As List(Of String)
            Dim result As New List(Of String)()

            If String.IsNullOrWhiteSpace(headerSection) OrElse String.IsNullOrWhiteSpace(bodySection) Then
                Return result
            End If

            Dim boundary As String =
                ExtractMimeHeaderParameter(headerSection, "Content-Type", "boundary")

            If String.IsNullOrWhiteSpace(boundary) Then
                Return result
            End If

            Dim tempDir As String = Path.Combine(Path.GetTempPath(), "ri_eml_" & Guid.NewGuid().ToString("N"))

            Try
                Directory.CreateDirectory(tempDir)

                Dim parts = bodySection.Split({$"--{boundary}"}, StringSplitOptions.RemoveEmptyEntries)
                Dim attachmentIndex As Integer = 0

                For Each part In parts
                    If part.StartsWith("--", StringComparison.Ordinal) Then Continue For

                    Dim partHeaderEnd = part.IndexOf(vbCrLf & vbCrLf, StringComparison.Ordinal)
                    If partHeaderEnd < 0 Then partHeaderEnd = part.IndexOf(vbLf & vbLf, StringComparison.Ordinal)
                    If partHeaderEnd < 0 Then Continue For

                    Dim partHeaders As String = part.Substring(0, partHeaderEnd)
                    Dim partBody As String = part.Substring(partHeaderEnd).TrimStart({CChar(vbCr), CChar(vbLf)})

                    Dim fileName As String = ExtractMimeHeaderParameter(partHeaders, "Content-Disposition", "filename")
                    If String.IsNullOrWhiteSpace(fileName) Then
                        fileName = ExtractMimeHeaderParameter(partHeaders, "Content-Type", "name")
                    End If

                    If String.IsNullOrWhiteSpace(fileName) Then
                        Continue For
                    End If

                    attachmentIndex += 1

                    Dim transferEncoding As String = ExtractMimeHeaderValue(partHeaders, "Content-Transfer-Encoding")
                    Dim data As Byte() = DecodeMimePartBytes(partBody, transferEncoding)

                    fileName = SanitizeAttachmentFileName(
                        fileName,
                        Path.GetFileNameWithoutExtension(sourceFileName) & $"_attachment_{attachmentIndex:000}.bin")

                    Dim attachmentText As String = ""

                    If data Is Nothing OrElse data.Length = 0 Then
                        attachmentText = $"[Skipped: attachment '{fileName}' could not be decoded from the .eml part]"
                    Else
                        Dim attachmentPath As String = Path.Combine(tempDir, fileName)
                        File.WriteAllBytes(attachmentPath, data)

                        attachmentText =
                            ExtractInlineAttachmentTextFromSavedFile(attachmentPath, depth + 1)
                    End If

                    Dim section As String = BuildInlineAttachmentSection(fileName, attachmentText)
                    If section <> "" Then
                        result.Add(section)
                    End If
                Next

            Catch
            Finally
                Try
                    If Directory.Exists(tempDir) Then
                        Directory.Delete(tempDir, recursive:=True)
                    End If
                Catch
                End Try
            End Try

            Return result
        End Function

        Private Shared Function ReadEmlSandboxedInternal(emlPath As String, depth As Integer) As String
            If String.IsNullOrWhiteSpace(emlPath) OrElse Not File.Exists(emlPath) Then
                Return "Error: File not found."
            End If

            If depth > 5 Then
                Return $"[Skipped: max nesting depth reached for {Path.GetFileName(emlPath)}]"
            End If

            Try
                Dim emlContent As String = ReadMailTextFileBestEffort(emlPath)
                If String.IsNullOrWhiteSpace(emlContent) Then Return "Error: Empty .eml file."

                Dim headerEnd = emlContent.IndexOf(vbCrLf & vbCrLf, StringComparison.Ordinal)
                If headerEnd < 0 Then headerEnd = emlContent.IndexOf(vbLf & vbLf, StringComparison.Ordinal)

                Dim headerSection As String
                Dim bodySection As String

                If headerEnd >= 0 Then
                    headerSection = emlContent.Substring(0, headerEnd)
                    bodySection = emlContent.Substring(headerEnd).TrimStart({CChar(vbCr), CChar(vbLf)})
                Else
                    headerSection = emlContent
                    bodySection = ""
                End If

                Dim fromText As String = ExtractMimeHeaderValue(headerSection, "From")
                Dim toText As String = ExtractMimeHeaderValue(headerSection, "To")
                Dim ccText As String = ExtractMimeHeaderValue(headerSection, "CC")
                Dim subjectText As String = ExtractMimeHeaderValue(headerSection, "Subject")
                Dim dateText As String = ExtractMimeHeaderValue(headerSection, "Date")
                Dim charsetName As String = ExtractMimeHeaderParameter(headerSection, "Content-Type", "charset")
                Dim boundary As String = ExtractMimeHeaderParameter(headerSection, "Content-Type", "boundary")

                Dim bodyText As String = ""
                If Not String.IsNullOrWhiteSpace(bodySection) Then
                    If Not String.IsNullOrWhiteSpace(boundary) Then
                        Dim parts = bodySection.Split({$"--{boundary}"}, StringSplitOptions.RemoveEmptyEntries)

                        For Each part In parts
                            If part.StartsWith("--", StringComparison.Ordinal) Then Continue For

                            Dim partHeaderEnd = part.IndexOf(vbCrLf & vbCrLf, StringComparison.Ordinal)
                            If partHeaderEnd < 0 Then partHeaderEnd = part.IndexOf(vbLf & vbLf, StringComparison.Ordinal)
                            If partHeaderEnd < 0 Then Continue For

                            Dim partHeaders As String = part.Substring(0, partHeaderEnd)
                            Dim partBody As String = part.Substring(partHeaderEnd).TrimStart({CChar(vbCr), CChar(vbLf)})

                            Dim partFileName As String =
                                ExtractMimeHeaderParameter(partHeaders, "Content-Disposition", "filename")
                            If partFileName = "" Then
                                partFileName = ExtractMimeHeaderParameter(partHeaders, "Content-Type", "name")
                            End If

                            If partFileName <> "" Then
                                Continue For
                            End If

                            Dim contentType As String = ExtractMimeHeaderValue(partHeaders, "Content-Type")
                            Dim transferEncoding As String = ExtractMimeHeaderValue(partHeaders, "Content-Transfer-Encoding")
                            Dim partCharset As String = ExtractMimeHeaderParameter(partHeaders, "Content-Type", "charset")

                            If contentType.IndexOf("text/plain", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                bodyText = DecodeMimePartText(partBody, transferEncoding, partCharset)
                                If bodyText <> "" Then Exit For
                            ElseIf contentType.IndexOf("text/html", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                bodyText = RemoveHTML(DecodeMimePartText(partBody, transferEncoding, partCharset)).Trim()
                                If bodyText <> "" Then Exit For
                            End If
                        Next
                    End If

                    If bodyText = "" Then
                        bodyText = DecodeMimePartText(bodySection, "", charsetName)
                        If bodyText = "" Then
                            bodyText = bodySection.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf).Trim()
                        End If
                    End If
                End If

                Dim sb As New StringBuilder()
                sb.AppendLine("═══════════════════════════════════════════════════")
                sb.AppendLine($"EMAIL MESSAGE (from {Path.GetFileName(emlPath)})")
                sb.AppendLine("═══════════════════════════════════════════════════")
                sb.AppendLine()

                If fromText <> "" Then sb.AppendLine("From: " & fromText)
                If toText <> "" Then sb.AppendLine("To: " & toText)
                If ccText <> "" Then sb.AppendLine("CC: " & ccText)
                If subjectText <> "" Then sb.AppendLine("Subject: " & subjectText)
                If dateText <> "" Then sb.AppendLine("Date: " & dateText)

                sb.AppendLine()
                sb.AppendLine("───────────────────────────────────────────────────")
                sb.AppendLine()

                If bodyText <> "" Then
                    sb.Append(bodyText.Trim())
                Else
                    sb.Append("[No readable body text found in the .eml file]")
                End If

                Dim attachmentSections =
                    ReadEmlInlineAttachmentSections(headerSection, bodySection, Path.GetFileName(emlPath), depth)

                For Each section In attachmentSections
                    If section <> "" Then
                        sb.AppendLine()
                        sb.AppendLine()
                        sb.Append(section)
                    End If
                Next

                Return sb.ToString().Trim()
            Catch ex As Exception
                Return $"Error reading .eml: {ex.Message}"
            End Try
        End Function

        Private Shared Function ReadMailTextFileBestEffort(filePath As String) As String
            Dim data As Byte() = System.IO.File.ReadAllBytes(filePath)

            If data Is Nothing OrElse data.Length = 0 Then Return ""

            Try
                If data.Length >= 3 AndAlso
           data(0) = &HEF AndAlso data(1) = &HBB AndAlso data(2) = &HBF Then
                    Return System.Text.Encoding.UTF8.GetString(data)
                End If

                If data.Length >= 2 AndAlso
           data(0) = &HFF AndAlso data(1) = &HFE Then
                    Return System.Text.Encoding.Unicode.GetString(data)
                End If

                If data.Length >= 2 AndAlso
           data(0) = &HFE AndAlso data(1) = &HFF Then
                    Return System.Text.Encoding.BigEndianUnicode.GetString(data)
                End If
            Catch
            End Try

            Try
                Dim utf8Strict As New System.Text.UTF8Encoding(False, True)
                Return utf8Strict.GetString(data)
            Catch
            End Try

            Try
                Return System.Text.Encoding.GetEncoding(1252).GetString(data)
            Catch
            End Try

            Return System.Text.Encoding.Default.GetString(data)
        End Function

        ' ═══════════════════════════════════════════════════════════════════════
        '  MSG — Sandboxed, dependency-free OLE Compound File + MAPI property reader
        ' ═══════════════════════════════════════════════════════════════════════

        ''' <summary>
        ''' Kept only for source compatibility with existing callers. This implementation does not call Outlook
        ''' and does not use the callback.
        ''' </summary>
        Public Delegate Function MsgReadCallback(msgPath As String, tempDir As String,
                                                  ByRef nestedFiles As System.Collections.Generic.List(Of String)) As String

        ''' <summary>
        ''' Extracts readable text from an Outlook .msg file without Outlook, MsgReader, OpenMcdf,
        ''' or any other external dependency. Reads the .msg file as an OLE Compound File and extracts
        ''' MAPI properties directly.
        ''' </summary>
        Public Shared Function ReadMsgSandboxed(msgPath As String,
                                                Optional msgReadFunc As MsgReadCallback = Nothing,
                                                Optional depth As Integer = 0) As String
            If System.String.IsNullOrWhiteSpace(msgPath) OrElse Not System.IO.File.Exists(msgPath) Then
                Return "Error: File not found."
            End If

            If depth > 5 Then
                Return "[Skipped: max nesting depth reached for " & System.IO.Path.GetFileName(msgPath) & "]"
            End If

            Try
                If Not IsMsgOleCompoundFile(msgPath) Then
                    Return "Error: The file is not a valid Outlook .msg file because it is not an OLE Compound File."
                End If

                Using oleDoc As New MsgOleCompoundDocument(msgPath)
                    Dim resultText As String = ReadMsgFromOleDocument(oleDoc, msgPath, depth)

                    If Not System.String.IsNullOrWhiteSpace(resultText) Then
                        Return resultText.Trim()
                    End If
                End Using

                Return "Error: The .msg file does not contain readable MAPI message properties."

            Catch ex As System.Exception
                Return "Error reading .msg: " & ex.Message
            End Try
        End Function

        Private Shared Function IsMsgOleCompoundFile(filePath As String) As Boolean
            If System.String.IsNullOrWhiteSpace(filePath) OrElse Not System.IO.File.Exists(filePath) Then
                Return False
            End If

            Dim expectedHeader As Byte() = New Byte() {&HD0, &HCF, &H11, &HE0, &HA1, &HB1, &H1A, &HE1}
            Dim actualHeader(7) As Byte

            Try
                Using fs As New System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)
                    If fs.Length < 8 Then Return False
                    Dim readCount As Integer = fs.Read(actualHeader, 0, actualHeader.Length)
                    If readCount <> actualHeader.Length Then Return False
                End Using

                For i As Integer = 0 To expectedHeader.Length - 1
                    If actualHeader(i) <> expectedHeader(i) Then
                        Return False
                    End If
                Next

                Return True
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Private Shared Function ReadMsgFromOleDocument(oleDoc As MsgOleCompoundDocument,
                                                       msgPath As String,
                                                       depth As Integer) As String
            Dim rootEntry As MsgOleDirectoryEntry = oleDoc.RootEntry
            Dim preferredCodePage As Integer = ReadMsgPreferredCodePage(oleDoc, rootEntry)

            Dim headerBlock As String = ReadMsgStringProperty(oleDoc, rootEntry, "007D", preferredCodePage)

            Dim subjectText As String = ReadMsgStringProperty(oleDoc, rootEntry, "0037", preferredCodePage)
            If System.String.IsNullOrWhiteSpace(subjectText) Then
                subjectText = ExtractHeaderValueFromMsgHeaderBlock(headerBlock, "Subject")
            End If
            subjectText = CleanFriendlyHeaderValue(subjectText)

            Dim senderName As String = ReadMsgStringProperty(oleDoc, rootEntry, "0C1A", preferredCodePage)
            Dim senderSmtp As String = ReadMsgStringProperty(oleDoc, rootEntry, "5D01", preferredCodePage)
            Dim senderEmail As String = ReadMsgStringProperty(oleDoc, rootEntry, "0C1F", preferredCodePage)

            If System.String.IsNullOrWhiteSpace(senderSmtp) Then
                senderSmtp = senderEmail
            End If

            Dim fromText As String = BuildFriendlyMsgAddressValue(
                ExtractHeaderValueFromMsgHeaderBlock(headerBlock, "From"),
                senderName,
                senderSmtp)

            Dim toText As String = ReadMsgRecipientList(oleDoc, rootEntry, preferredCodePage, 1)
            If System.String.IsNullOrWhiteSpace(toText) Then
                toText = BuildFriendlyMsgRecipientList(
                    ExtractHeaderValueFromMsgHeaderBlock(headerBlock, "To"),
                    ReadMsgStringProperty(oleDoc, rootEntry, "0E04", preferredCodePage))
            End If

            Dim ccText As String = ReadMsgRecipientList(oleDoc, rootEntry, preferredCodePage, 2)
            If System.String.IsNullOrWhiteSpace(ccText) Then
                ccText = BuildFriendlyMsgRecipientList(
                    ExtractHeaderValueFromMsgHeaderBlock(headerBlock, "CC"),
                    ReadMsgStringProperty(oleDoc, rootEntry, "0E03", preferredCodePage))
            End If

            Dim bccText As String = ReadMsgRecipientList(oleDoc, rootEntry, preferredCodePage, 3)
            If System.String.IsNullOrWhiteSpace(bccText) Then
                bccText = BuildFriendlyMsgRecipientList(
                    ExtractHeaderValueFromMsgHeaderBlock(headerBlock, "BCC"),
                    "")
            End If

            Dim sentText As String = ReadMsgDateProperty(oleDoc, rootEntry, "0039")
            If System.String.IsNullOrWhiteSpace(sentText) Then
                sentText = ReadMsgDateProperty(oleDoc, rootEntry, "0E06")
            End If
            If System.String.IsNullOrWhiteSpace(sentText) Then
                sentText = CleanFriendlyHeaderValue(ExtractHeaderValueFromMsgHeaderBlock(headerBlock, "Date"))
            End If

            Dim bodyText As String = ReadMsgBodyText(oleDoc, rootEntry, preferredCodePage)

            Dim hasAnyMailContent As Boolean =
                Not System.String.IsNullOrWhiteSpace(subjectText) OrElse
                Not System.String.IsNullOrWhiteSpace(fromText) OrElse
                Not System.String.IsNullOrWhiteSpace(toText) OrElse
                Not System.String.IsNullOrWhiteSpace(ccText) OrElse
                Not System.String.IsNullOrWhiteSpace(bodyText)

            If Not hasAnyMailContent Then
                Return ""
            End If

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("═══════════════════════════════════════════════════")
            sb.AppendLine("EMAIL MESSAGE (from " & System.IO.Path.GetFileName(msgPath) & ")")
            sb.AppendLine("═══════════════════════════════════════════════════")
            sb.AppendLine()

            If fromText <> "" Then sb.AppendLine("From: " & fromText)
            If toText <> "" Then sb.AppendLine("To: " & toText)
            If ccText <> "" Then sb.AppendLine("CC: " & ccText)
            If bccText <> "" Then sb.AppendLine("BCC: " & bccText)
            If subjectText <> "" Then sb.AppendLine("Subject: " & subjectText)
            If sentText <> "" Then sb.AppendLine("Date: " & sentText)

            sb.AppendLine()
            sb.AppendLine("───────────────────────────────────────────────────")
            sb.AppendLine()

            If bodyText <> "" Then
                sb.Append(bodyText.Trim())
            Else
                sb.Append("[No readable body text found in the .msg file]")
            End If

            Dim attachmentSections As System.Collections.Generic.List(Of String) =
                ReadMsgInlineAttachmentSections(oleDoc, rootEntry, System.IO.Path.GetFileName(msgPath), depth)

            For Each section As String In attachmentSections
                If section <> "" Then
                    sb.AppendLine()
                    sb.AppendLine()
                    sb.Append(section)
                End If
            Next

            Return sb.ToString().Trim()
        End Function

        Private Shared Function ReadMsgPreferredCodePage(oleDoc As MsgOleCompoundDocument,
                                                         parentEntry As MsgOleDirectoryEntry) As Integer
            Dim codePage As System.Nullable(Of Integer) = ReadMsgIntProperty(oleDoc, parentEntry, "3FFD")
            If codePage.HasValue AndAlso codePage.Value > 0 Then
                Return codePage.Value
            End If

            codePage = ReadMsgIntProperty(oleDoc, parentEntry, "3FDE")
            If codePage.HasValue AndAlso codePage.Value > 0 Then
                Return codePage.Value
            End If

            Return 1252
        End Function

        Private Shared Function ReadMsgStringProperty(oleDoc As MsgOleCompoundDocument,
                                                      parentEntry As MsgOleDirectoryEntry,
                                                      propertyId As String,
                                                      preferredCodePage As Integer) As String
            propertyId = If(propertyId, "").Trim().ToUpperInvariant()
            If propertyId = "" Then Return ""

            Dim unicodeBytes As Byte() = oleDoc.ReadChildStream(parentEntry, "__substg1.0_" & propertyId & "001F")
            If unicodeBytes IsNot Nothing AndAlso unicodeBytes.Length > 0 Then
                Return CleanMsgExtractedText(System.Text.Encoding.Unicode.GetString(RemoveTrailingNullBytes(unicodeBytes, 2)))
            End If

            Dim ansiBytes As Byte() = oleDoc.ReadChildStream(parentEntry, "__substg1.0_" & propertyId & "001E")
            If ansiBytes IsNot Nothing AndAlso ansiBytes.Length > 0 Then
                Dim enc As System.Text.Encoding = GetEncodingByNameSafe(preferredCodePage.ToString(System.Globalization.CultureInfo.InvariantCulture))
                Return CleanMsgExtractedText(enc.GetString(RemoveTrailingNullBytes(ansiBytes, 1)))
            End If

            Return ""
        End Function

        Private Shared Function ReadMsgBinaryProperty(oleDoc As MsgOleCompoundDocument,
                                                      parentEntry As MsgOleDirectoryEntry,
                                                      propertyId As String) As Byte()
            propertyId = If(propertyId, "").Trim().ToUpperInvariant()
            If propertyId = "" Then Return System.Array.Empty(Of Byte)()

            Dim bytes As Byte() = oleDoc.ReadChildStream(parentEntry, "__substg1.0_" & propertyId & "0102")
            If bytes IsNot Nothing AndAlso bytes.Length > 0 Then
                Return bytes
            End If

            bytes = oleDoc.ReadChildStream(parentEntry, "__substg1.0_" & propertyId & "000D")
            If bytes IsNot Nothing AndAlso bytes.Length > 0 Then
                Return bytes
            End If

            Return System.Array.Empty(Of Byte)()
        End Function

        Private Shared Function ReadMsgIntProperty(oleDoc As MsgOleCompoundDocument,
                                                   parentEntry As MsgOleDirectoryEntry,
                                                   propertyId As String) As System.Nullable(Of Integer)
            propertyId = If(propertyId, "").Trim().ToUpperInvariant()
            If propertyId = "" Then Return Nothing

            Dim bytes As Byte() = oleDoc.ReadChildStream(parentEntry, "__substg1.0_" & propertyId & "0003")
            If bytes Is Nothing OrElse bytes.Length < 4 Then
                Return Nothing
            End If

            Return System.BitConverter.ToInt32(bytes, 0)
        End Function

        Private Shared Function ReadMsgDateProperty(oleDoc As MsgOleCompoundDocument,
                                                    parentEntry As MsgOleDirectoryEntry,
                                                    propertyId As String) As String
            propertyId = If(propertyId, "").Trim().ToUpperInvariant()
            If propertyId = "" Then Return ""

            Dim bytes As Byte() = oleDoc.ReadChildStream(parentEntry, "__substg1.0_" & propertyId & "0040")
            If bytes Is Nothing OrElse bytes.Length < 8 Then Return ""

            Try
                Dim fileTime As Long = System.BitConverter.ToInt64(bytes, 0)
                If fileTime <= 0 Then Return ""

                Dim dt As System.DateTime = System.DateTime.FromFileTimeUtc(fileTime).ToLocalTime()
                Return dt.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)
            Catch ex As System.Exception
                Return ""
            End Try
        End Function

        Private Shared Function ReadMsgRecipientList(oleDoc As MsgOleCompoundDocument,
                                                     rootEntry As MsgOleDirectoryEntry,
                                                     preferredCodePage As Integer,
                                                     wantedRecipientType As Integer) As String
            Dim result As New System.Collections.Generic.List(Of String)()
            Dim recipientStorages As System.Collections.Generic.List(Of MsgOleDirectoryEntry) =
                oleDoc.GetChildStorages(rootEntry, "__recip_version1.0_#")

            For Each recipStorage As MsgOleDirectoryEntry In recipientStorages
                Dim recipientType As System.Nullable(Of Integer) = ReadMsgIntProperty(oleDoc, recipStorage, "0C15")

                If recipientType.HasValue AndAlso recipientType.Value <> wantedRecipientType Then
                    Continue For
                End If

                If Not recipientType.HasValue AndAlso wantedRecipientType <> 1 Then
                    Continue For
                End If

                Dim displayName As String = ReadMsgStringProperty(oleDoc, recipStorage, "3001", preferredCodePage)
                Dim smtpAddress As String = ReadMsgStringProperty(oleDoc, recipStorage, "39FE", preferredCodePage)
                Dim emailAddress As String = ReadMsgStringProperty(oleDoc, recipStorage, "3003", preferredCodePage)

                displayName = CleanFriendlyHeaderValue(displayName)
                smtpAddress = ExtractFirstEmailAddress(smtpAddress)
                emailAddress = CleanFriendlyHeaderValue(emailAddress)

                If smtpAddress = "" Then
                    smtpAddress = ExtractFirstEmailAddress(emailAddress)
                End If

                If smtpAddress = "" Then
                    smtpAddress = ExtractFirstEmailAddress(displayName)
                End If

                If displayName <> "" AndAlso smtpAddress <> "" Then
                    displayName = displayName.Replace(smtpAddress, "").Trim()
                    displayName = displayName.Trim("<"c, ">"c, ";"c, ","c, " "c)
                End If

                Dim itemText As String = ""

                If displayName <> "" AndAlso smtpAddress <> "" Then
                    itemText = displayName & " <" & smtpAddress & ">"
                ElseIf smtpAddress <> "" Then
                    itemText = smtpAddress
                ElseIf displayName <> "" Then
                    itemText = displayName
                End If

                If itemText <> "" AndAlso Not ContainsStringOrdinalIgnoreCase(result, itemText) Then
                    result.Add(itemText)
                End If
            Next

            Return System.String.Join("; ", result)
        End Function

        Private Shared Function ReadMsgBodyText(oleDoc As MsgOleCompoundDocument,
                                                rootEntry As MsgOleDirectoryEntry,
                                                preferredCodePage As Integer) As String
            Dim plainBody As String = ReadMsgStringProperty(oleDoc, rootEntry, "1000", preferredCodePage)
            plainBody = CleanMsgExtractedText(plainBody)

            If IsUsableMsgText(plainBody) Then
                Return plainBody
            End If

            Dim htmlBody As String = ReadMsgStringProperty(oleDoc, rootEntry, "1013", preferredCodePage)
            htmlBody = CleanMsgExtractedText(htmlBody)

            If IsUsableMsgText(htmlBody) Then
                Dim htmlAsText As String = CleanMsgExtractedText(RemoveHTML(htmlBody))
                If IsUsableMsgText(htmlAsText) Then
                    Return htmlAsText
                End If
            End If

            Dim htmlBytes As Byte() = ReadMsgBinaryProperty(oleDoc, rootEntry, "1013")
            If htmlBytes IsNot Nothing AndAlso htmlBytes.Length > 0 Then
                htmlBody = DecodeMsgTextBytes(htmlBytes, preferredCodePage)
                htmlBody = CleanMsgExtractedText(htmlBody)

                If IsUsableMsgText(htmlBody) Then
                    Dim htmlAsText As String = CleanMsgExtractedText(RemoveHTML(htmlBody))
                    If IsUsableMsgText(htmlAsText) Then
                        Return htmlAsText
                    End If
                End If
            End If

            Return ""
        End Function

        Private Shared Function ReadMsgInlineAttachmentSections(oleDoc As MsgOleCompoundDocument,
                                                                rootEntry As MsgOleDirectoryEntry,
                                                                sourceFileName As String,
                                                                depth As Integer) As System.Collections.Generic.List(Of String)
            Dim result As New System.Collections.Generic.List(Of String)()
            If depth > 5 Then Return result

            Dim tempDir As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ri_msg_att_" & System.Guid.NewGuid().ToString("N"))

            Try
                System.IO.Directory.CreateDirectory(tempDir)

                Dim attachmentStorages As System.Collections.Generic.List(Of MsgOleDirectoryEntry) =
                    oleDoc.GetChildStorages(rootEntry, "__attach_version1.0_#")

                Dim attachmentIndex As Integer = 0

                For Each attachStorage As MsgOleDirectoryEntry In attachmentStorages
                    attachmentIndex += 1

                    Dim fileName As String = ReadMsgStringProperty(oleDoc, attachStorage, "3707", ReadMsgPreferredCodePage(oleDoc, rootEntry))
                    If System.String.IsNullOrWhiteSpace(fileName) Then
                        fileName = ReadMsgStringProperty(oleDoc, attachStorage, "3704", ReadMsgPreferredCodePage(oleDoc, rootEntry))
                    End If

                    fileName = SanitizeAttachmentFileName(
                        fileName,
                        System.IO.Path.GetFileNameWithoutExtension(sourceFileName) & "_attachment_" & attachmentIndex.ToString("000", System.Globalization.CultureInfo.InvariantCulture) & ".bin")

                    Dim data As Byte() = ReadMsgBinaryProperty(oleDoc, attachStorage, "3701")
                    Dim attachmentText As String
                    Dim finalAttachmentName As String = fileName

                    If data Is Nothing OrElse data.Length = 0 Then
                        attachmentText = "[Skipped: attachment '" & fileName & "' has no readable binary content stream]"
                    Else
                        Dim attachmentPath As String = System.IO.Path.Combine(tempDir, fileName)
                        Dim suffix As Integer = 1

                        While System.IO.File.Exists(attachmentPath)
                            attachmentPath = System.IO.Path.Combine(
                                tempDir,
                                System.IO.Path.GetFileNameWithoutExtension(fileName) &
                                "_" & suffix.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                System.IO.Path.GetExtension(fileName))
                            suffix += 1
                        End While

                        System.IO.File.WriteAllBytes(attachmentPath, data)
                        finalAttachmentName = System.IO.Path.GetFileName(attachmentPath)
                        attachmentText = ExtractInlineAttachmentTextFromSavedFile(attachmentPath, depth + 1)
                    End If

                    Dim section As String = BuildInlineAttachmentSection(finalAttachmentName, attachmentText)
                    If section <> "" Then
                        result.Add(section)
                    End If
                Next

            Catch ex As System.Exception
            Finally
                Try
                    If System.IO.Directory.Exists(tempDir) Then
                        System.IO.Directory.Delete(tempDir, recursive:=True)
                    End If
                Catch ex As System.Exception
                End Try
            End Try

            Return result
        End Function

        Private Shared Function BuildFriendlyMsgAddressValue(headerValue As String,
                                                             directName As String,
                                                             directEmail As String) As String
            Dim cleanHeaderValue As String = CleanFriendlyHeaderValue(headerValue)
            Dim cleanName As String = CleanFriendlyHeaderValue(directName)
            Dim cleanEmail As String = CleanFriendlyHeaderValue(directEmail)

            Dim smtpFromHeader As String = ExtractFirstEmailAddress(cleanHeaderValue)
            Dim smtpFromDirect As String = ExtractFirstEmailAddress(cleanEmail)

            If smtpFromDirect <> "" Then
                cleanEmail = smtpFromDirect
            ElseIf smtpFromHeader <> "" Then
                cleanEmail = smtpFromHeader
            End If

            If cleanName <> "" AndAlso cleanEmail <> "" Then
                cleanName = cleanName.Replace(cleanEmail, "").Trim()
                cleanName = cleanName.Trim("<"c, ">"c, ";"c, ","c, " "c)
            End If

            If cleanName <> "" AndAlso cleanEmail <> "" Then
                Return cleanName & " <" & cleanEmail & ">"
            End If

            If cleanEmail <> "" Then Return cleanEmail
            If cleanHeaderValue <> "" Then Return cleanHeaderValue
            Return cleanName
        End Function

        Private Shared Function BuildFriendlyMsgRecipientList(headerValue As String,
                                                              directValue As String) As String
            Dim cleanHeaderValue As String = CleanFriendlyHeaderValue(headerValue)

            If cleanHeaderValue <> "" AndAlso
               cleanHeaderValue.IndexOf("/O=", System.StringComparison.OrdinalIgnoreCase) < 0 AndAlso
               cleanHeaderValue.IndexOf("EXCHANGELABS", System.StringComparison.OrdinalIgnoreCase) < 0 Then
                Return cleanHeaderValue
            End If

            Dim addresses As New System.Collections.Generic.List(Of String)()

            Try
                Dim matches As System.Text.RegularExpressions.MatchCollection =
                    System.Text.RegularExpressions.Regex.Matches(
                        If(directValue, ""),
                        "(?i)\b[a-z0-9._%+\-']+@[a-z0-9.\-]+\.[a-z]{2,}\b")

                For Each m As System.Text.RegularExpressions.Match In matches
                    Dim addr As String = m.Value.Trim()

                    If addr <> "" AndAlso Not ContainsStringOrdinalIgnoreCase(addresses, addr) Then
                        addresses.Add(addr)
                    End If
                Next
            Catch ex As System.Exception
            End Try

            If addresses.Count > 0 Then
                Return System.String.Join("; ", addresses)
            End If

            If cleanHeaderValue <> "" Then Return cleanHeaderValue
            Return CleanFriendlyHeaderValue(directValue)
        End Function

        Private Shared Function ExtractHeaderValueFromMsgHeaderBlock(headerBlock As String,
                                                                     headerName As String) As String
            If System.String.IsNullOrWhiteSpace(headerBlock) OrElse System.String.IsNullOrWhiteSpace(headerName) Then Return ""

            Try
                headerBlock = NormalizeHeaderBlockLineEndings(headerBlock)

                Dim pattern As String =
                    "(?im)^" &
                    System.Text.RegularExpressions.Regex.Escape(headerName) &
                    ":\s*(.+?)(?=\r?\n[!-9;-~]+:|\z)"

                Dim matchResult As System.Text.RegularExpressions.Match =
                    System.Text.RegularExpressions.Regex.Match(
                        headerBlock,
                        pattern,
                        System.Text.RegularExpressions.RegexOptions.Singleline Or
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase)

                If Not matchResult.Success Then Return ""

                Dim value As String = matchResult.Groups(1).Value
                value = System.Text.RegularExpressions.Regex.Replace(value, "\r?\n[ \t]+", " ")
                Return CleanFriendlyHeaderValue(value.Trim())
            Catch ex As System.Exception
                Return ""
            End Try
        End Function

        Private Shared Function ExtractFirstEmailAddress(value As String) As String
            If System.String.IsNullOrWhiteSpace(value) Then Return ""

            Try
                Dim matchResult As System.Text.RegularExpressions.Match =
                    System.Text.RegularExpressions.Regex.Match(
                        value,
                        "(?i)\b[a-z0-9._%+\-']+@[a-z0-9.\-]+\.[a-z]{2,}\b")

                If matchResult.Success Then
                    Return matchResult.Value.Trim()
                End If
            Catch ex As System.Exception
            End Try

            Return ""
        End Function

        Private Shared Function CleanMsgExtractedText(value As String) As String
            If value Is Nothing Then Return ""

            value = value.Replace(ChrW(0), "")
            value = value.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf)
            value = System.Text.RegularExpressions.Regex.Replace(value, "[\u0001-\u0008\u000B\u000C\u000E-\u001F]", "")
            value = value.Trim()

            Return value
        End Function

        Private Shared Function IsUsableMsgText(value As String) As Boolean
            If System.String.IsNullOrWhiteSpace(value) Then Return False

            Dim text As String = value.Trim()

            If text.IndexOf("__substg1.0_", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return False
            If text.IndexOf("__recip_version1.0_", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return False
            If text.IndexOf("__attach_version1.0_", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return False

            Dim totalCount As Integer = 0
            Dim badCount As Integer = 0

            For Each ch As Char In text
                Dim code As Integer = AscW(ch)

                If ch = vbCr OrElse ch = vbLf OrElse ch = vbTab Then
                    Continue For
                End If

                totalCount += 1

                If code = 0 Then
                    badCount += 1
                ElseIf code < 32 Then
                    badCount += 1
                ElseIf code >= &HD800 AndAlso code <= &HDFFF Then
                    badCount += 1
                End If
            Next

            If totalCount = 0 Then Return False

            Dim badRatio As Double = CDbl(badCount) / CDbl(totalCount)
            Return badRatio <= 0.03
        End Function

        Private Shared Function IsLikelyMsgNoiseLine(value As String) As Boolean
            If System.String.IsNullOrWhiteSpace(value) Then Return True

            Dim text As String = value.Trim()
            If text = "" Then Return True

            If text.IndexOf("__substg1.0_", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
            If text.IndexOf("__recip_version1.0_", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
            If text.IndexOf("__attach_version1.0_", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return True

            Dim totalCount As Integer = 0
            Dim badCount As Integer = 0

            For Each ch As Char In text
                Dim code As Integer = AscW(ch)

                If ch = vbCr OrElse ch = vbLf OrElse ch = vbTab Then
                    Continue For
                End If

                totalCount += 1

                If code = 0 Then
                    badCount += 1
                ElseIf code < 32 Then
                    badCount += 1
                ElseIf code >= &HD800 AndAlso code <= &HDFFF Then
                    badCount += 1
                End If
            Next

            If totalCount = 0 Then Return True

            Dim badRatio As Double = CDbl(badCount) / CDbl(totalCount)
            Return badRatio > 0.2
        End Function

        Private Shared Function DecodeMsgTextBytes(data As Byte(), preferredCodePage As Integer) As String
            If data Is Nothing OrElse data.Length = 0 Then Return ""

            Try
                If data.Length >= 2 AndAlso data(0) = &HFF AndAlso data(1) = &HFE Then
                    Return System.Text.Encoding.Unicode.GetString(data)
                End If

                If data.Length >= 2 AndAlso data(0) = &HFE AndAlso data(1) = &HFF Then
                    Return System.Text.Encoding.BigEndianUnicode.GetString(data)
                End If

                If data.Length >= 3 AndAlso data(0) = &HEF AndAlso data(1) = &HBB AndAlso data(2) = &HBF Then
                    Return System.Text.Encoding.UTF8.GetString(data)
                End If

                Dim strictUtf8 As New System.Text.UTF8Encoding(False, True)
                Return strictUtf8.GetString(data)
            Catch ex As System.Exception
            End Try

            Try
                Return GetEncodingByNameSafe(preferredCodePage.ToString(System.Globalization.CultureInfo.InvariantCulture)).GetString(data)
            Catch ex As System.Exception
            End Try

            Return System.Text.Encoding.Default.GetString(data)
        End Function

        Private Shared Function RemoveTrailingNullBytes(data As Byte(), characterWidth As Integer) As Byte()
            If data Is Nothing OrElse data.Length = 0 Then Return System.Array.Empty(Of Byte)()

            Dim endIndex As Integer = data.Length

            If characterWidth <= 1 Then
                While endIndex > 0 AndAlso data(endIndex - 1) = 0
                    endIndex -= 1
                End While
            Else
                While endIndex >= 2 AndAlso data(endIndex - 1) = 0 AndAlso data(endIndex - 2) = 0
                    endIndex -= 2
                End While
            End If

            If endIndex = data.Length Then Return data

            Dim result(endIndex - 1) As Byte
            System.Array.Copy(data, 0, result, 0, endIndex)
            Return result
        End Function

        Private Shared Function ContainsStringOrdinalIgnoreCase(items As System.Collections.Generic.List(Of String),
                                                                value As String) As Boolean
            If items Is Nothing Then Return False

            For Each item As String In items
                If System.String.Equals(item, value, System.StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private NotInheritable Class MsgOleDirectoryEntry
            Public Property Id As Integer
            Public Property Name As String
            Public Property ObjectType As Byte
            Public Property LeftSiblingId As Integer
            Public Property RightSiblingId As Integer
            Public Property ChildId As Integer
            Public Property StartSector As Integer
            Public Property StreamSize As Long

            Public ReadOnly Property IsStorage As Boolean
                Get
                    Return ObjectType = 1 OrElse ObjectType = 5
                End Get
            End Property

            Public ReadOnly Property IsStream As Boolean
                Get
                    Return ObjectType = 2
                End Get
            End Property
        End Class

        Private NotInheritable Class MsgOleCompoundDocument
            Implements System.IDisposable

            Private Const OleFreeSector As Integer = -1
            Private Const OleEndOfChain As Integer = -2
            Private Const OleFatSector As Integer = -3
            Private Const OleDifatSector As Integer = -4
            Private Const OleNoStream As Integer = -1

            Private ReadOnly _fileBytes As Byte()
            Private ReadOnly _sectorSize As Integer
            Private ReadOnly _miniSectorSize As Integer
            Private ReadOnly _miniStreamCutoffSize As Integer
            Private ReadOnly _fat As System.Collections.Generic.List(Of Integer)
            Private ReadOnly _miniFat As System.Collections.Generic.List(Of Integer)
            Private ReadOnly _entries As System.Collections.Generic.List(Of MsgOleDirectoryEntry)
            Private ReadOnly _miniStreamBytes As Byte()
            Private _disposed As Boolean

            Public Sub New(filePath As String)
                _fileBytes = System.IO.File.ReadAllBytes(filePath)

                If _fileBytes Is Nothing OrElse _fileBytes.Length < 512 Then
                    Throw New System.IO.InvalidDataException("The file is too small to be an OLE Compound File.")
                End If

                If Not HasValidOleHeader(_fileBytes) Then
                    Throw New System.IO.InvalidDataException("The file is not an OLE Compound File.")
                End If

                Dim byteOrder As UShort = ReadUInt16(_fileBytes, 28)
                If byteOrder <> &HFFFEUS Then
                    Throw New System.IO.InvalidDataException("Unsupported OLE byte order.")
                End If

                Dim sectorShift As UShort = ReadUInt16(_fileBytes, 30)
                Dim miniSectorShift As UShort = ReadUInt16(_fileBytes, 32)

                _sectorSize = 1 << sectorShift
                _miniSectorSize = 1 << miniSectorShift

                If _sectorSize < 512 OrElse _sectorSize > 4096 Then
                    Throw New System.IO.InvalidDataException("Unsupported OLE sector size.")
                End If

                If _miniSectorSize <= 0 OrElse _miniSectorSize > _sectorSize Then
                    Throw New System.IO.InvalidDataException("Unsupported OLE mini-sector size.")
                End If

                Dim numFatSectors As Integer = ReadInt32(_fileBytes, 44)
                Dim firstDirectorySector As Integer = ReadInt32(_fileBytes, 48)
                _miniStreamCutoffSize = ReadInt32(_fileBytes, 56)
                Dim firstMiniFatSector As Integer = ReadInt32(_fileBytes, 60)
                Dim numMiniFatSectors As Integer = ReadInt32(_fileBytes, 64)
                Dim firstDifatSector As Integer = ReadInt32(_fileBytes, 68)
                Dim numDifatSectors As Integer = ReadInt32(_fileBytes, 72)

                Dim fatSectorIds As System.Collections.Generic.List(Of Integer) =
                    LoadDifatSectorIds(numFatSectors, firstDifatSector, numDifatSectors)

                _fat = LoadFat(fatSectorIds)
                _entries = LoadDirectoryEntries(firstDirectorySector)

                If _entries.Count = 0 Then
                    Throw New System.IO.InvalidDataException("The OLE directory is empty.")
                End If

                _miniFat = LoadMiniFat(firstMiniFatSector, numMiniFatSectors)
                _miniStreamBytes = LoadMiniStreamBytes()
            End Sub

            Public ReadOnly Property RootEntry As MsgOleDirectoryEntry
                Get
                    Return _entries(0)
                End Get
            End Property

            Public Function ReadChildStream(parentEntry As MsgOleDirectoryEntry, childName As String) As Byte()
                Dim childEntry As MsgOleDirectoryEntry = GetChildByName(parentEntry, childName)
                If childEntry Is Nothing OrElse Not childEntry.IsStream Then
                    Return System.Array.Empty(Of Byte)()
                End If

                Return ReadStream(childEntry)
            End Function

            Public Function GetChildStorages(parentEntry As MsgOleDirectoryEntry,
                                             namePrefix As String) As System.Collections.Generic.List(Of MsgOleDirectoryEntry)
                Dim result As New System.Collections.Generic.List(Of MsgOleDirectoryEntry)()
                Dim children As System.Collections.Generic.List(Of MsgOleDirectoryEntry) = GetChildren(parentEntry)

                For Each child As MsgOleDirectoryEntry In children
                    If child IsNot Nothing AndAlso child.IsStorage AndAlso
                       child.Name.StartsWith(namePrefix, System.StringComparison.OrdinalIgnoreCase) Then
                        result.Add(child)
                    End If
                Next

                result.Sort(Function(a As MsgOleDirectoryEntry, b As MsgOleDirectoryEntry) As Integer
                                Return System.StringComparer.OrdinalIgnoreCase.Compare(a.Name, b.Name)
                            End Function)

                Return result
            End Function

            Private Function GetChildByName(parentEntry As MsgOleDirectoryEntry,
                                            childName As String) As MsgOleDirectoryEntry
                Dim children As System.Collections.Generic.List(Of MsgOleDirectoryEntry) = GetChildren(parentEntry)

                For Each child As MsgOleDirectoryEntry In children
                    If System.String.Equals(child.Name, childName, System.StringComparison.OrdinalIgnoreCase) Then
                        Return child
                    End If
                Next

                Return Nothing
            End Function

            Private Function GetChildren(parentEntry As MsgOleDirectoryEntry) As System.Collections.Generic.List(Of MsgOleDirectoryEntry)
                Dim result As New System.Collections.Generic.List(Of MsgOleDirectoryEntry)()

                If parentEntry Is Nothing OrElse parentEntry.ChildId = OleNoStream Then
                    Return result
                End If

                AddDirectoryTreeEntries(parentEntry.ChildId, result)
                Return result
            End Function

            Private Sub AddDirectoryTreeEntries(entryId As Integer,
                                                result As System.Collections.Generic.List(Of MsgOleDirectoryEntry))
                If entryId = OleNoStream OrElse entryId < 0 OrElse entryId >= _entries.Count Then
                    Return
                End If

                Dim entry As MsgOleDirectoryEntry = _entries(entryId)

                If entry.LeftSiblingId <> OleNoStream Then
                    AddDirectoryTreeEntries(entry.LeftSiblingId, result)
                End If

                If entry.ObjectType <> 0 Then
                    result.Add(entry)
                End If

                If entry.RightSiblingId <> OleNoStream Then
                    AddDirectoryTreeEntries(entry.RightSiblingId, result)
                End If
            End Sub

            Private Function ReadStream(entry As MsgOleDirectoryEntry) As Byte()
                If entry Is Nothing OrElse Not entry.IsStream Then Return System.Array.Empty(Of Byte)()
                If entry.StreamSize <= 0 Then Return System.Array.Empty(Of Byte)()

                If entry.StreamSize < _miniStreamCutoffSize AndAlso _miniFat IsNot Nothing AndAlso _miniFat.Count > 0 AndAlso _miniStreamBytes IsNot Nothing Then
                    Return ReadMiniStreamChain(entry.StartSector, entry.StreamSize)
                End If

                Return ReadRegularStreamChain(entry.StartSector, entry.StreamSize)
            End Function

            Private Function LoadDifatSectorIds(numFatSectors As Integer,
                                                firstDifatSector As Integer,
                                                numDifatSectors As Integer) As System.Collections.Generic.List(Of Integer)
                Dim result As New System.Collections.Generic.List(Of Integer)()

                For i As Integer = 0 To 108
                    Dim sectorId As Integer = ReadInt32(_fileBytes, 76 + i * 4)
                    If sectorId >= 0 Then
                        result.Add(sectorId)
                    End If
                Next

                Dim currentDifatSector As Integer = firstDifatSector
                Dim difatSectorsRead As Integer = 0

                While currentDifatSector >= 0 AndAlso currentDifatSector <> OleEndOfChain AndAlso difatSectorsRead < numDifatSectors
                    Dim sectorBytes As Byte() = GetSectorBytes(currentDifatSector)
                    Dim entriesPerDifatSector As Integer = (_sectorSize \ 4) - 1

                    For i As Integer = 0 To entriesPerDifatSector - 1
                        Dim sectorId As Integer = ReadInt32(sectorBytes, i * 4)
                        If sectorId >= 0 Then
                            result.Add(sectorId)
                        End If
                    Next

                    currentDifatSector = ReadInt32(sectorBytes, entriesPerDifatSector * 4)
                    difatSectorsRead += 1
                End While

                If result.Count < numFatSectors Then
                    Throw New System.IO.InvalidDataException("The OLE FAT sector list is incomplete.")
                End If

                If result.Count > numFatSectors Then
                    result.RemoveRange(numFatSectors, result.Count - numFatSectors)
                End If

                Return result
            End Function

            Private Function LoadFat(fatSectorIds As System.Collections.Generic.List(Of Integer)) As System.Collections.Generic.List(Of Integer)
                Dim result As New System.Collections.Generic.List(Of Integer)()

                For Each fatSectorId As Integer In fatSectorIds
                    Dim sectorBytes As Byte() = GetSectorBytes(fatSectorId)
                    Dim entriesPerSector As Integer = _sectorSize \ 4

                    For i As Integer = 0 To entriesPerSector - 1
                        result.Add(ReadInt32(sectorBytes, i * 4))
                    Next
                Next

                Return result
            End Function

            Private Function LoadDirectoryEntries(firstDirectorySector As Integer) As System.Collections.Generic.List(Of MsgOleDirectoryEntry)
                Dim directoryBytes As Byte() = ReadSectorChain(firstDirectorySector, _fat, -1)
                Dim result As New System.Collections.Generic.List(Of MsgOleDirectoryEntry)()

                If directoryBytes Is Nothing OrElse directoryBytes.Length = 0 Then
                    Return result
                End If

                Dim entryCount As Integer = directoryBytes.Length \ 128

                For i As Integer = 0 To entryCount - 1
                    Dim offset As Integer = i * 128
                    Dim nameLength As UShort = ReadUInt16(directoryBytes, offset + 64)
                    Dim entryName As String = ""

                    If nameLength >= 2 AndAlso nameLength <= 64 Then
                        entryName = System.Text.Encoding.Unicode.GetString(directoryBytes, offset, nameLength - 2)
                    End If

                    Dim entry As New MsgOleDirectoryEntry()
                    entry.Id = i
                    entry.Name = entryName
                    entry.ObjectType = directoryBytes(offset + 66)
                    entry.LeftSiblingId = ReadInt32(directoryBytes, offset + 68)
                    entry.RightSiblingId = ReadInt32(directoryBytes, offset + 72)
                    entry.ChildId = ReadInt32(directoryBytes, offset + 76)
                    entry.StartSector = ReadInt32(directoryBytes, offset + 116)

                    If _sectorSize = 512 Then
                        entry.StreamSize = ReadUInt32(directoryBytes, offset + 120)
                    Else
                        entry.StreamSize = ReadInt64(directoryBytes, offset + 120)
                    End If

                    result.Add(entry)
                Next

                Return result
            End Function

            Private Function LoadMiniFat(firstMiniFatSector As Integer,
                                         numMiniFatSectors As Integer) As System.Collections.Generic.List(Of Integer)
                Dim result As New System.Collections.Generic.List(Of Integer)()

                If firstMiniFatSector < 0 OrElse firstMiniFatSector = OleEndOfChain OrElse numMiniFatSectors <= 0 Then
                    Return result
                End If

                Dim miniFatBytes As Byte() = ReadSectorChain(firstMiniFatSector, _fat, CLng(numMiniFatSectors) * CLng(_sectorSize))
                If miniFatBytes Is Nothing OrElse miniFatBytes.Length = 0 Then
                    Return result
                End If

                Dim entryCount As Integer = miniFatBytes.Length \ 4
                For i As Integer = 0 To entryCount - 1
                    result.Add(ReadInt32(miniFatBytes, i * 4))
                Next

                Return result
            End Function

            Private Function LoadMiniStreamBytes() As Byte()
                If _entries Is Nothing OrElse _entries.Count = 0 Then Return System.Array.Empty(Of Byte)()

                Dim root As MsgOleDirectoryEntry = _entries(0)
                If root.StartSector < 0 OrElse root.StreamSize <= 0 Then Return System.Array.Empty(Of Byte)()

                Return ReadSectorChain(root.StartSector, _fat, root.StreamSize)
            End Function

            Private Function ReadRegularStreamChain(startSector As Integer, streamSize As Long) As Byte()
                Return ReadSectorChain(startSector, _fat, streamSize)
            End Function

            Private Function ReadMiniStreamChain(startMiniSector As Integer, streamSize As Long) As Byte()
                If startMiniSector < 0 OrElse _miniStreamBytes Is Nothing OrElse _miniStreamBytes.Length = 0 Then
                    Return System.Array.Empty(Of Byte)()
                End If

                Dim result As New System.IO.MemoryStream()
                Dim currentMiniSector As Integer = startMiniSector
                Dim remainingBytes As Long = streamSize
                Dim guard As Integer = 0

                While currentMiniSector >= 0 AndAlso currentMiniSector <> OleEndOfChain AndAlso remainingBytes > 0 AndAlso guard < _miniFat.Count + 1
                    Dim offset As Long = CLng(currentMiniSector) * CLng(_miniSectorSize)
                    If offset < 0 OrElse offset >= _miniStreamBytes.Length Then Exit While

                    Dim available As Integer = CInt(System.Math.Min(CLng(_miniSectorSize), CLng(_miniStreamBytes.Length) - offset))
                    Dim bytesToWrite As Integer = CInt(System.Math.Min(CLng(available), remainingBytes))

                    result.Write(_miniStreamBytes, CInt(offset), bytesToWrite)
                    remainingBytes -= bytesToWrite

                    If currentMiniSector < 0 OrElse currentMiniSector >= _miniFat.Count Then Exit While
                    currentMiniSector = _miniFat(currentMiniSector)
                    guard += 1
                End While

                Return result.ToArray()
            End Function

            Private Function ReadSectorChain(startSector As Integer,
                                             fatTable As System.Collections.Generic.List(Of Integer),
                                             maxBytes As Long) As Byte()
                If startSector < 0 OrElse fatTable Is Nothing OrElse fatTable.Count = 0 Then
                    Return System.Array.Empty(Of Byte)()
                End If

                Dim result As New System.IO.MemoryStream()
                Dim currentSector As Integer = startSector
                Dim remainingBytes As Long = maxBytes
                Dim guard As Integer = 0

                While currentSector >= 0 AndAlso currentSector <> OleEndOfChain AndAlso guard < fatTable.Count + 1
                    Dim sectorBytes As Byte() = GetSectorBytes(currentSector)

                    If maxBytes >= 0 Then
                        Dim bytesToWrite As Integer = CInt(System.Math.Min(CLng(sectorBytes.Length), remainingBytes))
                        If bytesToWrite <= 0 Then Exit While
                        result.Write(sectorBytes, 0, bytesToWrite)
                        remainingBytes -= bytesToWrite
                        If remainingBytes <= 0 Then Exit While
                    Else
                        result.Write(sectorBytes, 0, sectorBytes.Length)
                    End If

                    If currentSector < 0 OrElse currentSector >= fatTable.Count Then Exit While
                    currentSector = fatTable(currentSector)
                    guard += 1
                End While

                Return result.ToArray()
            End Function

            Private Function GetSectorBytes(sectorId As Integer) As Byte()
                If sectorId < 0 Then
                    Throw New System.IO.InvalidDataException("Invalid OLE sector id.")
                End If

                Dim offset As Long = 512L + CLng(sectorId) * CLng(_sectorSize)

                If offset < 0 OrElse offset + _sectorSize > _fileBytes.LongLength Then
                    Throw New System.IO.InvalidDataException("OLE sector points outside the file.")
                End If

                Dim result(_sectorSize - 1) As Byte
                System.Array.Copy(_fileBytes, CInt(offset), result, 0, _sectorSize)
                Return result
            End Function

            Private Shared Function HasValidOleHeader(data As Byte()) As Boolean
                If data Is Nothing OrElse data.Length < 8 Then Return False

                Dim expectedHeader As Byte() = New Byte() {&HD0, &HCF, &H11, &HE0, &HA1, &HB1, &H1A, &HE1}

                For i As Integer = 0 To expectedHeader.Length - 1
                    If data(i) <> expectedHeader(i) Then Return False
                Next

                Return True
            End Function

            Private Shared Function ReadUInt16(data As Byte(), offset As Integer) As UShort
                Return System.BitConverter.ToUInt16(data, offset)
            End Function

            Private Shared Function ReadInt32(data As Byte(), offset As Integer) As Integer
                Return System.BitConverter.ToInt32(data, offset)
            End Function

            Private Shared Function ReadUInt32(data As Byte(), offset As Integer) As UInteger
                Return System.BitConverter.ToUInt32(data, offset)
            End Function

            Private Shared Function ReadInt64(data As Byte(), offset As Integer) As Long
                Return System.BitConverter.ToInt64(data, offset)
            End Function

            Public Sub Dispose() Implements System.IDisposable.Dispose
                _disposed = True
            End Sub
        End Class

    End Class
End Namespace
