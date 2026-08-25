' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.DocProcessor.vb
' Purpose:
'   AutoPilot document-processing engine for Office attachments. Applies an
'   instruction via LLM to OpenXML package content while preserving structural
'   fidelity and formatting boundaries.
'
' Scope:
'  - DOCX processing (WordprocessingML)
'  - PPTX processing (DrawingML in slides and notes slides)
'  - XLSX processing (SpreadsheetML cells, formulas, and shared strings)
'
' Architecture:
'  - OpenXML-first pipeline (no Word/PowerPoint/Excel interop for core mutation).
'  - Unzips package parts, extracts processable text units, batches content for LLM,
'    parses structured responses, writes updates back to XML parts, then repacks.
'  - DOCX paragraph/run model:
'      * Preserves real formatting-group boundaries via `|` markers where applicable.
'      * Preserves footnote/endnote/field/manual-break boundary anchors via `‖` marker.
'      * Reapplies text proportionally when exact marker alignment is unavailable.
'  - Sub-part coverage:
'      * DOCX: `document.xml`, headers, footers, comments, footnotes, endnotes.
'      * PPTX: `slide*.xml` and `notesSlide*.xml`.
'      * XLSX: worksheet parts, shared strings table, workbook relationships,
'        and content-type registration when creating shared strings on demand.
'  - Batch strategy:
'      * Context-before/context-after windows; each LLM batch is explicitly treated as isolated.
'      * Paragraph/cell chunking with character cap and cancellation support.
'      * Caller-supplied document-wide consistency context is repeated through the instruction when provided.
'      * Structured response parsing (`[n] ...` for paragraph batches,
'        `[A1] ...` for spreadsheet cells).
'
' Safety & Integrity:
'  - Operates on copied output file, preserving input file unchanged.
'  - Maintains XML whitespace and declaration-sensitive saves for OOXML compatibility.
'  - Avoids namespace corruption by creating/importing elements in proper namespace context.
'  - Clears temporary extraction directories in `Finally` paths.
'
' Notes:
'  - Compare document generation for Word files is handled by tooling logic in
'    `ThisAddIn.AutoPIlot.Tools.vb` (`CreateWordCompareDocumentForAutoPilot`).
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' ═══════════════════════════════════════════════════════════════════════════
    '  CONSTANTS 
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Number of paragraphs included before the current batch as context.
    ''' </summary>
    Private Const AP_ContextBefore As Integer = 3

    ''' <summary>
    ''' Number of paragraphs included after the current batch as context.
    ''' </summary>
    Private Const AP_ContextAfter As Integer = 2

    ''' <summary>
    ''' Maximum number of paragraphs sent to the LLM per batch.
    ''' </summary>
    Private Const AP_ParagraphsPerBatch As Integer = 10

    ''' <summary>
    ''' Maximum total characters per batch.
    ''' </summary>
    Private Const AP_MaxCharsPerBatch As Integer = 15000

    ''' <summary>Marker inserted between formatting runs so the LLM can preserve boundaries.</summary>
    Private Const AP_RunBoundaryMarker As String = "|"

    ''' <summary>
    ''' Marker inserted at protected structural boundaries in text sent to the LLM.
    ''' U+2016 DOUBLE VERTICAL LINE — virtually never appears in legal documents.
    ''' </summary>
    Private Const AP_NoteRefMarker As String = "‖"

    ''' <summary>
    ''' Characters treated as non-breaking spaces that must be preserved.
    ''' U+00A0 = non-breaking space (geschütztes Leerzeichen)
    ''' U+202F = narrow no-break space (schmales geschütztes Leerzeichen)
    ''' U+2007 = figure space (ziffernbreites Leerzeichen)
    ''' </summary>
    Private Shared ReadOnly AP_NonBreakingSpaceChars As Char() = {ChrW(&HA0), ChrW(&H202F), ChrW(&H2007)}

    ' ═══════════════════════════════════════════════════════════════════════════
    '  DATA CLASSES (local to AutoPilot, avoid collision with Word add-in)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Represents a non-breaking space occurrence with surrounding word context for restoration.
    ''' </summary>
    Private Class APNonBreakingSpaceInfo
        ''' <summary>The non-breaking space character (U+00A0, U+202F, etc.).</summary>
        Public Property SpaceChar As Char
        ''' <summary>The word immediately before the non-breaking space (Nothing if at start).</summary>
        Public Property WordBefore As String
        ''' <summary>The word immediately after the non-breaking space (Nothing if at end).</summary>
        Public Property WordAfter As String
    End Class

    ''' <summary>
    ''' Stores a text node and its original content for a single run.
    ''' </summary>
    Private Class APTextRunInfo
        Public Property TextNode As System.Xml.XmlNode
        Public Property OriginalText As String
        Public Property HasNoteReferenceBefore As Boolean
        Public Property FormattingKey As String
    End Class

    ''' <summary>
    ''' Represents a paragraph, its runs, and processed text state.
    ''' </summary>
    Private Class APParagraphInfo
        Public Property Index As Integer
        Public Property TextRuns As List(Of APTextRunInfo)
        Public Property FullText As String

        ''' Builds text with | markers only between real formatting groups.
        Public Property MarkerText As String

        Public Property TranslatedText As String
        Public Property IsEmpty As Boolean
    End Class


    ''' <summary>
    ''' Returns a stable run-formatting key for DOCX/PPTX text. Adjacent runs with the same
    ''' key can be treated as one formatting group without exposing Office run fragmentation.
    ''' </summary>
    Private Shared Function APGetRunFormattingKey(textNode As System.Xml.XmlNode) As String
        If textNode Is Nothing Then Return ""

        Dim runNode As System.Xml.XmlNode = textNode.ParentNode
        While runNode IsNot Nothing AndAlso runNode.LocalName <> "r"
            runNode = runNode.ParentNode
        End While
        If runNode Is Nothing Then Return ""

        Dim sb As New System.Text.StringBuilder()
        For Each child As System.Xml.XmlNode In runNode.ChildNodes
            If child.NodeType = System.Xml.XmlNodeType.Element AndAlso child.LocalName = "rPr" Then
                sb.Append(child.OuterXml)
                Exit For
            End If
        Next

        Dim ancestor As System.Xml.XmlNode = runNode.ParentNode
        While ancestor IsNot Nothing AndAlso ancestor.LocalName <> "p"
            If ancestor.NodeType = System.Xml.XmlNodeType.Element Then
                sb.Append("|scope:")
                sb.Append(ancestor.LocalName)
                sb.Append(":")
                sb.Append(System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(ancestor).ToString())
            End If
            ancestor = ancestor.ParentNode
        End While

        Return sb.ToString()
    End Function

    Private Shared Function APBuildFormattingGroups(
        textRuns As List(Of APTextRunInfo),
        Optional runIndices As List(Of Integer) = Nothing) As List(Of List(Of Integer))

        Dim groups As New List(Of List(Of Integer))()
        If textRuns Is Nothing OrElse textRuns.Count = 0 Then Return groups

        Dim indices As New List(Of Integer)()
        If runIndices Is Nothing Then
            For i As Integer = 0 To textRuns.Count - 1
                indices.Add(i)
            Next
        Else
            indices.AddRange(runIndices)
        End If

        Dim currentGroup As List(Of Integer) = Nothing
        Dim previousKey As String = Nothing
        Dim havePrevious As Boolean = False

        For Each runIdx As Integer In indices
            If runIdx < 0 OrElse runIdx >= textRuns.Count Then Continue For
            Dim run As APTextRunInfo = textRuns(runIdx)
            If run Is Nothing OrElse String.IsNullOrEmpty(run.OriginalText) Then Continue For

            Dim currentKey As String = If(run.FormattingKey, "")
            If Not havePrevious OrElse Not String.Equals(previousKey, currentKey, StringComparison.Ordinal) Then
                currentGroup = New List(Of Integer)()
                groups.Add(currentGroup)
            End If

            currentGroup.Add(runIdx)
            previousKey = currentKey
            havePrevious = True
        Next

        Return groups
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  XLSX CONSTANTS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' SpreadsheetML namespace for worksheet elements.
    ''' </summary>
    Private Const AP_XlsxNs As String = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    ''' <summary>
    ''' Relationship namespace for workbook sheet-to-file mapping.
    ''' </summary>
    Private Const AP_XlsxRelNs As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


    ' ═══════════════════════════════════════════════════════════════════════════
    '  MAIN ENTRY POINT
    ' ═══════════════════════════════════════════════════════════════════════════


    ''' <summary>
    ''' Processes a DOCX file by applying the given instruction via LLM to all text paragraphs.
    ''' </summary>
    Private Async Function ProcessDocxForAutoPilot(inputPath As String, outputPath As String,
                                                    instruction As String, ct As CancellationToken) As Task(Of Boolean)
        Dim tempDir As String = Path.Combine(Path.GetTempPath(), AP_TempPrefix & "xml_" & Guid.NewGuid().ToString("N"))

        Try
            ApDashboardLog("DocProcessor: starting document processing", "step")

            ' Copy input to output (we modify the copy)
            File.Copy(inputPath, outputPath, overwrite:=True)

            ' Extract DOCX (ZIP)
            ZipFile.ExtractToDirectory(outputPath, tempDir)

            ' Process document.xml
            Dim documentXmlPath = Path.Combine(tempDir, "word", "document.xml")
            If Not File.Exists(documentXmlPath) Then Return False

            Dim xmlDoc As New System.Xml.XmlDocument()
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(documentXmlPath)

            Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
            nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

            Dim paragraphs = APExtractParagraphs(xmlDoc, nsMgr)
            If paragraphs.Count = 0 OrElse paragraphs.All(Function(p) p.IsEmpty) Then Return False

            ApDashboardLog($"DocProcessor: extracted {paragraphs.Count} paragraphs from document.xml", "step")

            Dim success = Await APProcessBatches(paragraphs, instruction, ct, "main body")
            If Not success Then Return False

            APApplyTranslations(paragraphs)

            ' Expand multi-line fills into separate XML paragraphs
            APExpandMultiLineParagraphs(paragraphs, nsMgr)

            xmlDoc.Save(documentXmlPath)

            ' Process headers, footers, comments, footnotes, endnotes
            ApDashboardLog("DocProcessor: processing sub-parts (headers, footers, etc.)", "step")
            Await APProcessSubParts(tempDir, instruction, ct)

            ' Repack DOCX
            File.Delete(outputPath)
            ZipFile.CreateFromDirectory(tempDir, outputPath, CompressionLevel.Optimal, False)

            ApDashboardLog("DocProcessor: document processing complete", "success")
            Return True

        Catch ex As System.Exception
            Debug.WriteLine("ProcessDocxForAutoPilot error: " & ex.Message)
            ApDashboardLog("DocProcessor: error - " & ex.Message, "error")
            Return False
        Finally
            If Directory.Exists(tempDir) Then
                Try : Directory.Delete(tempDir, recursive:=True) : Catch : End Try
            End If
        End Try
    End Function





    ''' <summary>
    ''' Expands paragraphs whose translated text contains line breaks into multiple
    ''' sibling w:p elements, cloning the original paragraph's formatting.
    ''' Must be called AFTER APApplyTranslations.
    ''' </summary>
    Private Sub APExpandMultiLineParagraphs(paragraphs As List(Of APParagraphInfo),
                                            nsMgr As System.Xml.XmlNamespaceManager)
        For Each para In paragraphs
            If para.IsEmpty OrElse String.IsNullOrEmpty(para.TranslatedText) Then Continue For
            If para.TextRuns.Count = 0 Then Continue For

            Dim lines As String() = para.TranslatedText.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            If lines.Length <= 1 Then Continue For

            Dim paraNode As System.Xml.XmlNode = para.TextRuns(0).TextNode
            While paraNode IsNot Nothing AndAlso paraNode.LocalName <> "p"
                paraNode = paraNode.ParentNode
            End While
            If paraNode Is Nothing OrElse paraNode.ParentNode Is Nothing Then Continue For

            Dim parentNode As System.Xml.XmlNode = paraNode.ParentNode

            APSetAllTextInParagraph(paraNode, nsMgr, lines(0).TrimEnd())

            Dim insertAfter As System.Xml.XmlNode = paraNode
            For lineIdx As Integer = 1 To lines.Length - 1
                Dim lineText As String = lines(lineIdx).TrimEnd()
                If String.IsNullOrEmpty(lineText) Then Continue For

                Dim newPara As System.Xml.XmlNode = paraNode.CloneNode(deep:=True)
                APSetAllTextInParagraph(newPara, nsMgr, lineText)
                parentNode.InsertAfter(newPara, insertAfter)
                insertAfter = newPara
            Next
        Next
    End Sub

    Private Sub APSetAllTextInParagraph(paraNode As System.Xml.XmlNode,
                                         nsMgr As System.Xml.XmlNamespaceManager,
                                         text As String)
        Dim tNodes = paraNode.SelectNodes(".//w:r/w:t", nsMgr)
        If tNodes Is Nothing OrElse tNodes.Count = 0 Then Return

        For i As Integer = 0 To tNodes.Count - 1
            If i = 0 Then
                APSetTextNode(tNodes(i), text)
            Else
                APSetTextNode(tNodes(i), "")
            End If
        Next
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PARAGRAPH EXTRACTION
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Extracts paragraphs and their text runs from the given document.
    ''' Detects footnote/endnote references, manual line breaks, AND complex field boundaries (cross-references,
    ''' merge fields, etc.) so that text redistribution never shifts content across them.
    ''' </summary>
    Private Function APExtractParagraphs(xmlDoc As System.Xml.XmlDocument,
                                          nsMgr As System.Xml.XmlNamespaceManager) As List(Of APParagraphInfo)
        Dim paragraphs As New List(Of APParagraphInfo)()
        Dim paraNodes = xmlDoc.SelectNodes("//w:p", nsMgr)
        Dim paraIndex As Integer = 0

        For Each paraNode As System.Xml.XmlNode In paraNodes

            Dim paraInfo As New APParagraphInfo() With {
                .Index = paraIndex,
                .TextRuns = New List(Of APTextRunInfo)(),
                .TranslatedText = Nothing
            }

            ' Identify w:t nodes to skip: footnoteRef / endnoteRef runs and their separator spaces
            ' (inside footnotes.xml / endnotes.xml)
            Dim refNodes As New HashSet(Of System.Xml.XmlNode)()
            Dim noteRefElements As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:r[w:footnoteRef or w:endnoteRef]", nsMgr)
            For Each refRunNode As System.Xml.XmlNode In noteRefElements
                For Each tInRef As System.Xml.XmlNode In refRunNode.SelectNodes(".//w:t", nsMgr)
                    refNodes.Add(tInRef)
                Next
                Dim nextSibling As System.Xml.XmlNode = refRunNode.NextSibling
                While nextSibling IsNot Nothing AndAlso nextSibling.NodeType <> System.Xml.XmlNodeType.Element
                    nextSibling = nextSibling.NextSibling
                End While
                If nextSibling IsNot Nothing AndAlso nextSibling.LocalName = "r" Then
                    Dim siblingTexts As System.Xml.XmlNodeList = nextSibling.SelectNodes(".//w:t", nsMgr)
                    If siblingTexts.Count = 1 Then
                        Dim sibText As String = siblingTexts(0).InnerText
                        If String.IsNullOrWhiteSpace(sibText) AndAlso sibText.Length <= 2 Then
                            refNodes.Add(siblingTexts(0))
                        End If
                    End If
                End If
            Next

            ' Identify w:r nodes containing w:footnoteReference / w:endnoteReference (document body)
            Dim bodyRefRunNodes As New HashSet(Of System.Xml.XmlNode)()
            Dim bodyRefElements As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:r[w:footnoteReference or w:endnoteReference]", nsMgr)
            For Each refRunNode As System.Xml.XmlNode In bodyRefElements
                bodyRefRunNodes.Add(refRunNode)
            Next

            ' ─── Complex field detection ───
            ' Complex fields (cross-references, merge fields, etc.) use w:fldChar elements.
            ' The w:r runs containing fldChar sit between text runs like footnote references.
            ' We treat them as boundaries and exclude field code text from extraction.
            Dim fldCharRuns As New HashSet(Of System.Xml.XmlNode)()
            Dim fldCharElements As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:r[w:fldChar]", nsMgr)
            For Each fldCharRunNode As System.Xml.XmlNode In fldCharElements
                fldCharRuns.Add(fldCharRunNode)
            Next

            ' Track field nesting to exclude w:t inside field code regions (begin→separate)
            If fldCharRuns.Count > 0 Then
                Dim fieldDepth As Integer = 0
                Dim inFieldCode As Boolean = False

                For Each childNode As System.Xml.XmlNode In paraNode.ChildNodes
                    If childNode.NodeType <> System.Xml.XmlNodeType.Element Then Continue For
                    If childNode.LocalName <> "r" Then Continue For

                    Dim fldChar As System.Xml.XmlNode = childNode.SelectSingleNode("w:fldChar", nsMgr)
                    If fldChar IsNot Nothing Then
                        Dim fldType As String = ""
                        Dim fldTypeAttr = fldChar.Attributes("w:fldCharType")
                        If fldTypeAttr IsNot Nothing Then fldType = fldTypeAttr.Value

                        Select Case fldType
                            Case "begin"
                                fieldDepth += 1
                                inFieldCode = True
                            Case "separate"
                                inFieldCode = False
                            Case "end"
                                fieldDepth -= 1
                                If fieldDepth <= 0 Then
                                    fieldDepth = 0
                                    inFieldCode = False
                                End If
                        End Select
                        Continue For
                    End If

                    If inFieldCode Then
                        For Each tNode As System.Xml.XmlNode In childNode.SelectNodes(".//w:t", nsMgr)
                            refNodes.Add(tNode)
                        Next
                    End If
                Next
            End If

            ' Protect manual Word line breaks (w:br / w:cr) as structural boundaries.
            ' The Outlook tool preserves formatting by default; these anchors also prevent
            ' otherwise-valid text redistribution from moving words across a manual break.
            Dim manualBreakBeforeTextNodes As New HashSet(Of System.Xml.XmlNode)()
            Dim orderedTextAndBreakNodes As System.Xml.XmlNodeList =
                paraNode.SelectNodes(".//w:t | .//w:br | .//w:cr", nsMgr)
            Dim havePreviousProcessableText As Boolean = False
            Dim manualBreakPending As Boolean = False

            For Each orderedNode As System.Xml.XmlNode In orderedTextAndBreakNodes
                If orderedNode.LocalName = "br" OrElse orderedNode.LocalName = "cr" Then
                    If havePreviousProcessableText Then manualBreakPending = True
                    Continue For
                End If

                If orderedNode.LocalName = "t" Then
                    If refNodes.Contains(orderedNode) Then Continue For
                    If manualBreakPending Then
                        manualBreakBeforeTextNodes.Add(orderedNode)
                    End If
                    havePreviousProcessableText = True
                    manualBreakPending = False
                End If
            Next

            Dim textNodes = paraNode.SelectNodes(".//w:t", nsMgr)
            Dim fullTextBuilder As New StringBuilder()

            For Each textNode As System.Xml.XmlNode In textNodes
                If refNodes.Contains(textNode) Then Continue For

                Dim text As String = textNode.InnerText

                ' Check if a footnoteReference/endnoteReference run OR a fldChar run
                ' sits between the previous text run and this one
                Dim hasNoteRefBefore As Boolean =
                    paraInfo.TextRuns.Count > 0 AndAlso manualBreakBeforeTextNodes.Contains(textNode)
                If paraInfo.TextRuns.Count > 0 AndAlso (bodyRefRunNodes.Count > 0 OrElse fldCharRuns.Count > 0) Then
                    Dim thisRun As System.Xml.XmlNode = textNode.ParentNode
                    Dim prevEl As System.Xml.XmlNode = thisRun.PreviousSibling
                    While prevEl IsNot Nothing
                        If prevEl.NodeType = System.Xml.XmlNodeType.Element Then
                            If prevEl.LocalName = "r" Then
                                If bodyRefRunNodes.Contains(prevEl) OrElse fldCharRuns.Contains(prevEl) Then
                                    hasNoteRefBefore = True
                                    Exit While
                                End If
                                Dim prevTexts As System.Xml.XmlNodeList = prevEl.SelectNodes(".//w:t", nsMgr)
                                Dim foundPrevTextRun As Boolean = False
                                For Each pt As System.Xml.XmlNode In prevTexts
                                    If Not refNodes.Contains(pt) Then
                                        foundPrevTextRun = True
                                        Exit For
                                    End If
                                Next
                                If foundPrevTextRun Then Exit While
                            End If
                        End If
                        prevEl = prevEl.PreviousSibling
                    End While
                End If

                paraInfo.TextRuns.Add(New APTextRunInfo() With {
                    .TextNode = textNode,
                    .OriginalText = text,
                    .HasNoteReferenceBefore = hasNoteRefBefore,
                    .FormattingKey = APGetRunFormattingKey(textNode)
                })
                ' Insert note-reference marker into FullText at the boundary
                If hasNoteRefBefore Then
                    fullTextBuilder.Append(AP_NoteRefMarker)
                End If

                fullTextBuilder.Append(text)
            Next

            paraInfo.FullText = fullTextBuilder.ToString()
            paraInfo.IsEmpty = String.IsNullOrWhiteSpace(paraInfo.FullText)

            If paraInfo.TextRuns.Count > 1 AndAlso Not paraInfo.IsEmpty Then
                paraInfo.MarkerText = APBuildMarkerAnnotatedText(paraInfo.TextRuns)
            Else
                paraInfo.MarkerText = Nothing
            End If

            paragraphs.Add(paraInfo)
            paraIndex += 1
        Next

        Return paragraphs
    End Function



    ''' <summary>
    ''' Builds text with | markers only between real formatting groups.
    ''' Only inserts markers between non-empty runs where formatting actually changes.
    ''' Also preserves ‖ markers at protected structural boundaries (notes, fields, manual breaks).
    ''' </summary>
    Private Shared Function APBuildMarkerAnnotatedText(textRuns As List(Of APTextRunInfo)) As String
        If textRuns Is Nothing OrElse textRuns.Count <= 1 Then Return Nothing

        Dim sb As New StringBuilder()
        Dim markerCount As Integer = 0
        Dim previousFormattingKey As String = Nothing
        Dim havePreviousVisibleRun As Boolean = False
        Dim hardBoundarySincePreviousVisibleRun As Boolean = False

        For i As Integer = 0 To textRuns.Count - 1
            Dim run As APTextRunInfo = textRuns(i)
            Dim runText As String = If(run.OriginalText, "")

            If run.HasNoteReferenceBefore Then
                sb.Append(AP_NoteRefMarker)
                hardBoundarySincePreviousVisibleRun = True
            End If

            If runText.Length > 0 Then
                Dim currentFormattingKey As String = If(run.FormattingKey, "")

                If havePreviousVisibleRun AndAlso
                   Not hardBoundarySincePreviousVisibleRun AndAlso
                   Not String.Equals(previousFormattingKey, currentFormattingKey, StringComparison.Ordinal) Then
                    sb.Append(AP_RunBoundaryMarker)
                    markerCount += 1
                End If

                sb.Append(runText)
                previousFormattingKey = currentFormattingKey
                havePreviousVisibleRun = True
                hardBoundarySincePreviousVisibleRun = False
            Else
                sb.Append(runText)
            End If
        Next

        If markerCount = 0 Then Return Nothing
        Return sb.ToString()
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  BATCH PROCESSING
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Batches paragraphs, sends them to the LLM, and stores processed results.
    ''' </summary>
    Private Async Function APProcessBatches(paragraphs As List(Of APParagraphInfo),
                                         instruction As String, ct As CancellationToken,
                                         Optional contentScope As String = "") As Task(Of Boolean)

        Dim processable = paragraphs.Where(Function(p) Not p.IsEmpty).ToList()
        If processable.Count = 0 Then Return True

        ' ─── Record non-breaking spaces for later restoration ───
        ' The LLM will normalize U+00A0, U+202F, etc. to regular spaces.
        ' We record context around each occurrence so we can restore them.
        Dim nbspRecords As New Dictionary(Of Integer, List(Of APNonBreakingSpaceInfo))()
        For i As Integer = 0 To processable.Count - 1
            Dim recorded = APRecordNonBreakingSpaces(processable(i).FullText)
            If recorded IsNot Nothing Then
                nbspRecords(i) = recorded
            End If
        Next

        ' Check if any paragraphs use markers
        Dim hasMarkers = processable.Any(Function(p) p.MarkerText IsNot Nothing)

        ' Build system prompt for document processing
        Dim systemPrompt As String =
            "You are a professional document processor. Apply the following instruction to the numbered paragraphs " &
            "in the [TEXTTOPROCESS] section." & vbCrLf & vbCrLf &
            "INSTRUCTION: " & instruction & vbCrLf & vbCrLf &
            "IMPORTANT BATCH SCOPE: The document is being processed in multiple sequential batches. This is an independent model call. " &
            "You do NOT have access to the full document and you do NOT retain hidden memory of earlier or later batches. In this call you see only " &
            "the current [TEXTTOPROCESS] plus the small [CONTEXT BEFORE] and [CONTEXT AFTER] windows supplied below. " &
            "The INSTRUCTION is repeated to every batch and is the authoritative document-wide contract. If it contains a " &
            "[DOCUMENT-WIDE CONSISTENCY CONTEXT] block, apply that block consistently throughout the document; it contains stable facts or decisions " &
            "that may not be visible in this local chunk and is context, not an additional operation. Reuse terminology and style choices visible in " &
            "[CONTEXT BEFORE] as precedent unless the document-wide instruction/context says otherwise. [CONTEXT AFTER] is unprocessed source context. " &
            "Do not invent facts, definitions, terminology choices, or other global conventions based on unseen text. " &
            "The INSTRUCTION describes the OVERALL task for the entire document — it does NOT mean that every paragraph in THIS batch needs to be changed. " &
            "Only modify paragraphs whose content is actually relevant to the INSTRUCTION. If a paragraph in this batch does not need any change, return it EXACTLY as-is." & vbCrLf & vbCrLf &
            "RULES:" & vbCrLf &
            "1. Process ONLY paragraphs inside [TEXTTOPROCESS], not the context sections." & vbCrLf &
            "2. Use [CONTEXT BEFORE] and [CONTEXT AFTER] to understand meaning, tone, and terminology." & vbCrLf &
            "3. Return each paragraph with its [n] marker exactly as shown — including paragraphs you did NOT change." & vbCrLf &
            "4. For paragraphs that DO NOT need changes based on their content, return them UNCHANGED — " &
            "copy the original text character for character." & vbCrLf &
            "5. For paragraphs that DO need changes, apply the INSTRUCTION precisely. Only modify what the " &
            "INSTRUCTION asks for. Do not rephrase, improve, or restructure text beyond the scope of the INSTRUCTION." & vbCrLf &
            "6. Maintain consistent terminology and style." & vbCrLf &
            "7. Return ONLY the [n] paragraphs, no explanations." & vbCrLf &
            "8. LANGUAGE RULE: NEVER translate or change the language of the document text unless the INSTRUCTION " &
            "EXPLICITLY requests translation to a specific target language. The INSTRUCTION itself may be in a " &
            "different language than the document — that does NOT mean you should translate. Apply the instruction " &
            "(e.g. correct, fill template, shorten) while preserving the original language of every paragraph." & vbCrLf &
            "9. TEMPLATE/PLACEHOLDER RULE: When filling a template or completing a document, preserve all " &
            "boilerplate, headings, and structural text verbatim. Only replace placeholders (text in brackets, " &
            "underscores, ellipsis …, guillemets, TBD, etc.) and resolve alternative patterns (e.g. Mr./Ms., " &
            "he/she, Herr/Frau) based on the data provided in the INSTRUCTION. A paragraph that contains no " &
            "placeholders and is not otherwise targeted by the INSTRUCTION must be returned unchanged." & vbCrLf &
            "10. When the INSTRUCTION requires ADDING content (e.g. listing multiple persons, adding paragraphs), " &
            "you MAY return multi-line text for a single [n] paragraph. Use actual line breaks between the lines — " &
            "each line will become a separate paragraph in the output document."

        Dim translationDictionary As String = Global.SharedLibrary.SharedLibrary.SharedMethods.GetTranslationDictionaryText(_context)
        If Not String.IsNullOrWhiteSpace(translationDictionary) Then
            systemPrompt &= vbCrLf & vbCrLf &
                "If the INSTRUCTION requires translation, apply the following translation dictionaries: " &
                translationDictionary
        End If

        Dim ruleNum As Integer = 11

        systemPrompt &= vbCrLf &
            $"{ruleNum}. Only modify text that is directly relevant to the INSTRUCTION. Leave all other text " &
            "EXACTLY as-is, character for character. When filling a template or completing a document, preserve " &
            "all boilerplate, headings, and structural text verbatim. Only replace placeholders (brackets, " &
            "underscores, ellipsis …, guillemets, TBD, etc.) and resolve alternative patterns (e.g. Mr./Ms. → Mr., " &
            "Herr/Frau → Herr) based on the data in the instruction."
        ruleNum += 1

        systemPrompt &= vbCrLf &
            $"{ruleNum}. When the INSTRUCTION requires ADDING content (e.g. listing multiple persons, " &
            "adding paragraphs), you MAY return multi-line text for a single [n] paragraph. Use actual " &
            "line breaks between the lines — each line will become a separate paragraph in the output document."
        ruleNum += 1

        If hasMarkers Then
            systemPrompt &= vbCrLf &
            $"{ruleNum}. Some paragraphs contain pipe characters (|) that mark formatting boundaries. " &
            "IMPORTANT: Preserve these | markers in your output at approximately the same positions " &
            "relative to the text. The number of | markers in each paragraph must stay EXACTLY the same. " &
            "Do NOT add or remove any | markers."
            ruleNum += 1
        End If

        Dim hasNoteRefMarkers As Boolean = processable.Any(Function(p) p.FullText.Contains(AP_NoteRefMarker))
        If hasNoteRefMarkers Then
            systemPrompt &= vbCrLf &
            $"{ruleNum}. Some paragraphs contain the character ‖ (double vertical line). " &
            "This marks a protected structural boundary such as a footnote/endnote reference, field boundary, or manual Word line break. " &
            "CRITICAL: Keep each ‖ at EXACTLY the same position relative to the surrounding words. " &
            "Do NOT move, add, or remove any ‖ characters."
            ruleNum += 1
        End If

        If Not String.IsNullOrWhiteSpace(contentScope) Then
            systemPrompt &= vbCrLf &
            $"{ruleNum}. Section context: The text to process comes from the document's {contentScope}. " &
            "Use this only as context for tone, conventions, and typical phrasing. " &
            "Do not add labels, headings, or explanations, and do not mention the section unless the source text already does."
            ruleNum += 1
        End If

        Dim batchIndex As Integer = 0
        Dim totalBatches As Integer = CInt(Math.Ceiling(processable.Count / CDbl(AP_ParagraphsPerBatch)))
        Dim currentBatch As Integer = 0

        ApDashboardLog($"DocProcessor: {processable.Count} paragraphs to process (~{totalBatches} batches)", "step")

        While batchIndex < processable.Count
            ct.ThrowIfCancellationRequested()

            currentBatch += 1

            Dim batchStart = batchIndex
            Dim batchEnd = Math.Min(batchIndex + AP_ParagraphsPerBatch - 1, processable.Count - 1)

            ' Adjust for character limit
            Dim batchChars As Integer = 0
            For j = batchStart To batchEnd
                batchChars += processable(j).FullText.Length
                If batchChars > AP_MaxCharsPerBatch AndAlso j > batchStart Then
                    batchEnd = j - 1
                    Exit For
                End If
            Next

            ApDashboardLog($"DocProcessor: batch {currentBatch}/{totalBatches} (paragraphs {batchStart + 1}-{batchEnd + 1})", "step")

            ' Build prompt
            Dim promptBuilder As New StringBuilder()

            ' Context Before
            Dim contextBeforeStart = Math.Max(0, batchStart - AP_ContextBefore)
            If contextBeforeStart < batchStart Then
                promptBuilder.AppendLine("[CONTEXT BEFORE - for reference only]")
                For j = contextBeforeStart To batchStart - 1
                    Dim contextText = If(processable(j).TranslatedText, processable(j).FullText)
                    ' Strip markers from context to avoid LLM echoing them into non-marker paragraphs
                    If contextText IsNot Nothing Then
                        contextText = contextText.Replace(AP_RunBoundaryMarker, "")
                        contextText = contextText.Replace(AP_NoteRefMarker, "")
                    End If
                    promptBuilder.AppendLine(contextText)
                Next
                promptBuilder.AppendLine()
            End If

            ' Paragraphs to process — use marker text when available
            promptBuilder.AppendLine("[TEXTTOPROCESS]")
            Dim batchNumber = 1
            For j = batchStart To batchEnd
                Dim paraText = If(processable(j).MarkerText, processable(j).FullText)
                promptBuilder.AppendLine("[" & batchNumber.ToString() & "] " & paraText)
                batchNumber += 1
            Next
            promptBuilder.AppendLine("[/TEXTTOPROCESS]")
            promptBuilder.AppendLine()

            ' Context After
            Dim contextAfterEnd = Math.Min(processable.Count - 1, batchEnd + AP_ContextAfter)
            If contextAfterEnd > batchEnd Then
                promptBuilder.AppendLine("[CONTEXT AFTER - for reference only]")
                For j = batchEnd + 1 To contextAfterEnd
                    Dim ctxAfterText As String = processable(j).FullText
                    If ctxAfterText IsNot Nothing Then
                        ctxAfterText = ctxAfterText.Replace(AP_NoteRefMarker, "")
                    End If
                    promptBuilder.AppendLine(ctxAfterText)
                Next
            End If

            ' Call LLM
            Dim llmResponse = Await LLM(systemPrompt, promptBuilder.ToString(),
                                         UseSecondAPI:=_apUseSecondApi,
                                         HideSplash:=True, EnsureUI:=False,
                                         cancellationToken:=ct)

            If String.IsNullOrWhiteSpace(llmResponse) Then
                ApDashboardLog($"DocProcessor: batch {currentBatch} returned empty response", "warn")
                Return False
            End If

            ' Parse response
            APParseResponse(llmResponse, processable, batchStart, batchEnd)

            ' ─── Restore non-breaking spaces in processed paragraphs ───
            For j As Integer = batchStart To batchEnd
                If processable(j).TranslatedText IsNot Nothing AndAlso nbspRecords.ContainsKey(j) Then
                    processable(j).TranslatedText = APRestoreNonBreakingSpaces(
                        processable(j).TranslatedText, nbspRecords(j))
                End If
            Next

            batchIndex = batchEnd + 1
        End While

        ApDashboardLog($"DocProcessor: all {currentBatch} batches completed", "success")
        Return True
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  RESPONSE PARSING
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Parses LLM output and assigns processed text to paragraph entries.
    ''' </summary>
    Private Sub APParseResponse(response As String, paragraphs As List(Of APParagraphInfo),
                                 batchStart As Integer, batchEnd As Integer)
        Dim pattern As New Regex("\[(\d+)\]\s*(.*?)(?=\s*\[\d+\]|$)", RegexOptions.Singleline)
        Dim matches = pattern.Matches(response)

        For Each m As Match In matches
            Dim num As Integer
            If Integer.TryParse(m.Groups(1).Value, num) Then
                Dim absoluteIndex = batchStart + num - 1
                If absoluteIndex >= batchStart AndAlso absoluteIndex <= batchEnd AndAlso absoluteIndex < paragraphs.Count Then
                    Dim processed = m.Groups(2).Value.Trim()
                    processed = Regex.Replace(processed, "^\[/?TEXTTOPROCESS\]", "", RegexOptions.IgnoreCase).Trim()
                    processed = Regex.Replace(processed, "\[/?TEXTTOPROCESS\]$", "", RegexOptions.IgnoreCase).Trim()
                    paragraphs(absoluteIndex).TranslatedText = processed
                End If
            End If
        Next
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  APPLYING TRANSLATIONS BACK TO XML
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Determines whether the processed paragraph text is visibly unchanged.
    ''' If unchanged, the original XML runs are left untouched so run-level formatting
    ''' and metadata remain exactly as they were.
    ''' </summary>
    ''' <param name="para">The paragraph being processed.</param>
    ''' <param name="translatedText">The LLM-processed paragraph text.</param>
    ''' <returns><c>True</c> if the visible paragraph text did not change.</returns>
    Private Function APIsParagraphTextVisiblyUnchanged(para As APParagraphInfo, translatedText As String) As Boolean
        If para Is Nothing OrElse translatedText Is Nothing Then Return False

        Dim originalVisibleText As String = If(para.FullText, "")
        Dim processedVisibleText As String = translatedText

        If para.TextRuns IsNot Nothing AndAlso para.TextRuns.Any(Function(r) r.HasNoteReferenceBefore) Then
            originalVisibleText = originalVisibleText.Replace(AP_NoteRefMarker, "")
            processedVisibleText = processedVisibleText.Replace(AP_NoteRefMarker, "")
        End If

        ' Only strip formatting markers when this paragraph was actually sent in marker mode
        ' and the original visible text itself does not already contain pipe characters.
        If para.MarkerText IsNot Nothing AndAlso Not originalVisibleText.Contains(AP_RunBoundaryMarker) Then
            processedVisibleText = processedVisibleText.Replace(AP_RunBoundaryMarker, "")
        End If

        Return String.Equals(originalVisibleText, processedVisibleText, StringComparison.Ordinal)
    End Function

    ''' <summary>
    ''' Applies translated text back into the XML text nodes.
    ''' Partitions runs at protected structural boundaries so that text
    ''' redistribution never moves content across a reference anchor.
    ''' When the LLM preserves ‖ markers, splits directly at their position (they are
    ''' authoritative). Falls back to anchor + word-boundary logic only when ‖ markers
    ''' are missing. The ‖ character is purely a positional placeholder — it indicates
    ''' where a footnote/endnote/field reference or manual line break sits but does not create an additional
    ''' formatting split beyond the existing run partition.
    ''' </summary>
    Private Sub APApplyTranslations(paragraphs As List(Of APParagraphInfo))
        For Each para In paragraphs
            If para.IsEmpty OrElse String.IsNullOrEmpty(para.TranslatedText) Then Continue For
            If para.TextRuns.Count = 0 Then Continue For

            Dim translatedText = para.TranslatedText

            ' If the visible text did not change, do not touch any text nodes.
            ' This prevents formatting drift caused by redistributing text across runs.
            If APIsParagraphTextVisiblyUnchanged(para, translatedText) Then Continue For

            ' Single run: simple replacement
            If para.TextRuns.Count = 1 Then
                translatedText = translatedText.Replace(AP_RunBoundaryMarker, "").Replace(AP_NoteRefMarker, "")
                APSetTextNode(para.TextRuns(0).TextNode, translatedText)
                Continue For
            End If

            ' ─── Check if any protected structural boundaries exist ───
            Dim hasNoteRefBoundaries As Boolean = para.TextRuns.Any(Function(r) r.HasNoteReferenceBefore)

            ' Fast path for paragraphs without note/field/manual-break anchors. Paragraphs with anchors
            ' are handled below segment-by-segment, where formatting markers can also be used safely.
            If Not hasNoteRefBoundaries Then
                If APTryApplyMarkerBasedDistribution(para, translatedText) Then
                    Continue For
                End If
                translatedText = translatedText.Replace(AP_RunBoundaryMarker, "")
                translatedText = translatedText.Replace(AP_NoteRefMarker, "")
                APDistributeProportional(para, translatedText)
                Continue For
            End If

            ' ─── From here: paragraph HAS protected structural boundaries ───
            ' Keep formatting markers until each note-delimited segment is applied.

            ' Partition runs into segments at protected structural boundaries
            Dim segments As New List(Of List(Of Integer))()
            Dim currentSegment As New List(Of Integer)()

            For runIdx As Integer = 0 To para.TextRuns.Count - 1
                If para.TextRuns(runIdx).HasNoteReferenceBefore AndAlso currentSegment.Count > 0 Then
                    segments.Add(currentSegment)
                    currentSegment = New List(Of Integer)()
                End If
                currentSegment.Add(runIdx)
            Next
            If currentSegment.Count > 0 Then segments.Add(currentSegment)

            ' Safety: if partitioning produced only 1 segment, fall back to proportional
            If segments.Count <= 1 Then
                translatedText = translatedText.Replace(AP_RunBoundaryMarker, "")
                translatedText = translatedText.Replace(AP_NoteRefMarker, "")
                APDistributeProportional(para, translatedText)
                Continue For
            End If

            ' ─── Split translated text across segments ───
            Dim expectedNoteRefCount As Integer = segments.Count - 1
            Dim actualNoteRefCount As Integer = 0
            For Each ch As Char In translatedText
                If ch = AP_NoteRefMarker(0) Then actualNoteRefCount += 1
            Next

            Dim segmentTexts As String() = Nothing

            If actualNoteRefCount = expectedNoteRefCount Then
                ' ─── LLM preserved ‖ markers: split directly on them ───
                ' The ‖ is a positional placeholder — it tells us where the footnote
                ' reference sits. We split the text at these positions and assign
                ' each piece to its corresponding segment of runs.
                segmentTexts = translatedText.Split(New String() {AP_NoteRefMarker}, StringSplitOptions.None)

                If segmentTexts.Length = segments.Count Then
                    ' Normalize spaces at segment boundaries: if the original text after
                    ' the footnote started with whitespace, ensure the space is on that
                    ' side (not trailing on the previous segment before the superscript).
                    For segIdx As Integer = 0 To segmentTexts.Length - 2
                        Dim nextSegFirstRunIdx As Integer = segments(segIdx + 1)(0)
                        Dim nextRunOriginal As String = para.TextRuns(nextSegFirstRunIdx).OriginalText
                        If nextRunOriginal.Length > 0 AndAlso Char.IsWhiteSpace(nextRunOriginal(0)) Then
                            Dim curSeg As String = segmentTexts(segIdx)
                            If curSeg.Length > 0 AndAlso curSeg(curSeg.Length - 1) = " "c Then
                                Dim trimmed As String = curSeg.TrimEnd(" "c)
                                Dim movedSpaces As String = curSeg.Substring(trimmed.Length)
                                segmentTexts(segIdx) = trimmed
                                segmentTexts(segIdx + 1) = movedSpaces & segmentTexts(segIdx + 1)
                            End If
                        End If
                    Next
                Else
                    segmentTexts = Nothing
                End If
            End If

            If segmentTexts Is Nothing Then
                ' ─── Fallback: anchor-based splitting (no ‖ or wrong count) ───
                Debug.WriteLine($"Note-ref marker mismatch for paragraph {para.Index}: expected {expectedNoteRefCount}, got {actualNoteRefCount}. Using anchor fallback.")
                translatedText = translatedText.Replace(AP_RunBoundaryMarker, "")
                translatedText = translatedText.Replace(AP_NoteRefMarker, "")

                Dim totalOrigLen As Integer = para.FullText.Replace(AP_NoteRefMarker, "").Length
                Dim translatedLen As Integer = translatedText.Length

                If totalOrigLen = 0 Then
                    APSetTextNode(para.TextRuns(0).TextNode, translatedText)
                    For idx As Integer = 1 To para.TextRuns.Count - 1
                        APSetTextNode(para.TextRuns(idx).TextNode, "")
                    Next
                    Continue For
                End If

                Dim segmentOrigTexts As New List(Of String)()
                For Each seg In segments
                    Dim sb As New StringBuilder()
                    For Each ri In seg
                        sb.Append(para.TextRuns(ri).OriginalText)
                    Next
                    segmentOrigTexts.Add(sb.ToString())
                Next

                Dim splitPoints As New List(Of Integer)()
                Dim searchFrom As Integer = 0

                For segIdx As Integer = 0 To segments.Count - 2
                    Dim segOrigText As String = segmentOrigTexts(segIdx)
                    Dim nextSegOrigText As String = segmentOrigTexts(segIdx + 1)
                    Dim cumulativeOrigLen As Integer = 0
                    For si As Integer = 0 To segIdx
                        cumulativeOrigLen += segmentOrigTexts(si).Length
                    Next
                    Dim proportion As Double = cumulativeOrigLen / CDbl(totalOrigLen)
                    Dim proportionalEnd As Integer = CInt(Math.Round(proportion * translatedLen))
                    proportionalEnd = Math.Min(proportionalEnd, translatedLen)

                    Dim splitPos As Integer = APFindNoteRefSplitPos(
                        translatedText, segOrigText, searchFrom, proportionalEnd, translatedLen, nextSegOrigText)
                    splitPos = Math.Min(splitPos, translatedLen)

                    If splitPos > searchFrom AndAlso splitPos <= translatedLen Then
                        Dim nextSegFirstRunIdx As Integer = segments(segIdx + 1)(0)
                        Dim nextRunOriginal As String = para.TextRuns(nextSegFirstRunIdx).OriginalText
                        If nextRunOriginal.Length > 0 AndAlso Char.IsWhiteSpace(nextRunOriginal(0)) Then
                            While splitPos > searchFrom AndAlso translatedText(splitPos - 1) = " "c
                                splitPos -= 1
                            End While
                        End If
                    End If

                    splitPoints.Add(splitPos)
                    searchFrom = splitPos
                Next
                splitPoints.Add(translatedLen)

                segmentTexts = New String(segments.Count - 1) {}
                Dim st As Integer = 0
                For i As Integer = 0 To segments.Count - 1
                    Dim en As Integer = splitPoints(i)
                    segmentTexts(i) = If(en > st, translatedText.Substring(st, en - st), "")
                    st = en
                Next
            End If

            ' ─── Distribute text within each segment independently ───
            ' Each segment is a group of runs between two footnote references.
            ' Within each segment, text is distributed proportionally across runs.
            For segIdx As Integer = 0 To segments.Count - 1
                Dim segText As String = segmentTexts(segIdx)
                Dim segRunIndices As List(Of Integer) = segments(segIdx)

                If APTryApplyMarkerBasedDistributionToRuns(para, segRunIndices, segText) Then
                    Continue For
                End If
                segText = segText.Replace(AP_RunBoundaryMarker, "")

                If segRunIndices.Count = 1 Then
                    APSetTextNode(para.TextRuns(segRunIndices(0)).TextNode, segText)
                Else
                    Dim segOrigLen As Integer = 0
                    For Each ri In segRunIndices
                        segOrigLen += para.TextRuns(ri).OriginalText.Length
                    Next
                    Dim segTransLen As Integer = segText.Length
                    Dim segCurrentPos As Integer = 0
                    Dim segCumOrig As Integer = 0

                    For ri As Integer = 0 To segRunIndices.Count - 1
                        Dim runIdx As Integer = segRunIndices(ri)
                        Dim run = para.TextRuns(runIdx)
                        segCumOrig += run.OriginalText.Length

                        If ri = segRunIndices.Count - 1 Then
                            Dim remaining As String = If(segCurrentPos < segTransLen,
                                                         segText.Substring(segCurrentPos), "")
                            APSetTextNode(run.TextNode, remaining)
                        Else
                            If segOrigLen = 0 Then
                                APSetTextNode(run.TextNode, "")
                                Continue For
                            End If

                            Dim prop As Double = segCumOrig / CDbl(segOrigLen)
                            Dim tgtEnd As Integer = CInt(Math.Round(prop * segTransLen))
                            tgtEnd = Math.Min(tgtEnd, segTransLen)

                            If tgtEnd <= segCurrentPos Then
                                APSetTextNode(run.TextNode, "")
                                Continue For
                            End If

                            Dim endPos As Integer = tgtEnd
                            Dim foundSpace As Boolean = False

                            If endPos < segTransLen AndAlso endPos > segCurrentPos Then
                                If segText(endPos) = " "c Then
                                    endPos += 1
                                    foundSpace = True
                                Else
                                    Dim searchMax As Integer = Math.Min(endPos + 15, segTransLen - 1)
                                    For sp As Integer = endPos To searchMax
                                        If segText(sp) = " "c Then
                                            endPos = sp + 1
                                            foundSpace = True
                                            Exit For
                                        End If
                                    Next
                                    If Not foundSpace Then
                                        For sp As Integer = endPos - 1 To segCurrentPos Step -1
                                            If segText(sp) = " "c Then
                                                endPos = sp + 1
                                                foundSpace = True
                                                Exit For
                                            End If
                                        Next
                                    End If
                                    If Not foundSpace Then
                                        Dim we As Integer = endPos
                                        While we < segTransLen AndAlso segText(we) <> " "c
                                            we += 1
                                        End While
                                        endPos = we
                                    End If
                                End If
                            End If

                            endPos = Math.Min(endPos, segTransLen)
                            APSetTextNode(run.TextNode, segText.Substring(segCurrentPos, endPos - segCurrentPos))
                            segCurrentPos = endPos
                        End If
                    Next
                End If
            Next

        Next
    End Sub

    ''' <summary>
    ''' Proportional distribution fallback when no note-reference boundaries exist.
    ''' </summary>
    Private Sub APDistributeProportional(para As APParagraphInfo, translatedText As String)
        Dim totalOrigLen As Integer = para.FullText.Length
        Dim translatedLen As Integer = translatedText.Length

        If totalOrigLen = 0 Then
            APSetTextNode(para.TextRuns(0).TextNode, translatedText)
            For idx As Integer = 1 To para.TextRuns.Count - 1
                APSetTextNode(para.TextRuns(idx).TextNode, "")
            Next
            Return
        End If

        Dim currentPos As Integer = 0
        Dim cumulativeOriginal As Integer = 0

        For runIdx As Integer = 0 To para.TextRuns.Count - 1
            Dim run = para.TextRuns(runIdx)
            cumulativeOriginal += run.OriginalText.Length

            If runIdx = para.TextRuns.Count - 1 Then
                Dim remaining = If(currentPos < translatedLen, translatedText.Substring(currentPos), "")
                APSetTextNode(run.TextNode, remaining)
            Else
                Dim proportion As Double = cumulativeOriginal / CDbl(totalOrigLen)
                Dim targetEndPos = CInt(Math.Round(proportion * translatedLen))
                targetEndPos = Math.Min(targetEndPos, translatedLen)

                If targetEndPos <= currentPos Then
                    APSetTextNode(run.TextNode, "")
                    Continue For
                End If

                Dim endPos = targetEndPos
                Dim foundSpace = False

                If endPos < translatedLen AndAlso endPos > currentPos Then
                    If translatedText(endPos) = " "c Then
                        endPos += 1
                        foundSpace = True
                    Else
                        Dim searchMax As Integer = Math.Min(endPos + 15, translatedLen - 1)
                        For searchPos = endPos To searchMax
                            If translatedText(searchPos) = " "c Then
                                endPos = searchPos + 1
                                foundSpace = True
                                Exit For
                            End If
                        Next
                        If Not foundSpace Then
                            For searchPos = endPos - 1 To currentPos Step -1
                                If translatedText(searchPos) = " "c Then
                                    endPos = searchPos + 1
                                    foundSpace = True
                                    Exit For
                                End If
                            Next
                        End If
                        If Not foundSpace Then
                            Dim wordEnd = endPos
                            While wordEnd < translatedLen AndAlso translatedText(wordEnd) <> " "c
                                wordEnd += 1
                            End While
                            If wordEnd < translatedLen AndAlso translatedText(wordEnd) = " "c Then
                                endPos = wordEnd + 1
                            Else
                                endPos = wordEnd
                            End If
                        End If
                    End If
                End If

                endPos = Math.Min(endPos, translatedLen)
                APSetTextNode(run.TextNode, translatedText.Substring(currentPos, endPos - currentPos))
                currentPos = endPos
            End If
        Next
    End Sub

    ''' <summary>
    ''' Finds the best split position in the translated text for a note-reference boundary
    ''' by anchoring on the last word(s) of the original segment text, or the first word(s)
    ''' of the next segment text. Falls back to proportional position with backward
    ''' word-boundary seeking if both anchoring strategies fail.
    ''' </summary>
    ''' <param name="translatedText">The full translated paragraph text.</param>
    ''' <param name="originalSegText">The original text of the segment before the note reference.</param>
    ''' <param name="currentPos">Start of the current unassigned region in translatedText.</param>
    ''' <param name="targetEndPos">Proportional end position (fallback).</param>
    ''' <param name="translatedLength">Total length of translatedText.</param>
    ''' <param name="nextSegOrigText">The original text of the segment after the note reference (Nothing if unavailable).</param>
    ''' <returns>The character index in translatedText where the split should occur.</returns>
    Private Shared Function APFindNoteRefSplitPos(
            translatedText As String,
            originalSegText As String,
            currentPos As Integer,
            targetEndPos As Integer,
            translatedLength As Integer,
            Optional nextSegOrigText As String = Nothing) As Integer

        ' ── Strategy 1: Anchor on the LAST word(s) of the current segment ──
        If Not String.IsNullOrWhiteSpace(originalSegText) AndAlso originalSegText.Length >= 2 Then
            Dim origTrimmed As String = originalSegText.TrimEnd()
            If origTrimmed.Length >= 2 Then
                Dim words As String() = origTrimmed.Split(" "c)
                Dim maxWords As Integer = Math.Min(words.Length, 3)

                For tryWords As Integer = maxWords To 1 Step -1
                    Dim fragment As String = String.Join(" ", words, words.Length - tryWords, tryWords)
                    If fragment.Length < 2 Then Continue For

                    Dim searchRegion As String = translatedText.Substring(currentPos)
                    Dim fragIdx As Integer = searchRegion.IndexOf(fragment, StringComparison.OrdinalIgnoreCase)

                    If fragIdx >= 0 Then
                        Dim splitPos As Integer = currentPos + fragIdx + fragment.Length
                        If splitPos >= translatedLength OrElse
                           translatedText(splitPos) = " "c OrElse
                           Char.IsPunctuation(translatedText(splitPos)) Then
                            Return splitPos
                        End If
                    End If
                Next
            End If
        End If

        ' ── Strategy 2: Anchor on the FIRST word(s) of the NEXT segment ──
        If Not String.IsNullOrWhiteSpace(nextSegOrigText) AndAlso nextSegOrigText.Length >= 2 Then
            Dim nextTrimmed As String = nextSegOrigText.TrimStart()
            If nextTrimmed.Length >= 2 Then
                Dim nextWords As String() = nextTrimmed.Split(" "c)
                Dim maxNextWords As Integer = Math.Min(nextWords.Length, 3)

                For tryWords As Integer = maxNextWords To 1 Step -1
                    Dim fragment As String = String.Join(" ", nextWords, 0, tryWords)
                    If fragment.Length < 2 Then Continue For

                    Dim earliestSearch As Integer = Math.Max(currentPos, CInt(currentPos + (targetEndPos - currentPos) * 0.4))
                    If earliestSearch >= translatedLength Then Continue For

                    Dim searchRegion As String = translatedText.Substring(earliestSearch)
                    Dim fragIdx As Integer = searchRegion.IndexOf(fragment, StringComparison.OrdinalIgnoreCase)

                    If fragIdx >= 0 Then
                        Dim matchStart As Integer = earliestSearch + fragIdx
                        Dim splitPos As Integer = matchStart
                        While splitPos > currentPos AndAlso translatedText(splitPos - 1) = " "c
                            splitPos -= 1
                        End While
                        If splitPos <= currentPos Then Continue For
                        If splitPos >= translatedLength OrElse
                           Char.IsWhiteSpace(translatedText(splitPos)) OrElse
                           (splitPos > 0 AndAlso (Char.IsPunctuation(translatedText(splitPos - 1)) OrElse
                                                   Char.IsWhiteSpace(translatedText(splitPos - 1)))) Then
                            Return splitPos
                        End If
                    End If
                Next
            End If
        End If

        ' ── Fallback: backward word-boundary search from proportional position ──
        Dim endPos As Integer = Math.Min(targetEndPos, translatedLength)
        If endPos > currentPos AndAlso endPos < translatedLength Then
            If translatedText(endPos) = " "c Then Return endPos
            For searchPos As Integer = endPos - 1 To currentPos Step -1
                If translatedText(searchPos) = " "c Then
                    Return searchPos + 1
                End If
            Next
        End If

        Return endPos
    End Function

    ''' <summary>
    ''' Attempts marker-based distribution (matching Word add-in's TryApplyMarkerBasedDistribution).
    ''' If the LLM returned text with exactly the right number of | markers, split on them
    ''' and assign each segment to the corresponding run. Returns True if successful.
    ''' </summary>
    Private Function APTryApplyMarkerBasedDistribution(para As APParagraphInfo, translatedText As String) As Boolean
        If para Is Nothing OrElse para.TextRuns Is Nothing Then Return False

        Dim runIndices As New List(Of Integer)()
        For i As Integer = 0 To para.TextRuns.Count - 1
            runIndices.Add(i)
        Next

        Return APTryApplyMarkerBasedDistributionToRuns(para, runIndices, translatedText)
    End Function

    Private Function APTryApplyMarkerBasedDistributionToRuns(
        para As APParagraphInfo,
        runIndices As List(Of Integer),
        translatedText As String) As Boolean

        If para Is Nothing OrElse para.TextRuns Is Nothing OrElse runIndices Is Nothing Then Return False

        Dim groups As List(Of List(Of Integer)) = APBuildFormattingGroups(para.TextRuns, runIndices)
        If groups.Count = 0 Then Return False

        Dim expectedMarkers As Integer = groups.Count - 1
        Dim actualMarkers As Integer = translatedText.Count(Function(c) c = AP_RunBoundaryMarker(0))
        If actualMarkers <> expectedMarkers Then
            Debug.WriteLine($"Formatting marker mismatch for paragraph {para.Index}: expected {expectedMarkers}, got {actualMarkers}. Falling back.")
            Return False
        End If

        Dim segments As String() = translatedText.Split(New String() {AP_RunBoundaryMarker}, StringSplitOptions.None)
        If segments.Length <> groups.Count Then Return False

        For groupIdx As Integer = 0 To groups.Count - 1
            APDistributeTextAcrossRuns(para, groups(groupIdx), segments(groupIdx))
        Next

        Return True
    End Function

    Private Sub APDistributeTextAcrossRuns(
        para As APParagraphInfo,
        runIndices As List(Of Integer),
        text As String)

        If runIndices Is Nothing OrElse runIndices.Count = 0 Then Return

        If runIndices.Count = 1 Then
            APSetTextNode(para.TextRuns(runIndices(0)).TextNode, text)
            Return
        End If

        Dim totalOrigLen As Integer = 0
        For Each runIdx As Integer In runIndices
            totalOrigLen += para.TextRuns(runIdx).OriginalText.Length
        Next

        If totalOrigLen <= 0 Then
            APSetTextNode(para.TextRuns(runIndices(0)).TextNode, text)
            For i As Integer = 1 To runIndices.Count - 1
                APSetTextNode(para.TextRuns(runIndices(i)).TextNode, "")
            Next
            Return
        End If

        Dim translatedLen As Integer = text.Length
        Dim currentPos As Integer = 0
        Dim cumulativeOriginal As Integer = 0

        For groupRunPos As Integer = 0 To runIndices.Count - 1
            Dim runIdx As Integer = runIndices(groupRunPos)
            Dim run As APTextRunInfo = para.TextRuns(runIdx)
            cumulativeOriginal += run.OriginalText.Length

            If groupRunPos = runIndices.Count - 1 Then
                Dim remaining As String = If(currentPos < translatedLen, text.Substring(currentPos), "")
                APSetTextNode(run.TextNode, remaining)
                Continue For
            End If

            Dim targetEndPos As Integer = CInt(Math.Round((cumulativeOriginal / CDbl(totalOrigLen)) * translatedLen))
            targetEndPos = Math.Min(targetEndPos, translatedLen)

            If targetEndPos <= currentPos Then
                APSetTextNode(run.TextNode, "")
                Continue For
            End If

            Dim endPos As Integer = targetEndPos
            Dim foundSpace As Boolean = False

            If endPos < translatedLen AndAlso endPos > currentPos Then
                If text(endPos) = " "c Then
                    endPos += 1
                    foundSpace = True
                Else
                    Dim searchMax As Integer = Math.Min(endPos + 15, translatedLen - 1)
                    For searchPos As Integer = endPos To searchMax
                        If text(searchPos) = " "c Then
                            endPos = searchPos + 1
                            foundSpace = True
                            Exit For
                        End If
                    Next

                    If Not foundSpace Then
                        For searchPos As Integer = endPos - 1 To currentPos Step -1
                            If text(searchPos) = " "c Then
                                endPos = searchPos + 1
                                foundSpace = True
                                Exit For
                            End If
                        Next
                    End If

                    If Not foundSpace Then
                        Dim wordEnd As Integer = endPos
                        While wordEnd < translatedLen AndAlso text(wordEnd) <> " "c
                            wordEnd += 1
                        End While
                        endPos = wordEnd
                    End If
                End If
            End If

            endPos = Math.Min(endPos, translatedLen)
            APSetTextNode(run.TextNode, text.Substring(currentPos, endPos - currentPos))
            currentPos = endPos
        Next
    End Sub

    ''' <summary>
    ''' Sets a w:t text node and ensures xml:space="preserve" when needed.
    ''' </summary>
    Private Sub APSetTextNode(textNode As System.Xml.XmlNode, text As String)
        textNode.InnerText = text
        If text.Length > 0 AndAlso (text.StartsWith(" ") OrElse text.EndsWith(" ") OrElse text.Contains("  ")) Then
            Dim xmlSpaceAttr = textNode.Attributes("xml:space")
            If xmlSpaceAttr Is Nothing Then
                xmlSpaceAttr = textNode.OwnerDocument.CreateAttribute("xml", "space", "http://www.w3.org/XML/1998/namespace")
                textNode.Attributes.Append(xmlSpaceAttr)
            End If
            xmlSpaceAttr.Value = "preserve"
        End If
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  SUB-PART PROCESSING (headers, footers, comments, footnotes, endnotes)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Processes header, footer, comment, footnote, and endnote parts in a DOCX package.
    ''' </summary>
    Private Async Function APProcessSubParts(tempDir As String, instruction As String, ct As CancellationToken) As Task
        Dim wordDir = Path.Combine(tempDir, "word")
        If Not Directory.Exists(wordDir) Then Return

        ' Process headers and footers
        For Each pattern In {"header*.xml", "footer*.xml"}
            For Each filePath In Directory.GetFiles(wordDir, pattern)
                Try
                    Dim isHeader As Boolean = Path.GetFileName(filePath).StartsWith("header", StringComparison.OrdinalIgnoreCase)
                    Dim contentScope As String = If(isHeader, "header", "footer")
                    Await APProcessXmlFile(filePath, instruction, ct, contentScope)
                Catch ex As System.Exception
                    Debug.WriteLine("APProcessSubParts error for " & Path.GetFileName(filePath) & ": " & ex.Message)
                End Try
            Next
        Next

        ' Process comments, footnotes, endnotes
        For Each fileName In {"comments.xml", "footnotes.xml", "endnotes.xml"}
            Dim filePath = Path.Combine(wordDir, fileName)
            If File.Exists(filePath) Then
                Try
                    Dim contentScope As String = ""
                    Select Case fileName.ToLowerInvariant()
                        Case "comments.xml"
                            contentScope = "comments"
                        Case "footnotes.xml"
                            contentScope = "footnotes"
                        Case "endnotes.xml"
                            contentScope = "endnotes"
                    End Select

                    Await APProcessXmlFile(filePath, instruction, ct, contentScope)
                Catch ex As System.Exception
                    Debug.WriteLine("APProcessSubParts error for " & fileName & ": " & ex.Message)
                End Try
            End If
        Next
    End Function

    ''' <summary>Processes a single XML file (header/footer/comments/etc.).</summary>
    Private Async Function APProcessXmlFile(filePath As String, instruction As String, ct As CancellationToken,
                                        Optional contentScope As String = "") As Task
        Dim xmlDoc As New System.Xml.XmlDocument()
        xmlDoc.PreserveWhitespace = True
        xmlDoc.Load(filePath)

        Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
        nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

        Dim paragraphs = APExtractParagraphs(xmlDoc, nsMgr)
        Dim processable = paragraphs.Where(Function(p) Not p.IsEmpty).ToList()
        If processable.Count = 0 Then Return

        Dim success = Await APProcessBatches(paragraphs, instruction, ct, contentScope)
        If success Then
            APApplyTranslations(paragraphs)
            APExpandMultiLineParagraphs(paragraphs, nsMgr)
            xmlDoc.Save(filePath)
        End If
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  ENTRY POINT: ProcessDocumentForAutoPilot
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Routes document processing to the appropriate format handler.
    ''' Applies format-specific alternate models (if defined) for simple tasks
    ''' (translation, correction) where a faster model suffices, but skips it
    ''' for complex tasks that benefit from the more capable primary model.
    ''' </summary>
    ''' <param name="sheetFilter">Optional: restrict Excel processing to these sheet names.</param>
    ''' <param name="useOfflineDocs">True for simple tasks (translate/correct) where OfflineDocs should be used;
    ''' False for complex tasks (anonymization, restructuring, etc.) that need the primary model.</param>
    Private Async Function ProcessDocumentForAutoPilot(inputPath As String, outputPath As String,
                                                       instruction As String, ct As CancellationToken,
                                                       Optional sheetFilter As List(Of String) = Nothing,
                                                       Optional useOfflineDocs As Boolean = True) As Task(Of Boolean)

        ' ── Parse and strip task_type tag from instruction ──
        Dim taskTypeMatch = Regex.Match(instruction, "\s*\(task_type\s*=\s*(\w+)\)\s*$", RegexOptions.IgnoreCase)
        If taskTypeMatch.Success Then
            Dim taskType = taskTypeMatch.Groups(1).Value.ToLowerInvariant()
            instruction = instruction.Substring(0, taskTypeMatch.Index).TrimEnd()
            ' Override useOfflineDocs based on task_type
            Select Case taskType
                Case "translate", "correct"
                    useOfflineDocs = True
                Case Else
                    useOfflineDocs = False
            End Select
        End If

        ' ── Determine format-specific SpecialTaskModel key ──
        Dim ext = Path.GetExtension(inputPath).ToLowerInvariant()
        Dim specialTaskKey As String = "OfflineDocs"
        Select Case ext
            Case ".pptx"
                specialTaskKey = "OfflineDocsPptx"
            Case ".xlsx"
                specialTaskKey = "OfflineDocsXlsx"
        End Select

        ' ── Try format-specific alternate model for simple document processing tasks ──
        ' Only apply for translate/correct — complex tasks (anonymization, restructuring,
        ' etc.) should use the more capable primary/secondary model selected by the user.
        Dim offlineDocsBackup As ModelConfig = Nothing
        Dim offlineDocsApplied As Boolean = False
        Dim previousUseSecondApi As Boolean = _apUseSecondApi

        If useOfflineDocs AndAlso Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
            Try
                ' Capture the current config BEFORE GetSpecialTaskModel overwrites it
                offlineDocsBackup = GetCurrentConfig(_context)

                ' Try format-specific key first, fall back to generic OfflineDocs
                Dim modelFound As Boolean = False
                If Not specialTaskKey.Equals("OfflineDocs", StringComparison.Ordinal) Then
                    modelFound = GetSpecialTaskModel(_context, INI_AlternateModelPath, specialTaskKey)
                End If
                If Not modelFound Then
                    modelFound = GetSpecialTaskModel(_context, INI_AlternateModelPath, "OfflineDocs")
                End If

                If modelFound Then
                    offlineDocsApplied = True
                    _apUseSecondApi = True
                    ApDashboardLog($"DocProcessor: {specialTaskKey} model applied (task_type: translate/correct)", "step")
                Else
                    ' GetSpecialTaskModel didn't find an entry — discard backup
                    offlineDocsBackup = Nothing
                End If
            Catch
                offlineDocsBackup = Nothing
            End Try
        ElseIf Not useOfflineDocs Then
            ApDashboardLog("DocProcessor: using primary model (task_type: other)", "step")
        End If

        Try
            Select Case ext
                Case ".docx", ".doc"
                    Return Await ProcessDocxForAutoPilot(inputPath, outputPath, instruction, ct)
                Case ".pptx"
                    Return Await ProcessPptxForAutoPilot(inputPath, outputPath, instruction, ct)
                Case ".xlsx"
                    Return Await ProcessXlsxForAutoPilot(inputPath, outputPath, instruction, ct, sheetFilter)
                Case Else
                    Debug.WriteLine($"ProcessDocumentForAutoPilot: unsupported extension '{ext}'")
                    Return False
            End Select
        Finally
            If offlineDocsApplied Then
                _apUseSecondApi = previousUseSecondApi
                If offlineDocsBackup IsNot Nothing Then
                    RestoreDefaults(_context, offlineDocsBackup)
                End If
            End If
        End Try
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  PPTX PROCESSING
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Processes a PPTX file by applying the given instruction via LLM to all text paragraphs.
    ''' PPTX uses DrawingML namespace (a:) with a:p → a:r → a:t structure.
    ''' Uses a presentation-aware system prompt that provides slide-level context so the LLM
    ''' understands each slide's position in the deck. For global operations (translate,
    ''' correct, shorten, etc.) all slides are processed normally. The slide outline helps
    ''' the LLM maintain consistency across the presentation.
    ''' </summary>
    Private Async Function ProcessPptxForAutoPilot(inputPath As String, outputPath As String,
                                                    instruction As String, ct As CancellationToken) As Task(Of Boolean)
        Dim tempDir As String = Path.Combine(Path.GetTempPath(), AP_TempPrefix & "pptx_" & Guid.NewGuid().ToString("N"))

        Try
            ApDashboardLog("DocProcessor: starting PowerPoint processing", "step")

            ' Copy input to output (we modify the copy)
            File.Copy(inputPath, outputPath, overwrite:=True)

            ' Extract PPTX (ZIP)
            ZipFile.ExtractToDirectory(outputPath, tempDir)

            Dim pptDir As String = Path.Combine(tempDir, "ppt")
            If Not Directory.Exists(pptDir) Then
                ApDashboardLog("DocProcessor: invalid PPTX structure - ppt directory not found", "error")
                Return False
            End If

            ' === Build a full-presentation outline for context ===
            Dim slidesDir As String = Path.Combine(pptDir, "slides")
            Dim notesDir As String = Path.Combine(pptDir, "notesSlides")
            Dim slideCount As Integer = 0
            Dim outlineBuilder As New StringBuilder()

            If Directory.Exists(slidesDir) Then
                Dim slideFiles = Directory.GetFiles(slidesDir, "slide*.xml").
                    OrderBy(Function(f)
                                Dim m = Regex.Match(Path.GetFileNameWithoutExtension(f), "\d+")
                                Return If(m.Success, Integer.Parse(m.Value), 0)
                            End Function).ToArray()

                slideCount = slideFiles.Length

                ' Build outline: extract first 200 chars of each slide for context
                For Each slideFile In slideFiles
                    Dim slideNum As String = Regex.Match(Path.GetFileNameWithoutExtension(slideFile), "\d+").Value
                    Try
                        Dim previewDoc As New System.Xml.XmlDocument()
                        previewDoc.PreserveWhitespace = True
                        previewDoc.Load(slideFile)
                        Dim previewNsMgr As New System.Xml.XmlNamespaceManager(previewDoc.NameTable)
                        previewNsMgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
                        Dim allText As New StringBuilder()
                        For Each tNode As System.Xml.XmlNode In previewDoc.SelectNodes("//a:r/a:t", previewNsMgr)
                            allText.Append(tNode.InnerText)
                            If allText.Length > 200 Then Exit For
                        Next
                        Dim preview = allText.ToString()
                        If preview.Length > 200 Then preview = preview.Substring(0, 200) & "..."
                        outlineBuilder.AppendLine($"  Slide {slideNum}: {preview}")
                    Catch
                        outlineBuilder.AppendLine($"  Slide {slideNum}: (could not preview)")
                    End Try
                Next

                ApDashboardLog($"DocProcessor: found {slideCount} slide(s) to process", "step")

                For Each slideFile In slideFiles
                    ct.ThrowIfCancellationRequested()

                    Dim slideNum As String = Regex.Match(Path.GetFileNameWithoutExtension(slideFile), "\d+").Value
                    Dim slideLabel As String = If(slideNum.Length > 0, $"Slide {slideNum}", Path.GetFileNameWithoutExtension(slideFile))

                    ApDashboardLog($"DocProcessor: processing {slideLabel}", "step")
                    Await APProcessPptxXmlPart(slideFile, instruction, ct,
                                               slideCount, slideNum, outlineBuilder.ToString())
                Next
            End If

            ' === Process notes slides ===
            If Directory.Exists(notesDir) Then
                For Each notesFile In Directory.GetFiles(notesDir, "notesSlide*.xml")
                    ct.ThrowIfCancellationRequested()

                    Dim notesNum As String = Regex.Match(Path.GetFileNameWithoutExtension(notesFile), "\d+").Value
                    Dim notesLabel As String = If(notesNum.Length > 0, $"Notes {notesNum}", Path.GetFileNameWithoutExtension(notesFile))

                    ApDashboardLog($"DocProcessor: processing {notesLabel}", "step")
                    ' Notes slides: no slide-level context needed, use standard batching
                    Await APProcessPptxXmlPart(notesFile, instruction, ct)
                Next
            End If

            ' Repack PPTX
            File.Delete(outputPath)
            ZipFile.CreateFromDirectory(tempDir, outputPath, CompressionLevel.Optimal, False)

            ApDashboardLog($"DocProcessor: PowerPoint processing complete ({slideCount} slides)", "success")
            Return True

        Catch ex As OperationCanceledException
            Throw
        Catch ex As System.Exception
            Debug.WriteLine("ProcessPptxForAutoPilot error: " & ex.Message)
            ApDashboardLog("DocProcessor: error - " & ex.Message, "error")
            Return False
        Finally
            If Directory.Exists(tempDir) Then
                Try : Directory.Delete(tempDir, recursive:=True) : Catch : End Try
            End If
        End Try
    End Function

    ''' <summary>
    ''' Processes a single PPTX XML part (slide, notes slide) by extracting
    ''' DrawingML paragraphs, sending text to the LLM, and writing back.
    ''' When slide context is provided, uses a presentation-aware system prompt that
    ''' gives the LLM the full deck outline and current slide position. This helps
    ''' global operations (translate, correct) maintain cross-slide consistency, and
    ''' helps targeted operations (add/modify specific slides) avoid corrupting others.
    ''' </summary>
    Private Async Function APProcessPptxXmlPart(xmlPath As String, instruction As String, ct As CancellationToken,
                                                 Optional totalSlides As Integer = 0,
                                                 Optional currentSlideNum As String = Nothing,
                                                 Optional presentationOutline As String = Nothing) As Task(Of Boolean)
        Try
            Dim xmlDoc As New System.Xml.XmlDocument()
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(xmlPath)

            Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
            nsMgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
            nsMgr.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main")
            nsMgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

            ' Extract paragraphs using DrawingML structure
            Dim paragraphs As List(Of APParagraphInfo) = APExtractPptxParagraphs(xmlDoc, nsMgr)

            Dim processable = paragraphs.Where(Function(p) Not p.IsEmpty).ToList()
            If processable.Count = 0 Then Return True

            ' Process paragraphs in batches — use PPTX-aware batching when context is available
            Dim success As Boolean
            If totalSlides > 0 AndAlso Not String.IsNullOrWhiteSpace(currentSlideNum) Then
                success = Await APProcessPptxBatches(paragraphs, instruction, ct,
                                                     totalSlides, currentSlideNum, presentationOutline)
            Else
                ' Fallback for notes slides or when no context is available
                success = Await APProcessBatches(paragraphs, instruction, ct)
            End If
            If Not success Then Return False

            ' Apply processed text back to XML nodes (reuses the same redistribution logic)
            APApplyTranslations(paragraphs)

            ' Save modified XML
            xmlDoc.Save(xmlPath)
            Return True

        Catch ex As OperationCanceledException
            Throw
        Catch ex As System.Exception
            Debug.WriteLine($"APProcessPptxXmlPart error for {Path.GetFileName(xmlPath)}: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' PPTX-specific batch processing that provides slide-level context to the LLM.
    ''' The system prompt includes the full presentation outline so the LLM understands
    ''' the current slide's position and content within the deck. This helps with:
    ''' - Global operations (translate, correct, shorten): maintains consistency across slides
    ''' - Targeted operations (modify specific slide): avoids corrupting unrelated content
    ''' The LLM is always instructed to process every paragraph — the outline is purely
    ''' contextual, not a filter. The "do not alter unrelated content" rule only applies
    ''' when the instruction explicitly targets specific slides.
    ''' </summary>
    Private Async Function APProcessPptxBatches(paragraphs As List(Of APParagraphInfo),
                                                 instruction As String, ct As CancellationToken,
                                                 totalSlides As Integer, currentSlideNum As String,
                                                 presentationOutline As String) As Task(Of Boolean)

        Dim processable = paragraphs.Where(Function(p) Not p.IsEmpty).ToList()
        If processable.Count = 0 Then Return True

        ' ─── Record non-breaking spaces for later restoration ───
        Dim nbspRecords As New Dictionary(Of Integer, List(Of APNonBreakingSpaceInfo))()
        For i As Integer = 0 To processable.Count - 1
            Dim recorded = APRecordNonBreakingSpaces(processable(i).FullText)
            If recorded IsNot Nothing Then
                nbspRecords(i) = recorded
            End If
        Next

        Dim hasMarkers = processable.Any(Function(p) p.MarkerText IsNot Nothing)

        ' ─── Build PPTX-aware system prompt ───
        Dim systemPrompt As New StringBuilder()
        systemPrompt.AppendLine("You are a professional presentation processor. Apply the following instruction to the numbered paragraphs " &
                                "in the [TEXTTOPROCESS] section.")
        systemPrompt.AppendLine()
        systemPrompt.AppendLine($"INSTRUCTION: {instruction}")
        systemPrompt.AppendLine()

        ' Add presentation context so the LLM knows what the deck looks like
        If Not String.IsNullOrWhiteSpace(presentationOutline) Then
            systemPrompt.AppendLine($"PRESENTATION OUTLINE ({totalSlides} slides total):")
            systemPrompt.AppendLine(presentationOutline)
            systemPrompt.AppendLine()
            systemPrompt.AppendLine($"You are now processing Slide {currentSlideNum} of {totalSlides}.")
            systemPrompt.AppendLine()
        End If

        systemPrompt.AppendLine("RULES:")
        systemPrompt.AppendLine("1. Process ONLY paragraphs inside [TEXTTOPROCESS], not the context sections.")
        systemPrompt.AppendLine("2. Use [CONTEXT BEFORE] and [CONTEXT AFTER] to understand meaning, tone, and terminology.")
        systemPrompt.AppendLine("3. Return each processed paragraph with its [n] marker exactly as shown.")
        systemPrompt.AppendLine("4. Apply the instruction faithfully to every paragraph. For global operations " &
                                "(translate, correct, rewrite, shorten, etc.) process ALL paragraphs.")
        systemPrompt.AppendLine("5. If the instruction explicitly targets a SPECIFIC slide (e.g. 'fix the last slide', " &
                                "'add content to slide 3') and this is NOT that slide, return every paragraph UNCHANGED " &
                                "— copy the original text exactly, including any | markers.")
        systemPrompt.AppendLine("6. Maintain consistent terminology and style with the rest of the presentation.")
        systemPrompt.AppendLine("7. Return ONLY the [n] processed paragraphs, no explanations.")

        Dim ruleNum As Integer = 8
        If hasMarkers Then
            systemPrompt.AppendLine($"{ruleNum}. Some paragraphs contain pipe characters (|) that mark formatting boundaries. " &
                "IMPORTANT: Preserve these | markers in your output at approximately the same positions " &
                "relative to the text. The number of | markers in each paragraph must stay EXACTLY the same. " &
                "Do NOT add or remove any | markers.")
            ruleNum += 1
        End If

        ' ─── Batch loop (same structure as APProcessBatches) ───
        Dim batchIndex As Integer = 0
        Dim totalBatches As Integer = CInt(Math.Ceiling(processable.Count / CDbl(AP_ParagraphsPerBatch)))
        Dim currentBatch As Integer = 0

        While batchIndex < processable.Count
            ct.ThrowIfCancellationRequested()

            currentBatch += 1

            Dim batchStart = batchIndex
            Dim batchEnd = Math.Min(batchIndex + AP_ParagraphsPerBatch - 1, processable.Count - 1)

            ' Adjust for character limit
            Dim batchChars As Integer = 0
            For j = batchStart To batchEnd
                batchChars += processable(j).FullText.Length
                If batchChars > AP_MaxCharsPerBatch AndAlso j > batchStart Then
                    batchEnd = j - 1
                    Exit For
                End If
            Next

            ' Build prompt
            Dim promptBuilder As New StringBuilder()

            ' Context Before
            Dim contextBeforeStart = Math.Max(0, batchStart - AP_ContextBefore)
            If contextBeforeStart < batchStart Then
                promptBuilder.AppendLine("[CONTEXT BEFORE - for reference only]")
                For j = contextBeforeStart To batchStart - 1
                    Dim contextText = If(processable(j).TranslatedText, processable(j).FullText)
                    If contextText IsNot Nothing Then
                        contextText = contextText.Replace(AP_RunBoundaryMarker, "")
                    End If
                    promptBuilder.AppendLine(contextText)
                Next
                promptBuilder.AppendLine()
            End If

            ' Paragraphs to process — use marker text when available
            promptBuilder.AppendLine("[TEXTTOPROCESS]")
            Dim batchNumber = 1
            For j = batchStart To batchEnd
                Dim paraText = If(processable(j).MarkerText, processable(j).FullText)
                promptBuilder.AppendLine("[" & batchNumber.ToString() & "] " & paraText)
                batchNumber += 1
            Next
            promptBuilder.AppendLine("[/TEXTTOPROCESS]")
            promptBuilder.AppendLine()

            ' Context After
            Dim contextAfterEnd = Math.Min(processable.Count - 1, batchEnd + AP_ContextAfter)
            If contextAfterEnd > batchEnd Then
                promptBuilder.AppendLine("[CONTEXT AFTER - for reference only]")
                For j = batchEnd + 1 To contextAfterEnd
                    promptBuilder.AppendLine(processable(j).FullText)
                Next
            End If

            ' Call LLM
            Dim llmResponse = Await LLM(systemPrompt.ToString(), promptBuilder.ToString(),
                                         UseSecondAPI:=_apUseSecondApi,
                                         HideSplash:=True, EnsureUI:=False,
                                         cancellationToken:=ct)

            If String.IsNullOrWhiteSpace(llmResponse) Then
                ApDashboardLog($"DocProcessor: slide {currentSlideNum} batch {currentBatch} returned empty response", "warn")
                Return False
            End If

            APParseResponse(llmResponse, processable, batchStart, batchEnd)

            ' Restore non-breaking spaces
            For j As Integer = batchStart To batchEnd
                If processable(j).TranslatedText IsNot Nothing AndAlso nbspRecords.ContainsKey(j) Then
                    processable(j).TranslatedText = APRestoreNonBreakingSpaces(
                        processable(j).TranslatedText, nbspRecords(j))
                End If
            Next

            batchIndex = batchEnd + 1
        End While

        Return True
    End Function

    ''' <summary>
    ''' Extracts paragraph information from a PPTX XML part (slide, notes, etc.).
    ''' DrawingML uses a:p → a:r → a:t structure instead of WordprocessingML's w:p → w:r → w:t.
    ''' Skips a:t nodes inside a:fld (field) elements — auto-generated content like slide numbers.
    ''' </summary>
    Private Function APExtractPptxParagraphs(xmlDoc As System.Xml.XmlDocument,
                                              nsMgr As System.Xml.XmlNamespaceManager) As List(Of APParagraphInfo)
        Dim paragraphs As New List(Of APParagraphInfo)()

        Dim paraNodes As System.Xml.XmlNodeList = xmlDoc.SelectNodes("//a:p", nsMgr)
        Dim paraIndex As Integer = 0

        For Each paraNode As System.Xml.XmlNode In paraNodes
            Dim paraInfo As New APParagraphInfo() With {
                .Index = paraIndex,
                .TextRuns = New List(Of APTextRunInfo)(),
                .TranslatedText = Nothing,
                .MarkerText = Nothing
            }

            ' Find all a:t (text) elements within a:r (run) elements
            Dim textNodes As System.Xml.XmlNodeList = paraNode.SelectNodes(".//a:r/a:t", nsMgr)
            Dim fullTextBuilder As New StringBuilder()

            ' Build a set of a:t nodes inside field elements to exclude
            Dim fieldTextNodes As New HashSet(Of System.Xml.XmlNode)()
            Dim fieldNodes As System.Xml.XmlNodeList = paraNode.SelectNodes(".//a:fld//a:t", nsMgr)
            For Each fldTextNode As System.Xml.XmlNode In fieldNodes
                fieldTextNodes.Add(fldTextNode)
            Next

            For Each textNode As System.Xml.XmlNode In textNodes
                ' Skip text inside field elements (slide numbers, dates, etc.)
                If fieldTextNodes.Contains(textNode) Then Continue For

                Dim text As String = textNode.InnerText

                paraInfo.TextRuns.Add(New APTextRunInfo() With {
                    .TextNode = textNode,
                    .OriginalText = text,
                    .HasNoteReferenceBefore = False,  ' PPTX doesn't have footnote references
                    .FormattingKey = APGetRunFormattingKey(textNode)
                })

                fullTextBuilder.Append(text)
            Next

            paraInfo.FullText = fullTextBuilder.ToString()
            paraInfo.IsEmpty = String.IsNullOrWhiteSpace(paraInfo.FullText)

            ' Build marker-annotated text for multi-run paragraphs
            If paraInfo.TextRuns.Count > 1 AndAlso Not paraInfo.IsEmpty Then
                paraInfo.MarkerText = APBuildMarkerAnnotatedText(paraInfo.TextRuns)
            Else
                paraInfo.MarkerText = Nothing
            End If

            paragraphs.Add(paraInfo)
            paraIndex += 1
        Next

        Return paragraphs
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  XLSX PROCESSING
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Processes an .xlsx file by extracting ALL cell content (text, numbers, formulas),
    ''' sending them through an Excel-specific batch pipeline, and writing the LLM's
    ''' changes back — preserving cell types, styles, and structure.
    ''' </summary>
    Private Async Function ProcessXlsxForAutoPilot(inputPath As String, outputPath As String,
                                                    instruction As String, ct As CancellationToken,
                                                    Optional sheetFilter As List(Of String) = Nothing) As Task(Of Boolean)

        Dim tempDir As String = Path.Combine(Path.GetTempPath(), $"{AN2}_apxlsx_{Guid.NewGuid():N}")

        Try
            ApDashboardLog("DocProcessor: starting Excel processing", "step")

            File.Copy(inputPath, outputPath, True)
            ZipFile.ExtractToDirectory(outputPath, tempDir)

            ' ── 1. Load shared strings table (may not exist) ──
            Dim sstPath = Path.Combine(tempDir, "xl", "sharedStrings.xml")
            Dim sstDoc As System.Xml.XmlDocument = Nothing
            Dim sstNsMgr As System.Xml.XmlNamespaceManager = Nothing
            Dim sharedStrings As New List(Of System.Xml.XmlNode)()

            If File.Exists(sstPath) Then
                sstDoc = New System.Xml.XmlDocument()
                sstDoc.PreserveWhitespace = True
                sstDoc.Load(sstPath)
                sstNsMgr = New System.Xml.XmlNamespaceManager(sstDoc.NameTable)
                sstNsMgr.AddNamespace("x", AP_XlsxNs)

                For Each siNode As System.Xml.XmlNode In sstDoc.SelectNodes("//x:si", sstNsMgr)
                    sharedStrings.Add(siNode)
                Next
            End If

            ' ── 2. Resolve sheet names → file mappings ──
            Dim sheetMap = APResolveXlsxSheets(tempDir)
            If sheetMap.Count = 0 Then
                ApDashboardLog("⚠ No worksheets found in workbook.", "warn")
                Return False
            End If

            ' ── 3. Process each sheet ──
            Dim anyProcessed As Boolean = False
            Dim sstCreatedDuringProcessing As Boolean = False

            For Each entry In sheetMap
                Dim sheetName = entry.Key
                Dim sheetXmlPath = entry.Value

                ' Apply sheet filter
                If sheetFilter IsNot Nothing AndAlso sheetFilter.Count > 0 Then
                    If Not sheetFilter.Any(Function(f) f.Equals(sheetName, StringComparison.OrdinalIgnoreCase)) Then
                        ApDashboardLog($"⏭ Skipping sheet: {sheetName} (not in filter)", "info")
                        Continue For
                    End If
                End If

                If Not File.Exists(sheetXmlPath) Then Continue For

                ApDashboardLog($"📊 Processing sheet: {sheetName}", "step")

                Dim sheetDoc As New System.Xml.XmlDocument()
                sheetDoc.PreserveWhitespace = True
                sheetDoc.Load(sheetXmlPath)

                Dim sheetNsMgr As New System.Xml.XmlNamespaceManager(sheetDoc.NameTable)
                sheetNsMgr.AddNamespace("x", AP_XlsxNs)

                ' Extract all cells with their content and metadata
                Dim cellEntries = APExtractAllXlsxCells(sheetDoc, sheetNsMgr, sharedStrings, sstNsMgr)

                If cellEntries.Count = 0 Then
                    ApDashboardLog($"⏭ Sheet '{sheetName}': no cells found.", "info")
                    Continue For
                End If

                ApDashboardLog($"📊 Sheet '{sheetName}': {cellEntries.Count} cell(s) extracted.", "info")

                ' Run through the Excel-specific batch pipeline
                Dim success = Await APProcessXlsxBatches(cellEntries, sheetName, instruction, ct)
                If Not success Then
                    ApDashboardLog($"⚠ Batch processing failed for sheet: {sheetName}", "warn")
                    Continue For
                End If

                ' Check if any text changes need an SST that doesn't exist yet
                Dim needsSst = cellEntries.Any(Function(c)
                                                   If c.NewValue Is Nothing Then Return False
                                                   If c.NewValue.StartsWith("=") Then Return False
                                                   If String.IsNullOrEmpty(c.NewValue) Then Return False
                                                   Dim numVal As Double
                                                   Return Not Double.TryParse(c.NewValue,
                                                        Globalization.NumberStyles.Any,
                                                        Globalization.CultureInfo.InvariantCulture, numVal)
                                               End Function)

                If needsSst AndAlso sstDoc Is Nothing Then
                    ' Create the shared strings table on demand
                    sstDoc = New System.Xml.XmlDocument()
                    sstDoc.PreserveWhitespace = True

                    Dim xmlDecl = sstDoc.CreateXmlDeclaration("1.0", "UTF-8", "yes")
                    sstDoc.AppendChild(xmlDecl)

                    Dim sstRoot = sstDoc.CreateElement("sst", AP_XlsxNs)
                    sstRoot.SetAttribute("count", "0")
                    sstRoot.SetAttribute("uniqueCount", "0")
                    sstDoc.AppendChild(sstRoot)

                    sstNsMgr = New System.Xml.XmlNamespaceManager(sstDoc.NameTable)
                    sstNsMgr.AddNamespace("x", AP_XlsxNs)

                    sstCreatedDuringProcessing = True
                    ApDashboardLog("DocProcessor: created shared strings table on demand", "step")
                End If

                ' Apply changes back to XML
                APApplyXlsxChanges(cellEntries, sheetDoc, sheetNsMgr, sharedStrings, sstDoc, sstNsMgr)

                ' Save sheet preserving original XML declaration
                APSaveXmlPreservingDeclaration(sheetDoc, sheetXmlPath)

                anyProcessed = True
                ApDashboardLog($"✓ Sheet '{sheetName}': processed successfully.", "info")
            Next

            ' ── 4. Save modified shared strings ──
            If sstDoc IsNot Nothing AndAlso anyProcessed Then
                Dim sstRoot = sstDoc.SelectSingleNode("//x:sst", sstNsMgr)
                If sstRoot IsNot Nothing Then
                    Dim newCount = sstDoc.SelectNodes("//x:si", sstNsMgr).Count.ToString()
                    Dim sstEl = DirectCast(sstRoot, System.Xml.XmlElement)
                    If sstEl.HasAttribute("count") Then
                        sstEl.SetAttribute("count", newCount)
                    End If
                    If sstEl.HasAttribute("uniqueCount") Then
                        sstEl.SetAttribute("uniqueCount", newCount)
                    End If
                End If

                APSaveXmlPreservingDeclaration(sstDoc, sstPath)

                ' If we created the SST from scratch, register it in the package
                If sstCreatedDuringProcessing Then
                    APRegisterSharedStringsInPackage(tempDir)
                End If
            End If

            ' ── 5. Repack .xlsx with forward-slash entry paths (OOXML requirement) ──
            If anyProcessed Then
                File.Delete(outputPath)
                Using zipStream As New FileStream(outputPath, FileMode.Create, FileAccess.Write)
                    Using archive As New ZipArchive(zipStream, ZipArchiveMode.Create)
                        For Each filePath In Directory.GetFiles(tempDir, "*", SearchOption.AllDirectories)
                            ' Build the entry name relative to tempDir, then normalize to forward slashes
                            Dim entryName = filePath.Substring(tempDir.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                            entryName = entryName.Replace("\"c, "/"c)

                            Dim entry = archive.CreateEntry(entryName, CompressionLevel.Fastest)
                            Using entryStream = entry.Open()
                                Using fileStream As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                                    fileStream.CopyTo(entryStream)
                                End Using
                            End Using
                        Next
                    End Using
                End Using
            End If

            ApDashboardLog($"DocProcessor: Excel processing complete", "success")
            Return anyProcessed

        Catch ex As OperationCanceledException
            Throw
        Catch ex As Exception
            Debug.WriteLine($"ProcessXlsxForAutoPilot error: {ex.Message}")
            ApDashboardLog($"⚠ Excel processing error: {ex.Message}", "warn")
            Return False
        Finally
            If Directory.Exists(tempDir) Then
                Try : Directory.Delete(tempDir, True) : Catch : End Try
            End If
        End Try
    End Function

    ''' <summary>
    ''' Registers a newly created sharedStrings.xml in [Content_Types].xml and
    ''' xl/_rels/workbook.xml.rels so Excel recognizes the part.
    ''' Called only when the original .xlsx had no shared strings table.
    ''' </summary>
    Private Sub APRegisterSharedStringsInPackage(tempDir As String)
        Const sstContentType As String = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
        Const sstRelType As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"

        ' ── [Content_Types].xml ──
        Dim contentTypesPath = Path.Combine(tempDir, "[Content_Types].xml")
        If File.Exists(contentTypesPath) Then
            Dim ctDoc As New System.Xml.XmlDocument()
            ctDoc.PreserveWhitespace = True
            ctDoc.Load(contentTypesPath)

            Dim ctNs As String = "http://schemas.openxmlformats.org/package/2006/content-types"
            Dim ctNsMgr As New System.Xml.XmlNamespaceManager(ctDoc.NameTable)
            ctNsMgr.AddNamespace("ct", ctNs)

            Dim existing = ctDoc.SelectSingleNode("//ct:Override[@PartName='/xl/sharedStrings.xml']", ctNsMgr)
            If existing Is Nothing Then
                Dim overrideEl = ctDoc.CreateElement("Override", ctNs)

                Dim partNameAttr = ctDoc.CreateAttribute("PartName")
                partNameAttr.Value = "/xl/sharedStrings.xml"
                overrideEl.Attributes.Append(partNameAttr)

                Dim contentTypeAttr = ctDoc.CreateAttribute("ContentType")
                contentTypeAttr.Value = sstContentType
                overrideEl.Attributes.Append(contentTypeAttr)

                ctDoc.DocumentElement.AppendChild(overrideEl)

                APSaveXmlPreservingDeclaration(ctDoc, contentTypesPath)

            End If
        End If

        ' ── xl/_rels/workbook.xml.rels ──
        Dim relsDir = Path.Combine(tempDir, "xl", "_rels")
        Dim relsPath = Path.Combine(relsDir, "workbook.xml.rels")

        If Not Directory.Exists(relsDir) Then
            Directory.CreateDirectory(relsDir)
        End If

        Dim relsNs As String = "http://schemas.openxmlformats.org/package/2006/relationships"

        If File.Exists(relsPath) Then
            Dim relsDoc As New System.Xml.XmlDocument()
            relsDoc.PreserveWhitespace = True
            relsDoc.Load(relsPath)

            Dim relsNsMgr As New System.Xml.XmlNamespaceManager(relsDoc.NameTable)
            relsNsMgr.AddNamespace("r", relsNs)

            ' Check if relationship already exists
            Dim existingRel = relsDoc.SelectSingleNode(
                "//r:Relationship[@Type='" & sstRelType & "']", relsNsMgr)
            If existingRel Is Nothing Then
                ' Determine next rId
                Dim maxId As Integer = 0
                Dim allRels = relsDoc.SelectNodes("//r:Relationship", relsNsMgr)
                For Each rel As System.Xml.XmlNode In allRels
                    Dim idVal = rel.Attributes("Id")?.Value
                    If idVal IsNot Nothing AndAlso idVal.StartsWith("rId") Then
                        Dim num As Integer
                        If Integer.TryParse(idVal.Substring(3), num) AndAlso num > maxId Then
                            maxId = num
                        End If
                    End If
                Next

                Dim newRel = relsDoc.CreateElement("Relationship", relsNs)

                Dim idAttr = relsDoc.CreateAttribute("Id")
                idAttr.Value = "rId" & (maxId + 1).ToString()
                newRel.Attributes.Append(idAttr)

                Dim typeAttr = relsDoc.CreateAttribute("Type")
                typeAttr.Value = sstRelType
                newRel.Attributes.Append(typeAttr)

                Dim targetAttr = relsDoc.CreateAttribute("Target")
                targetAttr.Value = "sharedStrings.xml"
                newRel.Attributes.Append(targetAttr)

                relsDoc.DocumentElement.AppendChild(newRel)

                APSaveXmlPreservingDeclaration(relsDoc, relsPath)
            End If
        End If
    End Sub


    ' ═══════════════════════════════════════════════════════════════════════════
    '  XLSX DATA CLASS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Represents a single Excel cell with its address, content, type, and XML node
    ''' references for write-back.
    ''' </summary>
    Private Class APXlsxCellInfo
        ''' <summary>Cell address, e.g. "A1", "B12".</summary>
        Public Property CellRef As String

        ''' <summary>Display text sent to the LLM. For shared strings: the resolved text.
        ''' For numbers: the numeric value. For formulas: "=FORMULA".</summary>
        Public Property DisplayText As String

        ''' <summary>Cell type: "text", "number", "formula", "boolean", "empty".</summary>
        Public Property CellType As String

        ''' <summary>The &lt;c&gt; element in sheet XML.</summary>
        Public Property CellElement As System.Xml.XmlElement

        ''' <summary>For shared string cells: the index into the shared strings table.</summary>
        Public Property SharedStringIndex As Integer

        ''' <summary>New value from LLM. Nothing = no change requested.</summary>
        Public Property NewValue As String
    End Class

    ' ═══════════════════════════════════════════════════════════════════════════
    '  XLSX CELL EXTRACTION (ALL cells, not just text)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Extracts ALL non-empty cells from a worksheet, including text, numbers,
    ''' formulas, and booleans — so the LLM can see the full picture and decide
    ''' what to change.
    ''' </summary>
    Private Function APExtractAllXlsxCells(
            sheetDoc As System.Xml.XmlDocument,
            sheetNsMgr As System.Xml.XmlNamespaceManager,
            sharedStrings As List(Of System.Xml.XmlNode),
            sstNsMgr As System.Xml.XmlNamespaceManager) As List(Of APXlsxCellInfo)

        Dim cells As New List(Of APXlsxCellInfo)()
        Dim cellNodes = sheetDoc.SelectNodes("//x:sheetData/x:row/x:c", sheetNsMgr)

        For Each cellNode As System.Xml.XmlElement In cellNodes
            Dim cellRef = cellNode.GetAttribute("r")
            If String.IsNullOrEmpty(cellRef) Then Continue For

            Dim cellType = cellNode.GetAttribute("t")
            Dim formulaNode = cellNode.SelectSingleNode("x:f", sheetNsMgr)
            Dim vNode = cellNode.SelectSingleNode("x:v", sheetNsMgr)

            Dim info As New APXlsxCellInfo() With {
                .CellRef = cellRef,
                .CellElement = cellNode,
                .SharedStringIndex = -1,
                .NewValue = Nothing
            }

            If formulaNode IsNot Nothing Then
                ' Formula cell — show the formula so the LLM can modify it
                Dim formulaText = formulaNode.InnerText
                Dim cachedValue = If(vNode?.InnerText, "")
                info.CellType = "formula"
                info.DisplayText = "=" & formulaText
            ElseIf cellType = "s" Then
                ' Shared string reference
                If vNode Is Nothing Then Continue For
                Dim ssIndex As Integer
                If Not Integer.TryParse(vNode.InnerText, ssIndex) Then Continue For
                If ssIndex < 0 OrElse ssIndex >= sharedStrings.Count Then Continue For

                Dim siNode = sharedStrings(ssIndex)
                Dim tNodes = siNode.SelectNodes(".//x:t", sstNsMgr)
                If tNodes Is Nothing OrElse tNodes.Count = 0 Then Continue For

                Dim fullText As New StringBuilder()
                For Each tNode As System.Xml.XmlNode In tNodes
                    fullText.Append(tNode.InnerText)
                Next

                info.CellType = "text"
                info.DisplayText = fullText.ToString()
                info.SharedStringIndex = ssIndex
            ElseIf cellType = "inlineStr" Then
                Dim isNode = cellNode.SelectSingleNode("x:is", sheetNsMgr)
                If isNode Is Nothing Then Continue For
                Dim tNodes = isNode.SelectNodes(".//x:t", sheetNsMgr)
                Dim fullText As New StringBuilder()
                For Each tNode As System.Xml.XmlNode In tNodes
                    fullText.Append(tNode.InnerText)
                Next
                info.CellType = "text"
                info.DisplayText = fullText.ToString()
            ElseIf cellType = "b" Then
                info.CellType = "boolean"
                info.DisplayText = If(vNode?.InnerText = "1", "TRUE", "FALSE")
            Else
                ' Numeric or empty
                If vNode IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(vNode.InnerText) Then
                    info.CellType = "number"
                    info.DisplayText = vNode.InnerText
                Else
                    Continue For ' Truly empty cell
                End If
            End If

            If String.IsNullOrEmpty(info.DisplayText) Then Continue For

            cells.Add(info)
        Next

        Return cells
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  XLSX BATCH PROCESSING (Excel-specific, not reusing APProcessBatches)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Sends Excel cells to the LLM in batches using a cell-address format.
    ''' The LLM sees ALL cells for context but only returns cells it wants to change.
    ''' </summary>
    Private Async Function APProcessXlsxBatches(
            cells As List(Of APXlsxCellInfo),
            sheetName As String,
            instruction As String,
            ct As CancellationToken) As Task(Of Boolean)

        If cells.Count = 0 Then Return True

        ' ── Build structural summary of the entire sheet ──
        ' This gives the LLM a bird's-eye view: column headings, data types,
        ' row count, and sample rows — so it understands what each column means
        ' even when processing cells deep into the sheet.
        Dim structureSummary As String = APBuildXlsxStructureSummary(cells, sheetName)

        ' ── Build column header map from row 1 cells (for per-batch header injection) ──
        Dim columnHeaders As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For Each c In cells
            Dim rowNum As String = ""
            Dim colLetter As String = ""
            For Each ch In c.CellRef
                If Char.IsLetter(ch) Then
                    colLetter &= ch
                Else
                    rowNum &= ch
                End If
            Next
            If rowNum = "1" AndAlso c.CellType = "text" AndAlso Not String.IsNullOrWhiteSpace(c.DisplayText) Then
                columnHeaders(colLetter) = c.DisplayText
            End If
        Next

        Dim systemPrompt As String =
            "You are a professional spreadsheet processor. Apply the following instruction to the cells " &
            "shown in the [CELLS] section." & vbCrLf & vbCrLf &
            "INSTRUCTION: " & instruction & vbCrLf & vbCrLf &
            If(Not String.IsNullOrWhiteSpace(structureSummary),
               structureSummary & vbCrLf, "") &
            "RULES:" & vbCrLf &
            "1. You will see cells in the format [CellRef] content (e.g. [A1] Hello)." & vbCrLf &
            "2. Return ONLY the cells you want to CHANGE, in the same [CellRef] format." & vbCrLf &
            "3. Cells you do NOT return will remain unchanged." & vbCrLf &
            "4. For formulas, include the = prefix (e.g. [B5] =SUM(A1:A4))." & vbCrLf &
            "5. For text values, just write the text (e.g. [A1] Translated text)." & vbCrLf &
            "6. For numbers, write the number (e.g. [C3] 42.5)." & vbCrLf &
            "7. To clear a cell, write [CellRef] (empty)." & vbCrLf &
            "8. Return ONLY [CellRef] lines, no explanations or commentary." & vbCrLf &
            "9. If no cells need changing, return exactly: NO_CHANGES" & vbCrLf &
            "10. When the instruction targets a specific column by name, ONLY change cells in that column. " &
            "Use the SPREADSHEET STRUCTURE section above to identify which column letter corresponds to which heading." & vbCrLf &
            "11. CRITICAL LANGUAGE RULE: The INSTRUCTION may be written in a different language than the cell content. " &
            "Unless the instruction EXPLICITLY asks for translation to a specific language, you MUST keep the cell " &
            "text in its ORIGINAL language. NEVER translate the content. Apply the instruction while preserving " &
            "the original language of the text."

        Dim batchIndex As Integer = 0
        Dim totalBatches As Integer = CInt(Math.Ceiling(cells.Count / CDbl(AP_ParagraphsPerBatch)))
        Dim currentBatch As Integer = 0

        ApDashboardLog($"DocProcessor: sheet '{sheetName}': {cells.Count} cells to process (~{totalBatches} batches)", "step")

        While batchIndex < cells.Count
            ct.ThrowIfCancellationRequested()

            currentBatch += 1

            Dim batchStart = batchIndex
            Dim batchEnd = Math.Min(batchIndex + AP_ParagraphsPerBatch - 1, cells.Count - 1)

            ' Adjust for character limit
            Dim batchChars As Integer = 0
            For j = batchStart To batchEnd
                batchChars += cells(j).DisplayText.Length + cells(j).CellRef.Length + 3
                If batchChars > AP_MaxCharsPerBatch AndAlso j > batchStart Then
                    batchEnd = j - 1
                    Exit For
                End If
            Next

            ApDashboardLog($"DocProcessor: sheet '{sheetName}' batch {currentBatch}/{totalBatches} ({batchEnd - batchStart + 1} cells)", "step")

            ' Build prompt
            Dim promptBuilder As New StringBuilder()

            ' ── Always include header row at the top for structural context ──
            If columnHeaders.Count > 0 Then
                Dim headerCells = cells.Where(Function(c)
                                                  Dim rn = ""
                                                  For Each ch In c.CellRef
                                                      If Not Char.IsLetter(ch) Then rn &= ch
                                                  Next
                                                  Return rn = "1"
                                              End Function).ToList()

                ' Only add if header row is NOT already in this batch
                Dim batchContainsRow1 = False
                For j = batchStart To batchEnd
                    Dim rn = ""
                    For Each ch In cells(j).CellRef
                        If Not Char.IsLetter(ch) Then rn &= ch
                    Next
                    If rn = "1" Then batchContainsRow1 = True : Exit For
                Next

                If Not batchContainsRow1 AndAlso headerCells.Count > 0 Then
                    promptBuilder.AppendLine("[COLUMN HEADERS - row 1, for reference only, do not modify]")
                    For Each hc In headerCells
                        promptBuilder.AppendLine($"[{hc.CellRef}] {If(hc.NewValue, hc.DisplayText)}")
                    Next
                    promptBuilder.AppendLine()
                End If
            End If

            ' Context before (previous batch cells for continuity)
            Dim contextStart = Math.Max(0, batchStart - AP_ContextBefore)
            If contextStart < batchStart Then
                promptBuilder.AppendLine("[CONTEXT - for reference only, do not modify]")
                For j = contextStart To batchStart - 1
                    Dim ctxText = If(cells(j).NewValue, cells(j).DisplayText)
                    promptBuilder.AppendLine($"[{cells(j).CellRef}] {ctxText}")
                Next
                promptBuilder.AppendLine()
            End If

            ' Cells to process
            promptBuilder.AppendLine("[CELLS]")
            For j = batchStart To batchEnd
                promptBuilder.AppendLine($"[{cells(j).CellRef}] {cells(j).DisplayText}")
            Next
            promptBuilder.AppendLine("[/CELLS]")

            ' Context after
            Dim contextEnd = Math.Min(cells.Count - 1, batchEnd + AP_ContextAfter)
            If contextEnd > batchEnd Then
                promptBuilder.AppendLine()
                promptBuilder.AppendLine("[CONTEXT - for reference only, do not modify]")
                For j = batchEnd + 1 To contextEnd
                    promptBuilder.AppendLine($"[{cells(j).CellRef}] {cells(j).DisplayText}")
                Next
            End If

            ' Call LLM
            Dim llmResponse = Await LLM(systemPrompt, promptBuilder.ToString(),
                                         UseSecondAPI:=_apUseSecondApi,
                                         HideSplash:=True, EnsureUI:=False,
                                         cancellationToken:=ct)

            If String.IsNullOrWhiteSpace(llmResponse) Then
                ApDashboardLog($"DocProcessor: sheet '{sheetName}' batch {currentBatch} returned empty response", "warn")
                Return False
            End If

            ' Parse response — only changed cells
            If Not llmResponse.Trim().Equals("NO_CHANGES", StringComparison.OrdinalIgnoreCase) Then
                APParseXlsxResponse(llmResponse, cells, batchStart, batchEnd)
            End If

            batchIndex = batchEnd + 1
        End While

        Dim changedCount = cells.Where(Function(c) c.NewValue IsNot Nothing).Count()
        ApDashboardLog($"DocProcessor: sheet '{sheetName}': {currentBatch} batches completed, {changedCount} cell(s) changed", "success")
        Return True
    End Function

    ''' <summary>
    ''' Parses LLM output for Excel cells. Expects lines like: [A1] new value
    ''' Only cells that appear in the response are marked as changed.
    ''' </summary>
    Private Sub APParseXlsxResponse(response As String, cells As List(Of APXlsxCellInfo),
                                     batchStart As Integer, batchEnd As Integer)
        ' Build a lookup from cell ref → cell info for the current batch
        Dim cellLookup As New Dictionary(Of String, APXlsxCellInfo)(StringComparer.OrdinalIgnoreCase)
        For j = batchStart To Math.Min(batchEnd, cells.Count - 1)
            cellLookup(cells(j).CellRef) = cells(j)
        Next

        ' Parse [CellRef] value lines
        Dim pattern As New Regex("^\[([A-Z]+\d+)\]\s*(.*)", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
        Dim matches = pattern.Matches(response)

        For Each m As Match In matches
            Dim cellRef = m.Groups(1).Value.ToUpperInvariant()
            Dim newValue = m.Groups(2).Value.Trim()

            ' Handle "(empty)" as clear
            If newValue.Equals("(empty)", StringComparison.OrdinalIgnoreCase) Then
                newValue = ""
            End If

            Dim cell As APXlsxCellInfo = Nothing
            If cellLookup.TryGetValue(cellRef, cell) Then
                ' Only mark as changed if the value actually differs
                If Not newValue.Equals(cell.DisplayText, StringComparison.Ordinal) Then
                    cell.NewValue = newValue
                End If
            End If
        Next
    End Sub


    ''' <summary>
    ''' Builds a concise structural summary of an Excel worksheet for LLM context.
    ''' Includes column headers, row count, sample rows, and data type profile.
    ''' This summary is prepended to the instruction so the batch-processing LLM
    ''' understands the full spreadsheet structure before seeing individual cells.
    ''' </summary>
    Private Function APBuildXlsxStructureSummary(
            cells As List(Of APXlsxCellInfo),
            sheetName As String) As String

        If cells.Count = 0 Then Return ""

        ' ── Parse cell refs into (column, row) tuples ──
        Dim parsed As New List(Of (Cell As APXlsxCellInfo, Col As String, Row As Integer))()
        For Each c In cells
            Dim col As String = ""
            Dim rowStr As String = ""
            For Each ch In c.CellRef
                If Char.IsLetter(ch) Then col &= ch Else rowStr &= ch
            Next
            Dim rowNum As Integer
            If Integer.TryParse(rowStr, rowNum) Then
                parsed.Add((c, col.ToUpperInvariant(), rowNum))
            End If
        Next

        If parsed.Count = 0 Then Return ""

        ' ── Identify header row (row 1 text cells) ──
        Dim headers = parsed.Where(Function(p) p.Row = 1 AndAlso p.Cell.CellType = "text") _
                            .OrderBy(Function(p) p.Col) _
                            .ToList()

        ' ── Compute basic stats ──
        Dim maxRow = parsed.Max(Function(p) p.Row)
        Dim allColumns = parsed.Select(Function(p) p.Col).Distinct().OrderBy(Function(c) c).ToList()
        Dim dataRowCount = If(headers.Count > 0, maxRow - 1, maxRow)

        Dim sb As New StringBuilder()
        sb.AppendLine($"[SPREADSHEET STRUCTURE — Sheet: ""{sheetName}""]")
        sb.AppendLine($"Columns: {allColumns.Count} ({String.Join(", ", allColumns)})")
        sb.AppendLine($"Data rows: {dataRowCount} (excluding header)")
        sb.AppendLine()

        ' ── Column headers with data type profile ──
        If headers.Count > 0 Then
            sb.AppendLine("Column headings (row 1):")
            For Each h In headers
                ' Determine predominant data type in this column (skip row 1)
                Dim colCells = parsed.Where(Function(p) p.Col = h.Col AndAlso p.Row > 1).ToList()
                Dim typeProfile As String = ""
                If colCells.Count > 0 Then
                    Dim typeCounts = colCells.GroupBy(Function(p) p.Cell.CellType) _
                                            .OrderByDescending(Function(g) g.Count()) _
                                            .Select(Function(g) g.Key) _
                                            .ToList()
                    typeProfile = $" [{String.Join("/", typeCounts)}, {colCells.Count} values]"
                End If
                sb.AppendLine($"  Column {h.Col} = ""{h.Cell.DisplayText}""{typeProfile}")
            Next
            sb.AppendLine()
        End If

        ' ── Sample rows (first 3 data rows) for pattern recognition ──
        Dim sampleRows = parsed.Where(Function(p) p.Row > 1) _
                               .GroupBy(Function(p) p.Row) _
                               .OrderBy(Function(g) g.Key) _
                               .Take(3) _
                               .ToList()

        If sampleRows.Count > 0 Then
            sb.AppendLine("Sample data (first rows):")
            For Each rowGroup In sampleRows
                Dim rowCells = rowGroup.OrderBy(Function(p) p.Col) _
                                      .Select(Function(p) $"[{p.Cell.CellRef}] {p.Cell.DisplayText}") _
                                      .ToList()
                sb.AppendLine($"  Row {rowGroup.Key}: {String.Join("  ", rowCells)}")
            Next
            sb.AppendLine()
        End If

        sb.AppendLine("[/SPREADSHEET STRUCTURE]")
        Return sb.ToString()
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  XLSX WRITE-BACK
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Applies LLM changes back to the SpreadsheetML XML. Handles type transitions:
    ''' text→formula, number→text, formula→text, etc.
    ''' Uses ImportNode-based element creation to avoid redundant xmlns attributes
    ''' that cause Excel to flag the file as corrupted.
    ''' </summary>
    Private Sub APApplyXlsxChanges(
            cells As List(Of APXlsxCellInfo),
            sheetDoc As System.Xml.XmlDocument,
            sheetNsMgr As System.Xml.XmlNamespaceManager,
            sharedStrings As List(Of System.Xml.XmlNode),
            sstDoc As System.Xml.XmlDocument,
            sstNsMgr As System.Xml.XmlNamespaceManager)

        For Each cell In cells
            If cell.NewValue Is Nothing Then Continue For ' No change

            Dim el = cell.CellElement
            Dim newVal = cell.NewValue

            ' Locate existing child nodes
            Dim existingF = el.SelectSingleNode("x:f", sheetNsMgr)
            Dim existingV = el.SelectSingleNode("x:v", sheetNsMgr)
            Dim existingIs = el.SelectSingleNode("x:is", sheetNsMgr)

            If newVal.StartsWith("=") Then
                ' ── New formula ──
                Dim formulaText = newVal.Substring(1)
                If existingF IsNot Nothing Then
                    existingF.InnerText = formulaText
                Else
                    ' Create <f> by parsing a fragment so it inherits the default namespace
                    ' without emitting an explicit xmlns attribute
                    Dim fNode = APCreateSheetElement(sheetDoc, "f", sheetNsMgr)
                    fNode.InnerText = formulaText
                    If existingV IsNot Nothing Then
                        el.InsertBefore(fNode, existingV)
                    Else
                        el.AppendChild(fNode)
                    End If
                End If
                ' Remove cached value — Excel will recalculate
                If existingV IsNot Nothing Then el.RemoveChild(existingV)
                If existingIs IsNot Nothing Then el.RemoveChild(existingIs)
                ' Remove type attribute (formula cells don't need t="s")
                el.RemoveAttribute("t")

            ElseIf String.IsNullOrEmpty(newVal) Then
                ' ── Clear cell ──
                If existingF IsNot Nothing Then el.RemoveChild(existingF)
                If existingV IsNot Nothing Then el.RemoveChild(existingV)
                If existingIs IsNot Nothing Then el.RemoveChild(existingIs)
                el.RemoveAttribute("t")

            Else
                ' ── Value (text or number) ──
                If existingF IsNot Nothing Then el.RemoveChild(existingF)
                If existingIs IsNot Nothing Then el.RemoveChild(existingIs)

                Dim numVal As Double
                Dim isNumeric = Double.TryParse(newVal, Globalization.NumberStyles.Any,
                                                 Globalization.CultureInfo.InvariantCulture, numVal)

                If isNumeric Then
                    ' Numeric value — write directly to <v>, no type attribute
                    el.RemoveAttribute("t")
                    If existingV IsNot Nothing Then
                        existingV.InnerText = numVal.ToString(Globalization.CultureInfo.InvariantCulture)
                    Else
                        Dim vNode = APCreateSheetElement(sheetDoc, "v", sheetNsMgr)
                        vNode.InnerText = numVal.ToString(Globalization.CultureInfo.InvariantCulture)
                        el.AppendChild(vNode)
                    End If
                Else
                    ' Text value — add as new shared string entry
                    If sstDoc IsNot Nothing Then
                        ' Create <si><t>text</t></si> using the SST document's namespace context
                        ' to avoid redundant xmlns attributes
                        Dim newSi = APCreateSstElement(sstDoc, "si", sstNsMgr)
                        Dim newT = APCreateSstElement(sstDoc, "t", sstNsMgr)
                        If newVal.StartsWith(" ") OrElse newVal.EndsWith(" ") Then
                            Dim spaceAttr = sstDoc.CreateAttribute("xml", "space", "http://www.w3.org/XML/1998/namespace")
                            spaceAttr.Value = "preserve"
                            newT.Attributes.Append(spaceAttr)
                        End If
                        newT.InnerText = newVal
                        newSi.AppendChild(newT)
                        sstDoc.DocumentElement.AppendChild(newSi)

                        Dim newIndex = sharedStrings.Count
                        sharedStrings.Add(newSi)

                        ' Update the cell to reference the new shared string
                        el.SetAttribute("t", "s")
                        If existingV IsNot Nothing Then
                            existingV.InnerText = newIndex.ToString()
                        Else
                            Dim vNode = APCreateSheetElement(sheetDoc, "v", sheetNsMgr)
                            vNode.InnerText = newIndex.ToString()
                            el.AppendChild(vNode)
                        End If
                    Else
                        ' No shared strings table — use inline string
                        el.SetAttribute("t", "inlineStr")
                        If existingV IsNot Nothing Then el.RemoveChild(existingV)
                        Dim isNode = APCreateSheetElement(sheetDoc, "is", sheetNsMgr)
                        Dim tNode = APCreateSheetElement(sheetDoc, "t", sheetNsMgr)
                        tNode.InnerText = newVal
                        isNode.AppendChild(tNode)
                        el.AppendChild(isNode)
                    End If
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Creates a SpreadsheetML element in the sheet document without emitting a redundant
    ''' xmlns attribute. Uses a parsed XML fragment that inherits the default namespace
    ''' from a parent context, then imports it into the target document.
    ''' </summary>
    Private Shared Function APCreateSheetElement(sheetDoc As System.Xml.XmlDocument,
                                                  localName As String,
                                                  nsMgr As System.Xml.XmlNamespaceManager) As System.Xml.XmlElement
        ' Build a minimal XML fragment with the namespace declared on a wrapper,
        ' so the inner element inherits it without its own xmlns attribute.
        Dim xml = $"<wrapper xmlns=""{AP_XlsxNs}""><{localName}/></wrapper>"
        Dim fragDoc As New System.Xml.XmlDocument()
        fragDoc.LoadXml(xml)
        Dim imported = sheetDoc.ImportNode(fragDoc.DocumentElement.FirstChild, True)
        Return DirectCast(imported, System.Xml.XmlElement)
    End Function

    ''' <summary>
    ''' Creates a SpreadsheetML element in the shared strings document without emitting
    ''' a redundant xmlns attribute.
    ''' </summary>
    Private Shared Function APCreateSstElement(sstDoc As System.Xml.XmlDocument,
                                                localName As String,
                                                nsMgr As System.Xml.XmlNamespaceManager) As System.Xml.XmlElement
        Dim xml = $"<wrapper xmlns=""{AP_XlsxNs}""><{localName}/></wrapper>"
        Dim fragDoc As New System.Xml.XmlDocument()
        fragDoc.LoadXml(xml)
        Dim imported = sstDoc.ImportNode(fragDoc.DocumentElement.FirstChild, True)
        Return DirectCast(imported, System.Xml.XmlElement)
    End Function


    ''' <summary>
    ''' Resolves workbook sheet names to worksheet XML file paths inside the extracted package.
    ''' </summary>
    ''' <returns>
    ''' Dictionary mapping sheet display name to physical worksheet part path.
    ''' Returns an empty dictionary if workbook metadata or relationships are missing.
    ''' </returns>
    Private Function APResolveXlsxSheets(tempDir As String) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Dim wbPath = Path.Combine(tempDir, "xl", "workbook.xml")
        Dim relsPath = Path.Combine(tempDir, "xl", "_rels", "workbook.xml.rels")
        If Not File.Exists(wbPath) OrElse Not File.Exists(relsPath) Then Return result

        Dim relsDoc As New System.Xml.XmlDocument()
        relsDoc.Load(relsPath)
        Dim relNsMgr As New System.Xml.XmlNamespaceManager(relsDoc.NameTable)
        relNsMgr.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")

        Dim relMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For Each relNode As System.Xml.XmlElement In relsDoc.SelectNodes("//r:Relationship", relNsMgr)
            Dim rId = relNode.GetAttribute("Id")
            Dim target = relNode.GetAttribute("Target")
            If Not String.IsNullOrWhiteSpace(rId) AndAlso Not String.IsNullOrWhiteSpace(target) Then
                relMap(rId) = target
            End If
        Next

        Dim wbDoc As New System.Xml.XmlDocument()
        wbDoc.Load(wbPath)
        Dim wbNsMgr As New System.Xml.XmlNamespaceManager(wbDoc.NameTable)
        wbNsMgr.AddNamespace("x", AP_XlsxNs)
        wbNsMgr.AddNamespace("r", AP_XlsxRelNs)

        For Each sheetNode As System.Xml.XmlElement In wbDoc.SelectNodes("//x:sheets/x:sheet", wbNsMgr)
            Dim name = sheetNode.GetAttribute("name")
            Dim rId = sheetNode.GetAttribute("r:id")
            If String.IsNullOrWhiteSpace(name) OrElse String.IsNullOrWhiteSpace(rId) Then Continue For

            Dim target As String = Nothing
            If relMap.TryGetValue(rId, target) Then
                Dim fullPath = Path.Combine(tempDir, "xl", target.Replace("/"c, Path.DirectorySeparatorChar))
                result(name) = fullPath
            End If
        Next

        Return result
    End Function

    ''' <summary>
    ''' A custom UTF8Encoding that forces the WebName to be uppercase "UTF-8".
    ''' This is required because Excel's OOXML parser is case-sensitive for the
    ''' encoding attribute in the XML declaration of certain parts.
    ''' </summary>
    Private Class UpperCaseUTF8Encoding
        Inherits UTF8Encoding

        Public Sub New(encoderShouldEmitUTF8Identifier As Boolean)
            MyBase.New(encoderShouldEmitUTF8Identifier)
        End Sub

        Public Overrides ReadOnly Property WebName As String
            Get
                Return MyBase.WebName.ToUpperInvariant()
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Saves an XmlDocument to a file path preserving the original XML declaration
    ''' and using uppercase UTF-8 without BOM — as required by OOXML packaging.
    ''' </summary>
    Private Shared Sub APSaveXmlPreservingDeclaration(xmlDoc As System.Xml.XmlDocument, filePath As String)
        ' Capture the original XML declaration before saving
        Dim originalDecl As System.Xml.XmlDeclaration = Nothing
        If xmlDoc.FirstChild IsNot Nothing AndAlso xmlDoc.FirstChild.NodeType = System.Xml.XmlNodeType.XmlDeclaration Then
            originalDecl = DirectCast(xmlDoc.FirstChild, System.Xml.XmlDeclaration)
        End If

        ' Use XmlWriterSettings to control output precisely
        Dim settings As New System.Xml.XmlWriterSettings()
        ' Use our custom encoding to force uppercase "UTF-8"
        settings.Encoding = New UpperCaseUTF8Encoding(False) ' no BOM
        settings.Indent = False
        settings.NewLineHandling = System.Xml.NewLineHandling.None
        settings.CloseOutput = True

        ' Preserve the standalone attribute from the original declaration
        If originalDecl IsNot Nothing AndAlso Not String.IsNullOrEmpty(originalDecl.Standalone) Then
            settings.OmitXmlDeclaration = False
        Else
            settings.OmitXmlDeclaration = False
        End If

        Using writer = System.Xml.XmlWriter.Create(filePath, settings)
            ' If the original had standalone="yes", we need to write the declaration manually
            ' because XmlWriterSettings doesn't support standalone directly in .NET Framework
            If originalDecl IsNot Nothing AndAlso
               Not String.IsNullOrEmpty(originalDecl.Standalone) Then
                writer.WriteStartDocument(originalDecl.Standalone.Equals("yes", StringComparison.OrdinalIgnoreCase))
            End If

            ' Write everything except the original XmlDeclaration (WriteStartDocument already wrote one)
            For Each child As System.Xml.XmlNode In xmlDoc.ChildNodes
                If child.NodeType <> System.Xml.XmlNodeType.XmlDeclaration Then
                    child.WriteTo(writer)
                End If
            Next
        End Using
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  NON-BREAKING SPACE PRESERVATION
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Scans text for non-breaking space characters and records their surrounding
    ''' word context so they can be restored after LLM processing.
    ''' </summary>
    ''' <param name="text">The original text to scan.</param>
    ''' <returns>A list of non-breaking space occurrences with context, or Nothing if none found.</returns>
    Private Shared Function APRecordNonBreakingSpaces(text As String) As List(Of APNonBreakingSpaceInfo)
        If String.IsNullOrEmpty(text) Then Return Nothing

        Dim hasNbsp As Boolean = False
        For Each ch As Char In text
            If AP_NonBreakingSpaceChars.Contains(ch) Then
                hasNbsp = True
                Exit For
            End If
        Next
        If Not hasNbsp Then Return Nothing

        Dim results As New List(Of APNonBreakingSpaceInfo)()

        Dim i As Integer = 0
        While i < text.Length
            If AP_NonBreakingSpaceChars.Contains(text(i)) Then
                Dim info As New APNonBreakingSpaceInfo() With {
                    .SpaceChar = text(i)
                }

                ' Find word before: scan backward from i-1 to find the token
                If i > 0 Then
                    Dim wordEnd As Integer = i - 1
                    While wordEnd >= 0 AndAlso text(wordEnd) = " "c
                        wordEnd -= 1
                    End While
                    If wordEnd >= 0 Then
                        Dim wordStart As Integer = wordEnd
                        While wordStart > 0 AndAlso text(wordStart - 1) <> " "c AndAlso Not AP_NonBreakingSpaceChars.Contains(text(wordStart - 1))
                            wordStart -= 1
                        End While
                        info.WordBefore = text.Substring(wordStart, wordEnd - wordStart + 1)
                    End If
                End If

                ' Find word after: scan forward from i+1 to find the token
                If i < text.Length - 1 Then
                    Dim wordStart As Integer = i + 1
                    While wordStart < text.Length AndAlso text(wordStart) = " "c
                        wordStart += 1
                    End While
                    If wordStart < text.Length Then
                        Dim wordEnd As Integer = wordStart
                        While wordEnd < text.Length - 1 AndAlso text(wordEnd + 1) <> " "c AndAlso Not AP_NonBreakingSpaceChars.Contains(text(wordEnd + 1))
                            wordEnd += 1
                        End While
                        info.WordAfter = text.Substring(wordStart, wordEnd - wordStart + 1)
                    End If
                End If

                results.Add(info)
            End If
            i += 1
        End While

        Return If(results.Count > 0, results, Nothing)
    End Function

    ''' <summary>
    ''' Restores non-breaking spaces in text that has been processed by the LLM,
    ''' using the recorded word context to find where they should be placed.
    ''' Only restores a non-breaking space when both surrounding words match.
    ''' </summary>
    ''' <param name="text">The LLM-processed text (with regular spaces).</param>
    ''' <param name="nbspInfos">The recorded non-breaking space positions from the original text.</param>
    ''' <returns>The text with non-breaking spaces restored where context matches.</returns>
    Private Shared Function APRestoreNonBreakingSpaces(text As String, nbspInfos As List(Of APNonBreakingSpaceInfo)) As String
        If String.IsNullOrEmpty(text) OrElse nbspInfos Is Nothing OrElse nbspInfos.Count = 0 Then Return text

        Dim replacements As New SortedList(Of Integer, Char)()

        For Each info In nbspInfos
            If String.IsNullOrEmpty(info.WordBefore) OrElse String.IsNullOrEmpty(info.WordAfter) Then Continue For

            Dim pattern As String = Regex.Escape(info.WordBefore) & "( +)" & Regex.Escape(info.WordAfter)
            Dim m As Match = Regex.Match(text, pattern, RegexOptions.None)

            While m.Success
                Dim spaceGroupStart As Integer = m.Groups(1).Index
                Dim spaceGroupLen As Integer = m.Groups(1).Length

                Dim alreadyClaimed As Boolean = False
                For spIdx As Integer = spaceGroupStart To spaceGroupStart + spaceGroupLen - 1
                    If replacements.ContainsKey(spIdx) Then
                        alreadyClaimed = True
                        Exit For
                    End If
                Next

                If Not alreadyClaimed Then
                    replacements(spaceGroupStart) = info.SpaceChar
                    Exit While
                End If

                m = m.NextMatch()
            End While
        Next

        If replacements.Count = 0 Then Return text

        Dim sb As New StringBuilder(text)
        For i As Integer = replacements.Count - 1 To 0 Step -1
            Dim pos As Integer = replacements.Keys(i)
            sb(pos) = replacements.Values(i)
        Next

        Return sb.ToString()
    End Function

End Class
