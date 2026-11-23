' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.TranslatePresentations.vb
' Purpose: Translates or corrects PowerPoint documents while preserving 100% of formatting
'          by editing only DrawingML text nodes (<a:t>) inside PPTX slide parts.
'
' Architecture / Key Ideas:
'  - Called from the unified ProcessWordDocuments pipeline in ThisAddIn.TranslateDocuments.vb
'    when a .pptx file is encountered.
'  - Reuses the same batch processing, LLM communication, text redistribution, and
'    marker-based formatting logic from ThisAddIn.TranslateDocuments.vb.
'  - OpenXML Processing: Operates directly on PPTX XML and modifies only <a:t>
'    nodes, preserving styles, runs, shapes, layouts, and presentation structure.
'  - Processes all slides and notes slides.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Processes a PPTX file using OpenXML to translate or correct text nodes.
    ''' PPTX uses DrawingML namespace (a:) with a:p → a:r → a:t structure.
    ''' Called from the unified ProcessWordDocuments pipeline.
    ''' </summary>
    Private Async Function ProcessPptxOpenXml(pptxPath As String, targetLanguage As String, mode As DocumentProcessMode) As System.Threading.Tasks.Task(Of Boolean)
        Dim tempDir As String = Path.Combine(Path.GetTempPath(), $"{AN2}_pptx_{Guid.NewGuid():N}")

        Try
            ' Extract PPTX (it's a ZIP just like DOCX)
            ZipFile.ExtractToDirectory(pptxPath, tempDir)

            Dim pptDir As String = Path.Combine(tempDir, "ppt")
            If Not Directory.Exists(pptDir) Then
                ShowCustomMessageBox("Invalid PPTX structure - ppt directory not found.")
                Return False
            End If

            ' === Collect all parts to process for accurate progress tracking ===
            Dim allParts As New List(Of (FilePath As String, Label As String))()

            Dim slidesDir As String = Path.Combine(pptDir, "slides")
            If Directory.Exists(slidesDir) Then
                Dim slideFiles = Directory.GetFiles(slidesDir, "slide*.xml").
                    OrderBy(Function(f)
                                Dim m = Regex.Match(Path.GetFileNameWithoutExtension(f), "\d+")
                                Return If(m.Success, Integer.Parse(m.Value), 0)
                            End Function).ToArray()

                For Each slideFile In slideFiles
                    Dim slideNum As String = Regex.Match(Path.GetFileNameWithoutExtension(slideFile), "\d+").Value
                    Dim slideLabel As String = If(slideNum.Length > 0, $"Slide {slideNum}", Path.GetFileNameWithoutExtension(slideFile))
                    allParts.Add((slideFile, $"{Path.GetFileName(pptxPath)} ({slideLabel})"))
                Next
            End If

            Dim notesDir As String = Path.Combine(pptDir, "notesSlides")
            If Directory.Exists(notesDir) Then
                For Each notesFile In Directory.GetFiles(notesDir, "notesSlide*.xml")
                    Dim notesNum As String = Regex.Match(Path.GetFileNameWithoutExtension(notesFile), "\d+").Value
                    Dim notesLabel As String = If(notesNum.Length > 0, $"Notes {notesNum}", Path.GetFileNameWithoutExtension(notesFile))
                    allParts.Add((notesFile, $"{Path.GetFileName(pptxPath)} ({notesLabel})"))
                Next
            End If

            ' === Process all parts with progress tracking ===
            ' Save and override progress max to show per-slide progress
            Dim savedProgressMax As Integer = ProgressBarModule.GlobalProgressMax
            Dim savedProgressValue As Integer = ProgressBarModule.GlobalProgressValue
            ProgressBarModule.GlobalProgressMax = allParts.Count

            For partIdx As Integer = 0 To allParts.Count - 1
                If ProgressBarModule.CancelOperation Then Return False

                ProgressBarModule.GlobalProgressValue = partIdx
                ProgressBarModule.GlobalProgressLabel = allParts(partIdx).Label

                If Not Await ProcessPptxXmlPart(allParts(partIdx).FilePath, targetLanguage, mode, allParts(partIdx).Label) Then
                    ProgressBarModule.GlobalProgressMax = savedProgressMax
                    ProgressBarModule.GlobalProgressValue = savedProgressValue
                    Return False
                End If
            Next

            ' Restore progress for the outer file loop
            ProgressBarModule.GlobalProgressMax = savedProgressMax
            ProgressBarModule.GlobalProgressValue = savedProgressValue

            ' Repack PPTX
            File.Delete(pptxPath)
            ZipFile.CreateFromDirectory(tempDir, pptxPath, CompressionLevel.Optimal, False)

            Return True

        Finally
            If Directory.Exists(tempDir) Then
                Try : Directory.Delete(tempDir, recursive:=True) : Catch : End Try
            End If
        End Try
    End Function

    ''' <summary>
    ''' Processes a single PPTX XML part (slide, notes slide, etc.) by extracting
    ''' DrawingML paragraphs, sending text to the LLM, and writing back.
    ''' </summary>
    Private Async Function ProcessPptxXmlPart(xmlPath As String, targetLanguage As String, mode As DocumentProcessMode, fileContext As String) As System.Threading.Tasks.Task(Of Boolean)
        Try
            Dim xmlDoc As New System.Xml.XmlDocument()
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(xmlPath)

            Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
            nsMgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
            nsMgr.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main")
            nsMgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

            ' Extract paragraphs using DrawingML structure
            Dim paragraphs As List(Of TranslateParagraphInfo) = ExtractPptxParagraphsFromXml(xmlDoc, nsMgr)

            Dim processableParagraphs = paragraphs.Where(Function(p) Not p.IsEmpty).ToList()
            If processableParagraphs.Count = 0 Then Return True

            ' Process paragraphs in batches (reuses the same LLM batching as Word)
            Dim success As Boolean = Await ProcessParagraphBatches(paragraphs, targetLanguage, mode, fileContext)
            If Not success Then Return False

            ' Apply processed text back to XML nodes (reuses the same redistribution logic)
            ApplyTranslationsToXml(paragraphs)

            ' Save modified XML
            xmlDoc.Save(xmlPath)
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ProcessPptxXmlPart error for {Path.GetFileName(xmlPath)}: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Extracts paragraph information from a PPTX XML part (slide, notes, etc.).
    ''' DrawingML uses a:p → a:r → a:t structure instead of WordprocessingML's w:p → w:r → w:t.
    ''' Only extracts plain text - no formatting codes sent to LLM.
    ''' </summary>
    Private Function ExtractPptxParagraphsFromXml(xmlDoc As System.Xml.XmlDocument, nsMgr As System.Xml.XmlNamespaceManager) As List(Of TranslateParagraphInfo)
        Dim paragraphs As New List(Of TranslateParagraphInfo)()

        ' Find all a:p (paragraph) elements in DrawingML
        Dim paraNodes As System.Xml.XmlNodeList = xmlDoc.SelectNodes("//a:p", nsMgr)
        Dim paraIndex As Integer = 0

        For Each paraNode As System.Xml.XmlNode In paraNodes
            Dim paraInfo As New TranslateParagraphInfo() With {
                .Index = paraIndex,
                .TextRuns = New List(Of TranslateTextRunInfo)(),
                .TranslatedText = Nothing,
                .MarkerText = Nothing
            }

            ' Find all a:t (text) elements within a:r (run) elements in this paragraph.
            ' Skip a:t nodes inside a:fld (field) elements — these are auto-generated
            ' content like slide numbers, dates, etc. that should not be translated.
            Dim textNodes As System.Xml.XmlNodeList = paraNode.SelectNodes(".//a:r/a:t", nsMgr)
            Dim fullTextBuilder As New StringBuilder()

            ' Build a set of a:t nodes inside field elements to exclude
            Dim fieldTextNodes As New HashSet(Of System.Xml.XmlNode)()
            Dim fieldNodes As System.Xml.XmlNodeList = paraNode.SelectNodes(".//a:fld//a:t", nsMgr)
            For Each fldTextNode As System.Xml.XmlNode In fieldNodes
                fieldTextNodes.Add(fldTextNode)
            Next

            For Each textNode As System.Xml.XmlNode In textNodes
                ' Skip text inside field elements
                If fieldTextNodes.Contains(textNode) Then Continue For

                Dim text As String = textNode.InnerText

                paraInfo.TextRuns.Add(New TranslateTextRunInfo() With {
                    .TextNode = textNode,
                    .OriginalText = text,
                    .HasNoteReferenceBefore = False  ' PPTX doesn't have footnote references
                })

                fullTextBuilder.Append(text)
            Next

            paraInfo.FullText = fullTextBuilder.ToString()
            paraInfo.IsEmpty = String.IsNullOrWhiteSpace(paraInfo.FullText)

            ' Build marker-annotated text when formatting markers are enabled
            If _useFormattingMarkers AndAlso Not paraInfo.IsEmpty AndAlso paraInfo.TextRuns.Count > 1 Then
                paraInfo.MarkerText = BuildMarkerAnnotatedText(paraInfo.TextRuns)
            End If

            paragraphs.Add(paraInfo)
            paraIndex += 1
        Next

        Return paragraphs
    End Function

End Class