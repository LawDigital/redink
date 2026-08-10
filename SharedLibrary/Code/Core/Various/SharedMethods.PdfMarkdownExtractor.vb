' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.PdfMarkdownExtractor.vb
' Purpose: Non-OCR, non-Interop PDF-to-Markdown extraction using PdfPig.
'
' Design goals:
'  - No Microsoft Office/Word/PDF COM interop.
'  - No OCR and no LLM dependency.
'  - Reuse the already-shipped PdfPig dependency.
'  - Recover reading order, paragraphs, headings, lists, simple tables,
'    captions, footnotes, repeated headers/footers/page numbers, multi-column
'    layouts and common PDF text artefacts heuristically.
'  - Preserve useful typography (bold/italic) where it can be inferred safely.
'  - Return "Error: ..." for PDFs that appear image-only/scanned so the caller
'    can deliberately route them into the existing OCR pipeline.
'
' PdfPig is used unchanged under its Apache-2.0 license.
' =============================================================================

Option Explicit On
Option Strict On
Option Infer On

Namespace SharedLibrary
    Partial Public Class SharedMethods

        Public NotInheritable Class PdfMarkdownExtractor

            Private Const MinimumDocumentCharacters As System.Int32 = 24
            Private Const MinimumCharactersPerPage As System.Int32 = 8
            Private Const PositionToleranceFraction As System.Double = 0.02R
            Private Const ColumnCenterMarginFraction As System.Double = 0.04R

            Private Sub New()
            End Sub

            Public NotInheritable Class PdfMarkdownOptions
                Public Sub New()
                    Me.RemoveRepeatedPageFurniture = True
                    Me.RemoveHeaders = True
                    Me.RemoveFooters = True
                    Me.RemovePageNumbers = True
                    Me.RemoveRepeatedSideFurniture = True
                    Me.RemoveWatermarks = True
                    Me.PreserveFootnotes = True
                    Me.EmitMarkdownFootnotes = True
                    Me.PreserveBold = True
                    Me.PreserveItalic = True
                    Me.DetectCaptions = True
                    Me.DetectColumns = True
                    Me.IncludePageBreakComments = False
                    Me.DetectHeadings = True
                    Me.DetectLists = True
                    Me.DetectTables = True
                    Me.JoinWrappedLines = True
                    Me.JoinAcrossPageBreaks = True
                    Me.RemoveSoftHyphens = True
                    Me.NormalizeLigatures = True
                    Me.MaximumHeadingLevel = 6
                    Me.HeaderFooterBandFraction = 0.12R
                    Me.FootnoteBandFraction = 0.28R
                    Me.RepeatedFurnitureMinimumPages = 3
                    Me.RepeatedFurniturePageRatio = 0.5R
                    Me.TableColumnToleranceFraction = 0.018R
                End Sub

                Public Property RemoveRepeatedPageFurniture As System.Boolean
                Public Property RemoveHeaders As System.Boolean
                Public Property RemoveFooters As System.Boolean
                Public Property RemovePageNumbers As System.Boolean
                Public Property RemoveRepeatedSideFurniture As System.Boolean
                Public Property RemoveWatermarks As System.Boolean
                Public Property PreserveFootnotes As System.Boolean
                Public Property EmitMarkdownFootnotes As System.Boolean
                Public Property PreserveBold As System.Boolean
                Public Property PreserveItalic As System.Boolean
                Public Property DetectCaptions As System.Boolean
                Public Property DetectColumns As System.Boolean
                Public Property IncludePageBreakComments As System.Boolean
                Public Property DetectHeadings As System.Boolean
                Public Property DetectLists As System.Boolean
                Public Property DetectTables As System.Boolean
                Public Property JoinWrappedLines As System.Boolean
                Public Property JoinAcrossPageBreaks As System.Boolean
                Public Property RemoveSoftHyphens As System.Boolean
                Public Property NormalizeLigatures As System.Boolean
                Public Property MaximumHeadingLevel As System.Int32
                Public Property HeaderFooterBandFraction As System.Double
                Public Property FootnoteBandFraction As System.Double
                Public Property RepeatedFurnitureMinimumPages As System.Int32
                Public Property RepeatedFurniturePageRatio As System.Double
                Public Property TableColumnToleranceFraction As System.Double
            End Class

            Private Enum PdfFurnitureKind
                None = 0
                Header = 1
                Footer = 2
                PageNumber = 3
                RepeatedSideFurniture = 4
                Watermark = 5
            End Enum

            Private Enum PdfSemanticKind
                Unknown = 0
                Paragraph = 1
                Heading = 2
                ListItem = 3
                Table = 4
                Footnote = 5
                Caption = 6
            End Enum

            Private NotInheritable Class PdfDocumentModel
                Public Sub New()
                    Me.Pages = New System.Collections.Generic.List(Of PdfPageModel)()
                    Me.BodyFontSize = 0.0R
                End Sub

                Public Property Pages As System.Collections.Generic.List(Of PdfPageModel)
                Public Property BodyFontSize As System.Double
            End Class

            Private NotInheritable Class PdfPageModel
                Public Sub New()
                    Me.Blocks = New System.Collections.Generic.List(Of PdfBlockModel)()
                    Me.PageNumber = 0
                    Me.Width = 0.0R
                    Me.Height = 0.0R
                End Sub

                Public Property PageNumber As System.Int32
                Public Property Width As System.Double
                Public Property Height As System.Double
                Public Property Blocks As System.Collections.Generic.List(Of PdfBlockModel)
            End Class

            Private NotInheritable Class PdfBlockModel
                Public Sub New()
                    Me.Lines = New System.Collections.Generic.List(Of PdfLineModel)()
                    Me.ReadingOrder = -1
                    Me.BlockIndex = -1
                End Sub

                Public Property Lines As System.Collections.Generic.List(Of PdfLineModel)
                Public Property ReadingOrder As System.Int32
                Public Property BlockIndex As System.Int32
            End Class

            Private NotInheritable Class PdfLineModel
                Public Sub New()
                    Me.Words = New System.Collections.Generic.List(Of PdfWordModel)()
                    Me.Text = System.String.Empty
                    Me.Left = 0.0R
                    Me.Right = 0.0R
                    Me.Top = 0.0R
                    Me.Bottom = 0.0R
                    Me.FontSize = 0.0R
                    Me.BoldRatio = 0.0R
                    Me.ItalicRatio = 0.0R
                    Me.BlockReadingOrder = -1
                    Me.BlockIndex = -1
                    Me.LineIndexInBlock = -1
                    Me.PageNumber = 0
                    Me.FurnitureKind = PdfFurnitureKind.None
                End Sub

                Public Property Words As System.Collections.Generic.List(Of PdfWordModel)
                Public Property Text As System.String
                Public Property Left As System.Double
                Public Property Right As System.Double
                Public Property Top As System.Double
                Public Property Bottom As System.Double
                Public Property FontSize As System.Double
                Public Property BoldRatio As System.Double
                Public Property ItalicRatio As System.Double
                Public Property BlockReadingOrder As System.Int32
                Public Property BlockIndex As System.Int32
                Public Property LineIndexInBlock As System.Int32
                Public Property PageNumber As System.Int32
                Public Property FurnitureKind As PdfFurnitureKind
            End Class

            Private NotInheritable Class PdfWordModel
                Public Sub New()
                    Me.Text = System.String.Empty
                    Me.FontName = System.String.Empty
                    Me.IsBold = False
                    Me.IsItalic = False
                    Me.IsSuperscript = False
                    Me.IsSubscript = False
                End Sub

                Public Property Text As System.String
                Public Property Left As System.Double
                Public Property Right As System.Double
                Public Property Top As System.Double
                Public Property Bottom As System.Double
                Public Property FontSize As System.Double
                Public Property FontName As System.String
                Public Property IsBold As System.Boolean
                Public Property IsItalic As System.Boolean
                Public Property IsSuperscript As System.Boolean
                Public Property IsSubscript As System.Boolean
            End Class

            Private NotInheritable Class ListMarkerInfo
                Public Sub New()
                    Me.MarkdownMarker = System.String.Empty
                    Me.PrefixLength = 0
                    Me.MarkerOnly = False
                    Me.IsOrdered = False
                    Me.RawMarker = System.String.Empty
                End Sub

                Public Property MarkdownMarker As System.String
                Public Property PrefixLength As System.Int32
                Public Property MarkerOnly As System.Boolean
                Public Property IsOrdered As System.Boolean
                Public Property RawMarker As System.String
            End Class

            Private NotInheritable Class TableCandidate
                Public Sub New()
                    Me.Rows = New System.Collections.Generic.List(Of System.Collections.Generic.List(Of System.String))()
                    Me.ColumnStarts = New System.Collections.Generic.List(Of System.Double)()
                    Me.ConsumedLineCount = 0
                End Sub

                Public Property Rows As System.Collections.Generic.List(Of System.Collections.Generic.List(Of System.String))
                Public Property ColumnStarts As System.Collections.Generic.List(Of System.Double)
                Public Property ConsumedLineCount As System.Int32
            End Class

            Private NotInheritable Class SemanticElement
                Public Sub New()
                    Me.Kind = PdfSemanticKind.Unknown
                    Me.Lines = New System.Collections.Generic.List(Of PdfLineModel)()
                    Me.Text = System.String.Empty
                    Me.HeadingLevel = 0
                    Me.ListMarker = System.String.Empty
                    Me.ListIndentLevel = 0
                    Me.PageNumber = 0
                    Me.Table = Nothing
                    Me.FootnoteId = System.String.Empty
                    Me.FootnoteMarker = System.String.Empty
                End Sub

                Public Property Kind As PdfSemanticKind
                Public Property Lines As System.Collections.Generic.List(Of PdfLineModel)
                Public Property Text As System.String
                Public Property HeadingLevel As System.Int32
                Public Property ListMarker As System.String
                Public Property ListIndentLevel As System.Int32
                Public Property PageNumber As System.Int32
                Public Property Table As TableCandidate
                Public Property FootnoteId As System.String
                Public Property FootnoteMarker As System.String
            End Class

            ''' <summary>
            ''' Reads a text-based PDF and returns Markdown. This method never invokes OCR.
            ''' Image-only/scanned PDFs return an Error: value so callers can choose the OCR path.
            ''' </summary>
            Public Shared Function ReadPdfAsMarkdown(
                pdfPath As System.String,
                Optional options As PdfMarkdownOptions = Nothing
            ) As System.String
                If System.String.IsNullOrWhiteSpace(pdfPath) OrElse Not System.IO.File.Exists(pdfPath) Then
                    Return "Error: File not found."
                End If

                If options Is Nothing Then
                    options = New PdfMarkdownOptions()
                End If

                Try
                    Dim model As PdfDocumentModel = ExtractDocumentModel(pdfPath, options)
                    Dim totalCharacters As System.Int32 = CountVisibleCharacters(model)
                    Dim minimumCharacters As System.Int32 = System.Math.Max(
                        MinimumDocumentCharacters,
                        model.Pages.Count * MinimumCharactersPerPage
                    )

                    If totalCharacters < minimumCharacters Then
                        Return "Error: PDF contains too little extractable text and likely requires OCR."
                    End If

                    model.BodyFontSize = DetermineBodyFontSize(model)
                    MarkPageNumbers(model, options)

                    If options.RemoveRepeatedPageFurniture AndAlso model.Pages.Count >= options.RepeatedFurnitureMinimumPages Then
                        MarkRepeatedPageFurniture(model, options)
                    End If

                    MarkSuperscriptAndSubscriptWords(model)

                    Dim footnoteIds As System.Collections.Generic.Dictionary(Of System.String, System.String) =
                        BuildFootnoteIdMap(model, options)

                    Dim elements As System.Collections.Generic.List(Of SemanticElement) =
                        BuildSemanticElements(model, options, footnoteIds)

                    MergeCompatibleSemanticElements(elements, model, options)

                    Dim markdown As System.String = RenderSemanticElements(elements, model, options, footnoteIds)
                    markdown = NormalizeMarkdown(markdown, options)

                    If System.String.IsNullOrWhiteSpace(markdown) Then
                        Return "Error: No text content found in PDF."
                    End If

                    Return markdown.TrimEnd()
                Catch ex As System.Exception
                    Return "Error reading PDF as Markdown: " & ex.Message
                End Try
            End Function

            ''' <summary>
            ''' Convenience helper that writes UTF-8 without BOM and returns the output path.
            ''' </summary>
            Public Shared Function WritePdfMarkdownFile(
                pdfPath As System.String,
                markdownFilePath As System.String,
                Optional options As PdfMarkdownOptions = Nothing
            ) As System.String
                If System.String.IsNullOrWhiteSpace(markdownFilePath) Then
                    Return "Error: Output Markdown path is empty."
                End If

                Try
                    Dim markdown As System.String = ReadPdfAsMarkdown(pdfPath, options)
                    If markdown.StartsWith("Error:", System.StringComparison.Ordinal) Then
                        Return markdown
                    End If

                    Dim outputDirectory As System.String = System.IO.Path.GetDirectoryName(markdownFilePath)
                    If Not System.String.IsNullOrWhiteSpace(outputDirectory) AndAlso Not System.IO.Directory.Exists(outputDirectory) Then
                        System.IO.Directory.CreateDirectory(outputDirectory)
                    End If

                    System.IO.File.WriteAllText(
                        markdownFilePath,
                        markdown,
                        New System.Text.UTF8Encoding(False)
                    )

                    Return markdownFilePath
                Catch ex As System.Exception
                    Return "Error writing Markdown file: " & ex.Message
                End Try
            End Function

            Private Shared Function ExtractDocumentModel(
                pdfPath As System.String,
                options As PdfMarkdownOptions
            ) As PdfDocumentModel
                Dim model As New PdfDocumentModel()

                Using document As UglyToad.PdfPig.PdfDocument = UglyToad.PdfPig.PdfDocument.Open(pdfPath)
                    For pageNumber As System.Int32 = 1 To document.NumberOfPages
                        Dim page = document.GetPage(pageNumber)
                        Dim pageModel As New PdfPageModel()
                        pageModel.PageNumber = pageNumber
                        pageModel.Width = System.Convert.ToDouble(page.Width, System.Globalization.CultureInfo.InvariantCulture)
                        pageModel.Height = System.Convert.ToDouble(page.Height, System.Globalization.CultureInfo.InvariantCulture)

                        Dim wordExtractor = UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor.NearestNeighbourWordExtractor.Instance
                        Dim words = wordExtractor.GetWords(page.Letters)
                        Dim pageSegmenter = UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter.DocstrumBoundingBoxes.Instance
                        Dim blocks = pageSegmenter.GetBlocks(words)
                        Dim readingOrderDetector = UglyToad.PdfPig.DocumentLayoutAnalysis.ReadingOrderDetector.UnsupervisedReadingOrderDetector.Instance
                        Dim orderedBlocks = readingOrderDetector.Get(blocks)

                        Dim blockIndex As System.Int32 = 0
                        For Each block In orderedBlocks
                            Dim blockModel As New PdfBlockModel()
                            blockModel.ReadingOrder = block.ReadingOrder
                            blockModel.BlockIndex = blockIndex

                            Dim lineIndex As System.Int32 = 0
                            For Each textLine In block.TextLines
                                Dim lineModel As PdfLineModel = ConvertTextLine(textLine, options)
                                If lineModel IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(lineModel.Text) Then
                                    lineModel.BlockReadingOrder = blockModel.ReadingOrder
                                    lineModel.BlockIndex = blockIndex
                                    lineModel.LineIndexInBlock = lineIndex
                                    lineModel.PageNumber = pageNumber
                                    blockModel.Lines.Add(lineModel)
                                    lineIndex += 1
                                End If
                            Next

                            If blockModel.Lines.Count > 0 Then
                                pageModel.Blocks.Add(blockModel)
                                blockIndex += 1
                            End If
                        Next

                        model.Pages.Add(pageModel)
                    Next
                End Using

                Return model
            End Function

            Private Shared Function ConvertTextLine(
                textLine As UglyToad.PdfPig.DocumentLayoutAnalysis.TextLine,
                options As PdfMarkdownOptions
            ) As PdfLineModel
                If textLine Is Nothing OrElse textLine.Words Is Nothing OrElse textLine.Words.Count = 0 Then
                    Return Nothing
                End If

                Dim result As New PdfLineModel()
                result.Text = NormalizeBasicText(textLine.Text).Trim()
                result.Left = textLine.BoundingBox.BottomLeft.X
                result.Right = textLine.BoundingBox.BottomLeft.X + textLine.BoundingBox.Width
                result.Bottom = textLine.BoundingBox.BottomLeft.Y
                result.Top = textLine.BoundingBox.BottomLeft.Y + textLine.BoundingBox.Height

                Dim fontSizes As New System.Collections.Generic.List(Of System.Double)()
                Dim boldCharacters As System.Int32 = 0
                Dim italicCharacters As System.Int32 = 0
                Dim totalCharacters As System.Int32 = 0

                For Each word As UglyToad.PdfPig.Content.Word In textLine.Words
                    If System.String.IsNullOrWhiteSpace(word.Text) Then
                        Continue For
                    End If

                    Dim wordModel As New PdfWordModel()
                    wordModel.Text = NormalizeBasicText(word.Text)
                    wordModel.Left = word.BoundingBox.BottomLeft.X
                    wordModel.Right = word.BoundingBox.BottomLeft.X + word.BoundingBox.Width
                    wordModel.Bottom = word.BoundingBox.BottomLeft.Y
                    wordModel.Top = word.BoundingBox.BottomLeft.Y + word.BoundingBox.Height
                    wordModel.FontName = If(word.FontName, System.String.Empty)
                    wordModel.IsBold = IsBoldFontName(wordModel.FontName)
                    wordModel.IsItalic = IsItalicFontName(wordModel.FontName)

                    Dim wordFontSizes As New System.Collections.Generic.List(Of System.Double)()
                    For Each letter As UglyToad.PdfPig.Content.Letter In word.Letters
                        If letter.FontSize > 0.0R Then
                            wordFontSizes.Add(letter.FontSize)
                            fontSizes.Add(letter.FontSize)
                        End If
                    Next
                    wordModel.FontSize = Median(wordFontSizes)
                    result.Words.Add(wordModel)

                    Dim characterCount As System.Int32 = wordModel.Text.Length
                    totalCharacters += characterCount
                    If wordModel.IsBold Then
                        boldCharacters += characterCount
                    End If
                    If wordModel.IsItalic Then
                        italicCharacters += characterCount
                    End If
                Next

                result.FontSize = Median(fontSizes)
                If totalCharacters > 0 Then
                    result.BoldRatio = boldCharacters / System.Convert.ToDouble(totalCharacters, System.Globalization.CultureInfo.InvariantCulture)
                    result.ItalicRatio = italicCharacters / System.Convert.ToDouble(totalCharacters, System.Globalization.CultureInfo.InvariantCulture)
                End If

                If result.Words.Count = 0 OrElse System.String.IsNullOrWhiteSpace(result.Text) Then
                    Return Nothing
                End If

                Return result
            End Function

            Private Shared Function IsBoldFontName(fontName As System.String) As System.Boolean
                If System.String.IsNullOrWhiteSpace(fontName) Then
                    Return False
                End If

                Dim normalized As System.String = fontName.ToLowerInvariant()
                Return normalized.Contains("bold") OrElse
                    normalized.Contains("black") OrElse
                    normalized.Contains("heavy") OrElse
                    normalized.Contains("semibold") OrElse
                    normalized.Contains("demi")
            End Function

            Private Shared Function IsItalicFontName(fontName As System.String) As System.Boolean
                If System.String.IsNullOrWhiteSpace(fontName) Then
                    Return False
                End If

                Dim normalized As System.String = fontName.ToLowerInvariant()
                Return normalized.Contains("italic") OrElse
                    normalized.Contains("oblique") OrElse
                    normalized.Contains("kursiv")
            End Function

            Private Shared Sub MarkSuperscriptAndSubscriptWords(model As PdfDocumentModel)
                For Each page As PdfPageModel In model.Pages
                    For Each block As PdfBlockModel In page.Blocks
                        For Each line As PdfLineModel In block.Lines
                            If line.FontSize <= 0.0R Then
                                Continue For
                            End If

                            Dim baseline As System.Double = MedianWordBottom(line.Words)
                            For Each word As PdfWordModel In line.Words
                                If word.FontSize <= 0.0R OrElse word.FontSize >= line.FontSize * 0.92R Then
                                    Continue For
                                End If

                                Dim offset As System.Double = word.Bottom - baseline
                                If offset > line.FontSize * 0.18R Then
                                    word.IsSuperscript = True
                                ElseIf offset < -(line.FontSize * 0.12R) Then
                                    word.IsSubscript = True
                                End If
                            Next
                        Next
                    Next
                Next
            End Sub

            Private Shared Function MedianWordBottom(words As System.Collections.Generic.List(Of PdfWordModel)) As System.Double
                Dim values As New System.Collections.Generic.List(Of System.Double)()
                For Each word As PdfWordModel In words
                    values.Add(word.Bottom)
                Next
                Return Median(values)
            End Function

            Private Shared Function CountVisibleCharacters(model As PdfDocumentModel) As System.Int32
                Dim total As System.Int32 = 0
                For Each page As PdfPageModel In model.Pages
                    For Each block As PdfBlockModel In page.Blocks
                        For Each line As PdfLineModel In block.Lines
                            For Each character As System.Char In line.Text
                                If Not System.Char.IsWhiteSpace(character) Then
                                    total += 1
                                End If
                            Next
                        Next
                    Next
                Next
                Return total
            End Function

            Private Shared Function DetermineBodyFontSize(model As PdfDocumentModel) As System.Double
                Dim weightedBuckets As New System.Collections.Generic.Dictionary(Of System.Int32, System.Int64)()

                For Each page As PdfPageModel In model.Pages
                    For Each block As PdfBlockModel In page.Blocks
                        For Each line As PdfLineModel In block.Lines
                            If line.FontSize <= 0.0R Then
                                Continue For
                            End If
                            If IsLikelyPageNumberText(line.Text) Then
                                Continue For
                            End If
                            If line.Text.Length < 8 Then
                                Continue For
                            End If

                            Dim bucket As System.Int32 = System.Convert.ToInt32(System.Math.Round(line.FontSize * 4.0R))
                            Dim weight As System.Int64 = System.Math.Max(1, CountNonWhitespaceCharacters(line.Text))
                            If Not weightedBuckets.ContainsKey(bucket) Then
                                weightedBuckets(bucket) = 0
                            End If
                            weightedBuckets(bucket) += weight
                        Next
                    Next
                Next

                If weightedBuckets.Count = 0 Then
                    Return 0.0R
                End If

                Dim bestBucket As System.Int32 = 0
                Dim bestWeight As System.Int64 = -1
                For Each pair As System.Collections.Generic.KeyValuePair(Of System.Int32, System.Int64) In weightedBuckets
                    If pair.Value > bestWeight Then
                        bestWeight = pair.Value
                        bestBucket = pair.Key
                    End If
                Next

                Return bestBucket / 4.0R
            End Function

            Private Shared Function CountNonWhitespaceCharacters(value As System.String) As System.Int32
                Dim total As System.Int32 = 0
                If System.String.IsNullOrEmpty(value) Then
                    Return total
                End If
                For Each character As System.Char In value
                    If Not System.Char.IsWhiteSpace(character) Then
                        total += 1
                    End If
                Next
                Return total
            End Function

            Private Shared Sub MarkPageNumbers(model As PdfDocumentModel, options As PdfMarkdownOptions)
                If Not options.RemovePageNumbers Then
                    Return
                End If

                For Each page As PdfPageModel In model.Pages
                    For Each block As PdfBlockModel In page.Blocks
                        For Each line As PdfLineModel In block.Lines
                            If IsLikelyPageNumber(line, page) Then
                                line.FurnitureKind = PdfFurnitureKind.PageNumber
                            End If
                        Next
                    Next
                Next
            End Sub

            Private Shared Function IsLikelyPageNumber(line As PdfLineModel, page As PdfPageModel) As System.Boolean
                If line Is Nothing OrElse page Is Nothing Then
                    Return False
                End If

                If Not IsLikelyPageNumberText(line.Text) Then
                    Return False
                End If

                If page.Height <= 0.0R Then
                    Return False
                End If

                Dim bandHeight As System.Double = page.Height * 0.18R
                Dim inTopBand As System.Boolean = line.Top >= page.Height - bandHeight
                Dim inBottomBand As System.Boolean = line.Bottom <= bandHeight
                Return inTopBand OrElse inBottomBand
            End Function

            Private Shared Function IsLikelyPageNumberText(text As System.String) As System.Boolean
                If System.String.IsNullOrWhiteSpace(text) Then
                    Return False
                End If

                Return System.Text.RegularExpressions.Regex.IsMatch(
                    text.Trim(),
                    "^(?:(?:page|seite|s\.)\s*)?[-–—]?\s*\d{1,5}\s*(?:(?:/|of|von)\s*\d{1,5})?\s*[-–—]?$",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                )
            End Function

            Private Shared Sub MarkRepeatedPageFurniture(
                model As PdfDocumentModel,
                options As PdfMarkdownOptions
            )
                Dim occurrences As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.HashSet(Of System.Int32))(System.StringComparer.OrdinalIgnoreCase)
                Dim locations As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of PdfLineModel))(System.StringComparer.OrdinalIgnoreCase)

                For Each page As PdfPageModel In model.Pages
                    For Each block As PdfBlockModel In page.Blocks
                        For Each line As PdfLineModel In block.Lines
                            If line.FurnitureKind = PdfFurnitureKind.PageNumber Then
                                Continue For
                            End If

                            Dim key As System.String = NormalizeFurnitureKey(line.Text)
                            If key.Length < 2 Then
                                Continue For
                            End If

                            Dim pages As System.Collections.Generic.HashSet(Of System.Int32) = Nothing
                            If Not occurrences.TryGetValue(key, pages) Then
                                pages = New System.Collections.Generic.HashSet(Of System.Int32)()
                                occurrences(key) = pages
                                locations(key) = New System.Collections.Generic.List(Of PdfLineModel)()
                            End If
                            pages.Add(page.PageNumber)
                            locations(key).Add(line)
                        Next
                    Next
                Next

                Dim requiredCount As System.Int32 = System.Math.Max(
                    options.RepeatedFurnitureMinimumPages,
                    System.Convert.ToInt32(System.Math.Ceiling(model.Pages.Count * options.RepeatedFurniturePageRatio))
                )

                For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Collections.Generic.HashSet(Of System.Int32)) In occurrences
                    If pair.Value.Count < requiredCount Then
                        Continue For
                    End If

                    Dim repeatedLines As System.Collections.Generic.List(Of PdfLineModel) = locations(pair.Key)
                    For Each line As PdfLineModel In repeatedLines
                        Dim page As PdfPageModel = model.Pages(line.PageNumber - 1)
                        Dim bandHeight As System.Double = page.Height * System.Math.Max(0.02R, System.Math.Min(0.3R, options.HeaderFooterBandFraction))
                        Dim inTopBand As System.Boolean = line.Top >= page.Height - bandHeight
                        Dim inBottomBand As System.Boolean = line.Bottom <= bandHeight
                        Dim nearSide As System.Boolean = line.Left <= page.Width * 0.06R OrElse line.Right >= page.Width * 0.94R
                        Dim unusuallyLarge As System.Boolean = model.BodyFontSize > 0.0R AndAlso line.FontSize >= model.BodyFontSize * 1.45R
                        Dim shortText As System.Boolean = line.Text.Trim().Length <= 100

                        If inTopBand AndAlso options.RemoveHeaders Then
                            line.FurnitureKind = PdfFurnitureKind.Header
                        ElseIf inBottomBand AndAlso options.RemoveFooters Then
                            line.FurnitureKind = PdfFurnitureKind.Footer
                        ElseIf nearSide AndAlso shortText AndAlso options.RemoveRepeatedSideFurniture Then
                            line.FurnitureKind = PdfFurnitureKind.RepeatedSideFurniture
                        ElseIf unusuallyLarge AndAlso shortText AndAlso options.RemoveWatermarks Then
                            line.FurnitureKind = PdfFurnitureKind.Watermark
                        End If
                    Next
                Next
            End Sub

            Private Shared Function NormalizeFurnitureKey(value As System.String) As System.String
                Dim result As System.String = NormalizeBasicText(value)
                result = System.Text.RegularExpressions.Regex.Replace(result, "\d+", "#")
                result = System.Text.RegularExpressions.Regex.Replace(result, "\s+", " ").Trim().ToLowerInvariant()
                Return result
            End Function

            Private Shared Function GetOrderedPageLines(page As PdfPageModel, options As PdfMarkdownOptions) As System.Collections.Generic.List(Of PdfLineModel)
                Dim lines As New System.Collections.Generic.List(Of PdfLineModel)()
                For Each block As PdfBlockModel In page.Blocks
                    For Each line As PdfLineModel In block.Lines
                        If ShouldIncludeLine(line, options) Then
                            lines.Add(line)
                        End If
                    Next
                Next

                ' PdfPig often emits a bullet/number marker as a separate text block from the
                ' first line of the item. Associate those fragments geometrically BEFORE any
                ' reading-order or semantic processing. This is much more reliable than hoping
                ' that a global sort happens to place marker and text next to one another.
                If options.DetectLists Then
                    lines = AttachDetachedListMarkers(lines, page)
                End If

                ' PdfPig block reading order is useful as metadata, but it is not reliable enough
                ' for detached list markers. A bullet glyph is frequently emitted as its own block
                ' and can otherwise be placed several paragraphs away from the associated text.
                ' For ordinary pages, reconstruct the visual reading order geometrically.
                If Not options.DetectColumns OrElse page.Width <= 0.0R OrElse lines.Count < 8 Then
                    Return OrderLinesByVisualRows(lines)
                End If

                If Not LooksLikeMultiColumnPage(lines, page) Then
                    Return OrderLinesByVisualRows(lines)
                End If

                Return OrderLinesByColumnsAndSpanningRows(lines, page)
            End Function

            Private Shared Function AttachDetachedListMarkers(
                sourceLines As System.Collections.Generic.List(Of PdfLineModel),
                page As PdfPageModel
            ) As System.Collections.Generic.List(Of PdfLineModel)
                Dim result As New System.Collections.Generic.List(Of PdfLineModel)(sourceLines)
                If result.Count < 2 Then
                    Return result
                End If

                Dim consumed As New System.Collections.Generic.HashSet(Of PdfLineModel)()
                Dim synthetic As New System.Collections.Generic.List(Of PdfLineModel)()

                For Each markerLine As PdfLineModel In sourceLines
                    If consumed.Contains(markerLine) Then
                        Continue For
                    End If

                    Dim markerInfo As ListMarkerInfo = Nothing
                    If Not TryGetListMarker(markerLine.Text, markerInfo) OrElse Not markerInfo.MarkerOnly Then
                        Continue For
                    End If

                    Dim markerFontSize As System.Double = markerLine.FontSize
                    If markerFontSize <= 0.0R Then
                        markerFontSize = 10.0R
                    End If

                    Dim baselineTolerance As System.Double = System.Math.Max(1.5R, markerFontSize * 0.45R)
                    Dim maximumHorizontalGap As System.Double
                    If page.Width > 0.0R Then
                        maximumHorizontalGap = System.Math.Max(markerFontSize * 5.0R, page.Width * 0.14R)
                    Else
                        maximumHorizontalGap = markerFontSize * 8.0R
                    End If

                    Dim bestCandidate As PdfLineModel = Nothing
                    Dim bestScore As System.Double = System.Double.MaxValue

                    For Each candidate As PdfLineModel In sourceLines
                        If candidate Is markerLine OrElse consumed.Contains(candidate) Then
                            Continue For
                        End If
                        If candidate.PageNumber <> markerLine.PageNumber Then
                            Continue For
                        End If

                        Dim candidateMarker As ListMarkerInfo = Nothing
                        If TryGetListMarker(candidate.Text, candidateMarker) AndAlso candidateMarker.MarkerOnly Then
                            Continue For
                        End If

                        ' The item text must be to the right of the detached marker.
                        If candidate.Left <= markerLine.Left Then
                            Continue For
                        End If

                        Dim baselineDifference As System.Double = System.Math.Abs(candidate.Bottom - markerLine.Bottom)
                        Dim candidateTolerance As System.Double = System.Math.Max(
                            baselineTolerance,
                            System.Math.Max(1.5R, candidate.FontSize * 0.45R)
                        )
                        If baselineDifference > candidateTolerance Then
                            Continue For
                        End If

                        Dim horizontalGap As System.Double = candidate.Left - markerLine.Right
                        If horizontalGap > maximumHorizontalGap Then
                            Continue For
                        End If

                        ' Baseline agreement is more important than horizontal distance.
                        Dim score As System.Double = baselineDifference * 100.0R + System.Math.Max(0.0R, horizontalGap)
                        If score < bestScore Then
                            bestScore = score
                            bestCandidate = candidate
                        End If
                    Next

                    If bestCandidate Is Nothing Then
                        Continue For
                    End If

                    Dim combined As New PdfLineModel()
                    combined.Text = markerInfo.RawMarker.Trim() & " " & bestCandidate.Text.TrimStart()
                    combined.Left = System.Math.Min(markerLine.Left, bestCandidate.Left)
                    combined.Right = System.Math.Max(markerLine.Right, bestCandidate.Right)
                    combined.Top = System.Math.Max(markerLine.Top, bestCandidate.Top)
                    combined.Bottom = bestCandidate.Bottom
                    combined.FontSize = bestCandidate.FontSize
                    combined.BoldRatio = bestCandidate.BoldRatio
                    combined.ItalicRatio = bestCandidate.ItalicRatio
                    combined.BlockReadingOrder = bestCandidate.BlockReadingOrder
                    combined.BlockIndex = bestCandidate.BlockIndex
                    combined.LineIndexInBlock = bestCandidate.LineIndexInBlock
                    combined.PageNumber = bestCandidate.PageNumber
                    combined.FurnitureKind = PdfFurnitureKind.None

                    For Each word As PdfWordModel In markerLine.Words
                        combined.Words.Add(word)
                    Next
                    For Each word As PdfWordModel In bestCandidate.Words
                        combined.Words.Add(word)
                    Next

                    consumed.Add(markerLine)
                    consumed.Add(bestCandidate)
                    synthetic.Add(combined)
                Next

                If consumed.Count = 0 Then
                    Return result
                End If

                result.Clear()
                For Each line As PdfLineModel In sourceLines
                    If Not consumed.Contains(line) Then
                        result.Add(line)
                    End If
                Next
                result.AddRange(synthetic)

                Return result
            End Function

            Private Shared Function ShouldIncludeLine(line As PdfLineModel, options As PdfMarkdownOptions) As System.Boolean
                If line Is Nothing OrElse System.String.IsNullOrWhiteSpace(line.Text) Then
                    Return False
                End If

                Select Case line.FurnitureKind
                    Case PdfFurnitureKind.PageNumber
                        Return Not options.RemovePageNumbers
                    Case PdfFurnitureKind.Header
                        Return Not options.RemoveHeaders
                    Case PdfFurnitureKind.Footer
                        Return Not options.RemoveFooters
                    Case PdfFurnitureKind.RepeatedSideFurniture
                        Return Not options.RemoveRepeatedSideFurniture
                    Case PdfFurnitureKind.Watermark
                        Return Not options.RemoveWatermarks
                    Case Else
                        Return True
                End Select
            End Function

            Private Shared Function CompareLinesByExistingReadingOrder(x As PdfLineModel, y As PdfLineModel) As System.Int32
                Dim orderCompare As System.Int32 = x.BlockReadingOrder.CompareTo(y.BlockReadingOrder)
                If orderCompare <> 0 Then
                    Return orderCompare
                End If
                Dim blockCompare As System.Int32 = x.BlockIndex.CompareTo(y.BlockIndex)
                If blockCompare <> 0 Then
                    Return blockCompare
                End If
                Return x.LineIndexInBlock.CompareTo(y.LineIndexInBlock)
            End Function

            Private Shared Function CompareLinesStrictBaselineTopDown(x As PdfLineModel, y As PdfLineModel) As System.Int32
                Dim baselineCompare As System.Int32 = y.Bottom.CompareTo(x.Bottom)
                If baselineCompare <> 0 Then
                    Return baselineCompare
                End If

                Dim leftCompare As System.Int32 = x.Left.CompareTo(y.Left)
                If leftCompare <> 0 Then
                    Return leftCompare
                End If

                Return CompareLinesByExistingReadingOrder(x, y)
            End Function

            Private Shared Function CompareLinesLeftToRight(x As PdfLineModel, y As PdfLineModel) As System.Int32
                Dim leftCompare As System.Int32 = x.Left.CompareTo(y.Left)
                If leftCompare <> 0 Then
                    Return leftCompare
                End If

                Dim baselineCompare As System.Int32 = y.Bottom.CompareTo(x.Bottom)
                If baselineCompare <> 0 Then
                    Return baselineCompare
                End If

                Return CompareLinesByExistingReadingOrder(x, y)
            End Function

            Private Shared Function OrderLinesByVisualRows(
                sourceLines As System.Collections.Generic.List(Of PdfLineModel)
            ) As System.Collections.Generic.List(Of PdfLineModel)
                Dim result As New System.Collections.Generic.List(Of PdfLineModel)()
                If sourceLines Is Nothing OrElse sourceLines.Count = 0 Then
                    Return result
                End If

                Dim sorted As New System.Collections.Generic.List(Of PdfLineModel)(sourceLines)
                sorted.Sort(AddressOf CompareLinesStrictBaselineTopDown)

                Dim fontSizes As New System.Collections.Generic.List(Of System.Double)()
                For Each line As PdfLineModel In sorted
                    If line.FontSize > 0.0R Then
                        fontSizes.Add(line.FontSize)
                    End If
                Next

                Dim typicalFontSize As System.Double = Median(fontSizes)
                If typicalFontSize <= 0.0R Then
                    typicalFontSize = 10.0R
                End If

                ' IMPORTANT: this tolerance is fixed for the whole page. Do not use a
                ' pair-dependent comparator here: it violates transitivity and can cause
                ' List.Sort to move several detached bullet glyphs ahead of their text.
                Dim rowTolerance As System.Double = System.Math.Max(1.25R, typicalFontSize * 0.28R)

                Dim currentRow As New System.Collections.Generic.List(Of PdfLineModel)()
                Dim rowAnchorBaseline As System.Double = sorted(0).Bottom

                For Each line As PdfLineModel In sorted
                    If currentRow.Count = 0 Then
                        currentRow.Add(line)
                        rowAnchorBaseline = line.Bottom
                        Continue For
                    End If

                    If System.Math.Abs(line.Bottom - rowAnchorBaseline) <= rowTolerance Then
                        currentRow.Add(line)
                    Else
                        currentRow.Sort(AddressOf CompareLinesLeftToRight)
                        result.AddRange(currentRow)
                        currentRow = New System.Collections.Generic.List(Of PdfLineModel)()
                        currentRow.Add(line)
                        rowAnchorBaseline = line.Bottom
                    End If
                Next

                If currentRow.Count > 0 Then
                    currentRow.Sort(AddressOf CompareLinesLeftToRight)
                    result.AddRange(currentRow)
                End If

                Return result
            End Function

            Private Shared Function LooksLikeMultiColumnPage(
                lines As System.Collections.Generic.List(Of PdfLineModel),
                page As PdfPageModel
            ) As System.Boolean
                Dim center As System.Double = page.Width / 2.0R
                Dim margin As System.Double = page.Width * ColumnCenterMarginFraction
                Dim leftCount As System.Int32 = 0
                Dim rightCount As System.Int32 = 0
                Dim spanningCount As System.Int32 = 0

                For Each line As PdfLineModel In lines
                    If line.Right < center - margin Then
                        leftCount += 1
                    ElseIf line.Left > center + margin Then
                        rightCount += 1
                    Else
                        spanningCount += 1
                    End If
                Next

                Return leftCount >= 4 AndAlso rightCount >= 4 AndAlso spanningCount <= System.Math.Max(6, lines.Count \ 3)
            End Function

            Private Shared Function OrderLinesByColumnsAndSpanningRows(
                lines As System.Collections.Generic.List(Of PdfLineModel),
                page As PdfPageModel
            ) As System.Collections.Generic.List(Of PdfLineModel)
                Dim result As New System.Collections.Generic.List(Of PdfLineModel)()
                Dim center As System.Double = page.Width / 2.0R
                Dim margin As System.Double = page.Width * ColumnCenterMarginFraction

                Dim spanning As New System.Collections.Generic.List(Of PdfLineModel)()
                For Each line As PdfLineModel In lines
                    If line.Left <= center + margin AndAlso line.Right >= center - margin Then
                        spanning.Add(line)
                    End If
                Next
                spanning = OrderLinesByVisualRows(spanning)

                Dim upperBoundary As System.Double = page.Height + 1.0R
                For spanIndex As System.Int32 = 0 To spanning.Count
                    Dim lowerBoundary As System.Double
                    If spanIndex < spanning.Count Then
                        lowerBoundary = spanning(spanIndex).Top
                    Else
                        lowerBoundary = -1.0R
                    End If

                    AppendColumnBand(lines, result, center, margin, upperBoundary, lowerBoundary)

                    If spanIndex < spanning.Count Then
                        If Not result.Contains(spanning(spanIndex)) Then
                            result.Add(spanning(spanIndex))
                        End If
                        upperBoundary = spanning(spanIndex).Bottom
                    End If
                Next

                For Each line As PdfLineModel In lines
                    If Not result.Contains(line) Then
                        result.Add(line)
                    End If
                Next

                Return result
            End Function

            Private Shared Sub AppendColumnBand(
                lines As System.Collections.Generic.List(Of PdfLineModel),
                result As System.Collections.Generic.List(Of PdfLineModel),
                center As System.Double,
                margin As System.Double,
                upperBoundary As System.Double,
                lowerBoundary As System.Double
            )
                Dim leftLines As New System.Collections.Generic.List(Of PdfLineModel)()
                Dim rightLines As New System.Collections.Generic.List(Of PdfLineModel)()

                For Each line As PdfLineModel In lines
                    If result.Contains(line) Then
                        Continue For
                    End If
                    If line.Top > upperBoundary OrElse line.Top <= lowerBoundary Then
                        Continue For
                    End If
                    If line.Right < center - margin Then
                        leftLines.Add(line)
                    ElseIf line.Left > center + margin Then
                        rightLines.Add(line)
                    End If
                Next

                leftLines = OrderLinesByVisualRows(leftLines)
                rightLines = OrderLinesByVisualRows(rightLines)
                result.AddRange(leftLines)
                result.AddRange(rightLines)
            End Sub

            Private Shared Function CompareLinesTopDown(x As PdfLineModel, y As PdfLineModel) As System.Int32
                Dim topCompare As System.Int32 = y.Top.CompareTo(x.Top)
                If topCompare <> 0 Then
                    Return topCompare
                End If
                Return x.Left.CompareTo(y.Left)
            End Function

            Private Shared Function BuildFootnoteIdMap(
                model As PdfDocumentModel,
                options As PdfMarkdownOptions
            ) As System.Collections.Generic.Dictionary(Of System.String, System.String)
                Dim result As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
                If Not options.PreserveFootnotes Then
                    Return result
                End If

                Dim usedIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)

                For Each page As PdfPageModel In model.Pages
                    Dim lines As System.Collections.Generic.List(Of PdfLineModel) = GetOrderedPageLines(page, options)
                    For Each line As PdfLineModel In lines
                        Dim marker As System.String = System.String.Empty
                        Dim content As System.String = System.String.Empty
                        If TryParseFootnoteStart(line, page, model.BodyFontSize, options, marker, content) Then
                            Dim key As System.String = FootnoteMapKey(page.PageNumber, marker)
                            If Not result.ContainsKey(key) Then
                                Dim baseId As System.String = SanitizeFootnoteId(marker)
                                If System.String.IsNullOrWhiteSpace(baseId) Then
                                    baseId = "fn"
                                End If
                                Dim candidate As System.String = baseId
                                If usedIds.Contains(candidate) Then
                                    candidate = "p" & page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) & "-" & baseId
                                End If
                                Dim suffix As System.Int32 = 2
                                While usedIds.Contains(candidate)
                                    candidate = "p" & page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) & "-" & baseId & "-" & suffix.ToString(System.Globalization.CultureInfo.InvariantCulture)
                                    suffix += 1
                                End While
                                usedIds.Add(candidate)
                                result(key) = candidate
                            End If
                        End If
                    Next
                Next

                Return result
            End Function

            Private Shared Function FootnoteMapKey(pageNumber As System.Int32, marker As System.String) As System.String
                Return pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) & "|" & marker.Trim().ToLowerInvariant()
            End Function

            Private Shared Function SanitizeFootnoteId(marker As System.String) As System.String
                Dim value As System.String = System.Text.RegularExpressions.Regex.Replace(marker.Trim(), "[^A-Za-z0-9_-]+", System.String.Empty)
                Return value
            End Function

            Private Shared Function BuildSemanticElements(
                model As PdfDocumentModel,
                options As PdfMarkdownOptions,
                footnoteIds As System.Collections.Generic.Dictionary(Of System.String, System.String)
            ) As System.Collections.Generic.List(Of SemanticElement)
                Dim elements As New System.Collections.Generic.List(Of SemanticElement)()

                For pageIndex As System.Int32 = 0 To model.Pages.Count - 1
                    Dim page As PdfPageModel = model.Pages(pageIndex)
                    Dim lines As System.Collections.Generic.List(Of PdfLineModel) = GetOrderedPageLines(page, options)
                    Dim index As System.Int32 = 0

                    While index < lines.Count
                        Dim line As PdfLineModel = lines(index)

                        Dim footnoteMarker As System.String = System.String.Empty
                        Dim footnoteContent As System.String = System.String.Empty
                        If options.PreserveFootnotes AndAlso TryParseFootnoteStart(line, page, model.BodyFontSize, options, footnoteMarker, footnoteContent) Then
                            Dim footnote As SemanticElement = ReadFootnote(lines, index, page, model.BodyFontSize, options, footnoteIds)
                            elements.Add(footnote)
                            Continue While
                        End If

                        If options.DetectTables Then
                            Dim table As TableCandidate = TryReadTable(lines, index, page, options)
                            If table IsNot Nothing AndAlso table.ConsumedLineCount >= 2 Then
                                Dim tableElement As New SemanticElement()
                                tableElement.Kind = PdfSemanticKind.Table
                                tableElement.Table = table
                                tableElement.PageNumber = page.PageNumber
                                For tableLineIndex As System.Int32 = 0 To table.ConsumedLineCount - 1
                                    tableElement.Lines.Add(lines(index + tableLineIndex))
                                Next
                                elements.Add(tableElement)
                                index += table.ConsumedLineCount
                                Continue While
                            End If
                        End If

                        Dim headingLevel As System.Int32 = 0
                        If options.DetectHeadings Then
                            headingLevel = DetectHeadingLevel(line, model.BodyFontSize, options.MaximumHeadingLevel)
                        End If

                        If headingLevel > 0 Then
                            Dim heading As SemanticElement = ReadHeading(lines, index, model.BodyFontSize, options)
                            elements.Add(heading)
                            Continue While
                        End If

                        If options.DetectCaptions AndAlso IsLikelyCaption(line) Then
                            Dim caption As New SemanticElement()
                            caption.Kind = PdfSemanticKind.Caption
                            caption.PageNumber = page.PageNumber
                            caption.Lines.Add(line)
                            caption.Text = line.Text
                            elements.Add(caption)
                            index += 1
                            Continue While
                        End If

                        Dim markerInfo As ListMarkerInfo = Nothing
                        If options.DetectLists AndAlso TryGetListMarker(line.Text, markerInfo) Then
                            Dim listItem As SemanticElement = ReadListItem(lines, index, page, model.BodyFontSize, options)
                            ' Never emit orphan marker glyphs as empty Markdown list items.
                            ' PdfPig can return a bullet/dash as a separate block even when its
                            ' associated text was classified elsewhere. Keeping the text is safer
                            ' than producing a meaningless "-" entry.
                            If listItem IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(listItem.Text) Then
                                elements.Add(listItem)
                            End If
                            Continue While
                        End If

                        Dim paragraph As SemanticElement = ReadParagraph(lines, index, page, model.BodyFontSize, options)
                        If paragraph IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(paragraph.Text) Then
                            elements.Add(paragraph)
                        End If
                    End While
                Next

                Return elements
            End Function

            Private Shared Function ReadHeading(
                lines As System.Collections.Generic.List(Of PdfLineModel),
                ByRef index As System.Int32,
                bodyFontSize As System.Double,
                options As PdfMarkdownOptions
            ) As SemanticElement
                Dim firstLine As PdfLineModel = lines(index)
                Dim level As System.Int32 = DetectHeadingLevel(firstLine, bodyFontSize, options.MaximumHeadingLevel)
                Dim result As New SemanticElement()
                result.Kind = PdfSemanticKind.Heading
                result.PageNumber = firstLine.PageNumber
                result.HeadingLevel = level
                result.Lines.Add(firstLine)
                index += 1

                While index < lines.Count
                    Dim nextLine As PdfLineModel = lines(index)
                    If Not AreCompatibleHeadingLines(firstLine, nextLine, bodyFontSize, options.MaximumHeadingLevel) Then
                        Exit While
                    End If
                    If Not IsReasonableHeadingContinuation(result.Lines(result.Lines.Count - 1), nextLine, bodyFontSize) Then
                        Exit While
                    End If
                    result.Lines.Add(nextLine)
                    index += 1
                End While

                result.Text = JoinLineModelsPlain(result.Lines)
                Return result
            End Function

            Private Shared Function AreCompatibleHeadingLines(
                firstLine As PdfLineModel,
                secondLine As PdfLineModel,
                bodyFontSize As System.Double,
                maximumHeadingLevel As System.Int32
            ) As System.Boolean
                Dim firstLevel As System.Int32 = DetectHeadingLevel(firstLine, bodyFontSize, maximumHeadingLevel)
                Dim secondLevel As System.Int32 = DetectHeadingLevel(secondLine, bodyFontSize, maximumHeadingLevel)

                If firstLevel <= 0 OrElse secondLevel <= 0 OrElse firstLevel <> secondLevel Then
                    Return False
                End If

                Dim allowedFontDifference As System.Double = System.Math.Max(firstLine.FontSize, secondLine.FontSize) * 0.08R
                If System.Math.Abs(firstLine.FontSize - secondLine.FontSize) > allowedFontDifference Then
                    Return False
                End If

                If System.Math.Abs(firstLine.Left - secondLine.Left) > System.Math.Max(12.0R, bodyFontSize * 1.8R) Then
                    Return False
                End If

                Return True
            End Function

            Private Shared Function IsReasonableHeadingContinuation(
                previousLine As PdfLineModel,
                currentLine As PdfLineModel,
                bodyFontSize As System.Double
            ) As System.Boolean
                If previousLine.PageNumber <> currentLine.PageNumber Then
                    Return False
                End If

                Dim gap As System.Double = previousLine.Bottom - currentLine.Top
                Dim allowedGap As System.Double = System.Math.Max(bodyFontSize * 1.6R, previousLine.FontSize * 1.5R)
                Return gap <= allowedGap
            End Function

            Private Shared Function DetectHeadingLevel(
                line As PdfLineModel,
                bodyFontSize As System.Double,
                maximumHeadingLevel As System.Int32
            ) As System.Int32
                If bodyFontSize <= 0.0R OrElse line.FontSize <= 0.0R Then
                    Return 0
                End If

                Dim text As System.String = line.Text.Trim()
                If text.Length = 0 OrElse text.Length > 220 Then
                    Return 0
                End If

                If IsLikelyPageNumberText(text) Then
                    Return 0
                End If

                Dim wordCount As System.Int32 = System.Text.RegularExpressions.Regex.Matches(text, "\S+").Count
                If wordCount > 30 Then
                    Return 0
                End If

                Dim ratio As System.Double = line.FontSize / bodyFontSize
                Dim numberingEvidence As System.Boolean = HasHeadingNumbering(text)
                Dim sentenceLike As System.Boolean = LooksLikeNormalSentence(text)
                Dim level As System.Int32 = 0

                If ratio >= 1.8R Then
                    level = 1
                ElseIf ratio >= 1.55R Then
                    level = 2
                ElseIf ratio >= 1.35R Then
                    level = 3
                ElseIf ratio >= 1.18R Then
                    level = 4
                ElseIf ratio >= 1.08R AndAlso line.BoldRatio >= 0.6R AndAlso wordCount <= 18 Then
                    level = 5
                ElseIf numberingEvidence AndAlso line.BoldRatio >= 0.55R AndAlso wordCount <= 20 AndAlso ratio >= 0.98R Then
                    level = 5
                End If

                If level = 0 Then
                    Return 0
                End If

                If sentenceLike AndAlso ratio < 1.18R AndAlso Not numberingEvidence Then
                    Return 0
                End If

                Return System.Math.Min(System.Math.Max(1, maximumHeadingLevel), level)
            End Function

            Private Shared Function HasHeadingNumbering(text As System.String) As System.Boolean
                Return System.Text.RegularExpressions.Regex.IsMatch(
                    text,
                    "^\s*(?:(?:\d+(?:\.\d+){0,5})|(?:[IVXLCDM]+)|(?:[A-Z]))[\.)]?\s+",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                )
            End Function

            Private Shared Function LooksLikeNormalSentence(text As System.String) As System.Boolean
                If text.Length < 45 Then
                    Return False
                End If
                Dim punctuationCount As System.Int32 = System.Text.RegularExpressions.Regex.Matches(text, "[,;:]").Count
                Return punctuationCount >= 2 AndAlso text.EndsWith(".", System.StringComparison.Ordinal)
            End Function

            Private Shared Function TryGetListMarker(text As System.String, ByRef markerInfo As ListMarkerInfo) As System.Boolean
                markerInfo = Nothing
                If System.String.IsNullOrWhiteSpace(text) Then
                    Return False
                End If

                ' Requiring whitespace or end-of-line after ordered markers prevents decimal
                ' fragments such as "4.0" from being interpreted as list item "4." + "0".
                Dim patterns As System.String() = {
                    "^\s*([-•▪▫◦●○*+‣⁃▸])\s*(.*)$",
                    "^\s*(\(?\d{1,4}[\.)])(?=\s|$)\s*(.*)$",
                    "^\s*(\([A-Za-z]\))(?=\s|$)\s*(.*)$",
                    "^\s*([A-Za-z][\.)])(?=\s|$)\s+(.*)$",
                    "^\s*(\([ivxlcdmIVXLCDM]{1,8}\))(?=\s|$)\s*(.*)$",
                    "^\s*([ivxlcdmIVXLCDM]{1,8}[\.)])(?=\s|$)\s+(.*)$"
                }

                For patternIndex As System.Int32 = 0 To patterns.Length - 1
                    Dim match As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(text, patterns(patternIndex))
                    If Not match.Success Then
                        Continue For
                    End If

                    Dim rawMarker As System.String = match.Groups(1).Value
                    Dim remainder As System.String = match.Groups(2).Value

                    ' A wrapped decimal/version such as "4. 0" is still not a list marker.
                    If patternIndex = 1 AndAlso rawMarker.EndsWith(".", System.StringComparison.Ordinal) AndAlso
                        System.Text.RegularExpressions.Regex.IsMatch(remainder, "^\d") Then
                        Continue For
                    End If

                    Dim result As New ListMarkerInfo()
                    result.RawMarker = rawMarker
                    result.MarkerOnly = remainder.Trim().Length = 0
                    result.PrefixLength = match.Groups(2).Index
                    result.IsOrdered = patternIndex <> 0

                    If patternIndex = 1 Then
                        Dim numberMatch As System.Text.RegularExpressions.Match =
                            System.Text.RegularExpressions.Regex.Match(rawMarker, "\d{1,4}")
                        If numberMatch.Success Then
                            result.MarkdownMarker = numberMatch.Value & ". "
                        Else
                            result.MarkdownMarker = "1. "
                        End If
                    ElseIf result.IsOrdered Then
                        ' Markdown has no portable native alpha/Roman ordered-list syntax.
                        ' Keep these as visible source markers rather than fabricating numbering.
                        result.MarkdownMarker = rawMarker.Trim() & " "
                    Else
                        result.MarkdownMarker = "- "
                    End If

                    markerInfo = result
                    Return True
                Next

                Return False
            End Function

            Private Shared Function ReadListItem(
                lines As System.Collections.Generic.List(Of PdfLineModel),
                ByRef index As System.Int32,
                page As PdfPageModel,
                bodyFontSize As System.Double,
                options As PdfMarkdownOptions
            ) As SemanticElement
                Dim firstLine As PdfLineModel = lines(index)
                Dim markerInfo As ListMarkerInfo = Nothing
                If Not TryGetListMarker(firstLine.Text, markerInfo) Then
                    Return Nothing
                End If

                Dim result As New SemanticElement()
                result.Kind = PdfSemanticKind.ListItem
                result.PageNumber = firstLine.PageNumber
                result.ListMarker = markerInfo.MarkdownMarker
                result.ListIndentLevel = EstimateListIndent(firstLine, page)
                result.Lines.Add(firstLine)
                index += 1

                While index < lines.Count
                    Dim current As PdfLineModel = lines(index)
                    Dim nextMarker As ListMarkerInfo = Nothing
                    If TryGetListMarker(current.Text, nextMarker) Then
                        Exit While
                    End If
                    If DetectHeadingLevel(current, bodyFontSize, options.MaximumHeadingLevel) > 0 Then
                        Exit While
                    End If
                    If options.DetectCaptions AndAlso IsLikelyCaption(current) Then
                        Exit While
                    End If
                    If options.PreserveFootnotes AndAlso IsLikelyFootnoteLine(current, page, bodyFontSize, options) Then
                        Exit While
                    End If
                    If options.DetectTables Then
                        Dim possibleTable As TableCandidate = TryReadTable(lines, index, page, options)
                        If possibleTable IsNot Nothing AndAlso possibleTable.ConsumedLineCount >= 2 Then
                            Exit While
                        End If
                    End If
                    If Not ShouldContinueListItem(result.Lines(result.Lines.Count - 1), current, firstLine.Left, bodyFontSize, page) Then
                        Exit While
                    End If

                    result.Lines.Add(current)
                    index += 1
                End While

                result.Text = JoinListItemPlainText(result.Lines, markerInfo)
                Return result
            End Function

            Private Shared Function EstimateListIndent(line As PdfLineModel, page As PdfPageModel) As System.Int32
                If page.Width <= 0.0R Then
                    Return 0
                End If

                Dim normalized As System.Double = line.Left / page.Width
                If normalized < 0.12R Then
                    Return 0
                ElseIf normalized < 0.2R Then
                    Return 1
                ElseIf normalized < 0.28R Then
                    Return 2
                End If
                Return 3
            End Function

            Private Shared Function ShouldContinueListItem(
                previous As PdfLineModel,
                current As PdfLineModel,
                listLeft As System.Double,
                bodyFontSize As System.Double,
                page As PdfPageModel
            ) As System.Boolean
                If current.PageNumber <> previous.PageNumber Then
                    Return False
                End If

                Dim verticalGap As System.Double = previous.Bottom - current.Top
                Dim maxGap As System.Double = System.Math.Max(bodyFontSize * 1.9R, previous.FontSize * 1.8R)
                If verticalGap > maxGap Then
                    Return False
                End If

                Dim leftTolerance As System.Double = System.Math.Max(bodyFontSize * 2.0R, page.Width * 0.035R)
                Return current.Left >= listLeft - leftTolerance
            End Function

            Private Shared Function JoinListItemPlainText(
                lines As System.Collections.Generic.List(Of PdfLineModel),
                markerInfo As ListMarkerInfo
            ) As System.String
                Dim parts As New System.Collections.Generic.List(Of System.String)()
                For index As System.Int32 = 0 To lines.Count - 1
                    Dim text As System.String = lines(index).Text
                    If index = 0 Then
                        Dim prefixLength As System.Int32 = System.Math.Min(markerInfo.PrefixLength, text.Length)
                        text = text.Substring(prefixLength).TrimStart()
                    End If
                    If text.Length > 0 Then
                        parts.Add(text)
                    End If
                Next
                Return JoinParagraphLines(parts)
            End Function

            Private Shared Function ReadParagraph(
                lines As System.Collections.Generic.List(Of PdfLineModel),
                ByRef index As System.Int32,
                page As PdfPageModel,
                bodyFontSize As System.Double,
                options As PdfMarkdownOptions
            ) As SemanticElement
                If index < 0 OrElse index >= lines.Count Then
                    Return Nothing
                End If

                Dim result As New SemanticElement()
                result.Kind = PdfSemanticKind.Paragraph
                result.PageNumber = lines(index).PageNumber
                result.Lines.Add(lines(index))
                index += 1

                If options.JoinWrappedLines Then
                    While index < lines.Count
                        Dim nextLine As PdfLineModel = lines(index)
                        If options.DetectHeadings AndAlso DetectHeadingLevel(nextLine, bodyFontSize, options.MaximumHeadingLevel) > 0 Then
                            Exit While
                        End If
                        Dim markerInfo As ListMarkerInfo = Nothing
                        If options.DetectLists AndAlso TryGetListMarker(nextLine.Text, markerInfo) Then
                            Exit While
                        End If
                        If options.DetectCaptions AndAlso IsLikelyCaption(nextLine) Then
                            Exit While
                        End If
                        If options.PreserveFootnotes AndAlso IsLikelyFootnoteLine(nextLine, page, bodyFontSize, options) Then
                            Exit While
                        End If
                        If options.DetectTables Then
                            Dim nextTable As TableCandidate = TryReadTable(lines, index, page, options)
                            If nextTable IsNot Nothing AndAlso nextTable.ConsumedLineCount >= 2 Then
                                Exit While
                            End If
                        End If
                        If Not ShouldContinueParagraph(result.Lines(result.Lines.Count - 1), nextLine, page, bodyFontSize) Then
                            Exit While
                        End If

                        result.Lines.Add(nextLine)
                        index += 1
                    End While
                End If

                result.Text = JoinLineModelsPlain(result.Lines)
                Return result
            End Function

            Private Shared Function ShouldContinueParagraph(
                previous As PdfLineModel,
                current As PdfLineModel,
                page As PdfPageModel,
                bodyFontSize As System.Double
            ) As System.Boolean
                If previous.PageNumber <> current.PageNumber Then
                    Return False
                End If

                Dim verticalGap As System.Double = previous.Bottom - current.Top
                Dim normalGap As System.Double = System.Math.Max(bodyFontSize * 1.65R, previous.FontSize * 1.55R)
                If verticalGap > normalGap Then
                    Return False
                End If

                Dim indentDifference As System.Double = System.Math.Abs(current.Left - previous.Left)
                Dim indentTolerance As System.Double = System.Math.Max(bodyFontSize * 2.25R, page.Width * 0.035R)
                If indentDifference > indentTolerance Then
                    Return False
                End If

                If previous.BlockIndex <> current.BlockIndex Then
                    Dim blockGapTolerance As System.Double = System.Math.Max(bodyFontSize * 1.25R, previous.FontSize * 1.2R)
                    If verticalGap > blockGapTolerance Then
                        Return False
                    End If
                End If

                Return True
            End Function

            Private Shared Function IsLikelyCaption(line As PdfLineModel) As System.Boolean
                Dim text As System.String = line.Text.Trim()
                Return System.Text.RegularExpressions.Regex.IsMatch(
                    text,
                    "^(?:(?:Abb(?:ildung)?\.?|Tabelle|Fig(?:ure)?\.?|Table)\s*[A-Za-z0-9IVXLCDM.-]+\s*[:.\-–—]?)\s+.+$",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                )
            End Function

            Private Shared Function IsLikelyFootnoteLine(
                line As PdfLineModel,
                page As PdfPageModel,
                bodyFontSize As System.Double,
                options As PdfMarkdownOptions
            ) As System.Boolean
                If Not options.PreserveFootnotes OrElse line Is Nothing OrElse page Is Nothing Then
                    Return False
                End If
                If bodyFontSize <= 0.0R OrElse line.FontSize <= 0.0R Then
                    Return False
                End If
                If page.Height <= 0.0R Then
                    Return False
                End If

                Dim bandHeight As System.Double = page.Height * System.Math.Max(0.12R, System.Math.Min(0.45R, options.FootnoteBandFraction))
                Dim inBottomBand As System.Boolean = line.Bottom <= bandHeight
                Dim smallerFont As System.Boolean = line.FontSize <= bodyFontSize * 0.9R
                Return inBottomBand AndAlso smallerFont AndAlso Not IsLikelyPageNumberText(line.Text)
            End Function

            Private Shared Function TryParseFootnoteStart(
                line As PdfLineModel,
                page As PdfPageModel,
                bodyFontSize As System.Double,
                options As PdfMarkdownOptions,
                ByRef marker As System.String,
                ByRef content As System.String
            ) As System.Boolean
                marker = System.String.Empty
                content = System.String.Empty
                If Not IsLikelyFootnoteLine(line, page, bodyFontSize, options) Then
                    Return False
                End If

                Dim match As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(
                    line.Text,
                    "^\s*(?:\[)?(\d{1,4}|[A-Za-z]|[*†‡])(?:\]|[\.)])?\s+(.+)$"
                )

                If Not match.Success Then
                    Return False
                End If

                marker = match.Groups(1).Value.Trim()
                content = match.Groups(2).Value.Trim()
                Return marker.Length > 0 AndAlso content.Length > 0
            End Function

            Private Shared Function ReadFootnote(
                lines As System.Collections.Generic.List(Of PdfLineModel),
                ByRef index As System.Int32,
                page As PdfPageModel,
                bodyFontSize As System.Double,
                options As PdfMarkdownOptions,
                footnoteIds As System.Collections.Generic.Dictionary(Of System.String, System.String)
            ) As SemanticElement
                Dim result As New SemanticElement()
                result.Kind = PdfSemanticKind.Footnote
                result.PageNumber = page.PageNumber

                Dim marker As System.String = System.String.Empty
                Dim content As System.String = System.String.Empty
                Dim startLine As PdfLineModel = lines(index)
                If Not TryParseFootnoteStart(startLine, page, bodyFontSize, options, marker, content) Then
                    result.Text = startLine.Text
                    result.Lines.Add(startLine)
                    index += 1
                    Return result
                End If

                result.FootnoteMarker = marker
                Dim mapKey As System.String = FootnoteMapKey(page.PageNumber, marker)
                If footnoteIds.ContainsKey(mapKey) Then
                    result.FootnoteId = footnoteIds(mapKey)
                Else
                    result.FootnoteId = SanitizeFootnoteId(marker)
                End If

                result.Lines.Add(startLine)
                Dim parts As New System.Collections.Generic.List(Of System.String)()
                parts.Add(content)
                index += 1

                While index < lines.Count
                    Dim nextLine As PdfLineModel = lines(index)
                    If Not IsLikelyFootnoteLine(nextLine, page, bodyFontSize, options) Then
                        Exit While
                    End If

                    Dim nextMarker As System.String = System.String.Empty
                    Dim nextContent As System.String = System.String.Empty
                    If TryParseFootnoteStart(nextLine, page, bodyFontSize, options, nextMarker, nextContent) Then
                        Exit While
                    End If

                    result.Lines.Add(nextLine)
                    parts.Add(nextLine.Text)
                    index += 1
                End While

                result.Text = JoinParagraphLines(parts)
                Return result
            End Function

            Private Shared Function TryReadTable(
                lines As System.Collections.Generic.List(Of PdfLineModel),
                startIndex As System.Int32,
                page As PdfPageModel,
                options As PdfMarkdownOptions
            ) As TableCandidate
                If startIndex < 0 OrElse startIndex >= lines.Count Then
                    Return Nothing
                End If

                Dim firstCells As System.Collections.Generic.List(Of System.String) = Nothing
                Dim firstStarts As System.Collections.Generic.List(Of System.Double) = Nothing
                SplitLineIntoTableCells(lines(startIndex), page, firstCells, firstStarts)
                If firstCells Is Nothing OrElse firstCells.Count < 2 Then
                    Return Nothing
                End If

                Dim candidate As New TableCandidate()
                candidate.Rows.Add(firstCells)
                candidate.ColumnStarts.AddRange(firstStarts)
                candidate.ConsumedLineCount = 1
                Dim expectedColumns As System.Int32 = firstCells.Count
                Dim tolerance As System.Double = System.Math.Max(6.0R, page.Width * System.Math.Max(0.005R, options.TableColumnToleranceFraction))

                For index As System.Int32 = startIndex + 1 To lines.Count - 1
                    Dim cells As System.Collections.Generic.List(Of System.String) = Nothing
                    Dim starts As System.Collections.Generic.List(Of System.Double) = Nothing
                    SplitLineIntoTableCells(lines(index), page, cells, starts)
                    If cells Is Nothing OrElse cells.Count <> expectedColumns Then
                        Exit For
                    End If
                    If Not ColumnStartsCompatible(candidate.ColumnStarts, starts, tolerance) Then
                        Exit For
                    End If
                    candidate.Rows.Add(cells)
                    candidate.ConsumedLineCount += 1
                Next

                If candidate.Rows.Count < 2 Then
                    Return Nothing
                End If

                Return candidate
            End Function

            Private Shared Sub SplitLineIntoTableCells(
                line As PdfLineModel,
                page As PdfPageModel,
                ByRef cells As System.Collections.Generic.List(Of System.String),
                ByRef starts As System.Collections.Generic.List(Of System.Double)
            )
                cells = New System.Collections.Generic.List(Of System.String)()
                starts = New System.Collections.Generic.List(Of System.Double)()
                If line.Words.Count < 2 Then
                    Return
                End If

                Dim gapThreshold As System.Double = System.Math.Max(12.0R, page.Width * 0.025R)
                Dim builder As New System.Text.StringBuilder()
                Dim cellStart As System.Double = line.Words(0).Left

                For wordIndex As System.Int32 = 0 To line.Words.Count - 1
                    Dim word As PdfWordModel = line.Words(wordIndex)
                    If builder.Length > 0 Then
                        Dim previousWord As PdfWordModel = line.Words(wordIndex - 1)
                        Dim gap As System.Double = word.Left - previousWord.Right
                        If gap >= gapThreshold Then
                            cells.Add(builder.ToString().Trim())
                            starts.Add(cellStart)
                            builder.Clear()
                            cellStart = word.Left
                        Else
                            builder.Append(" ")
                        End If
                    End If
                    builder.Append(word.Text)
                Next

                If builder.Length > 0 Then
                    cells.Add(builder.ToString().Trim())
                    starts.Add(cellStart)
                End If

                If cells.Count < 2 Then
                    cells.Clear()
                    starts.Clear()
                End If
            End Sub

            Private Shared Function ColumnStartsCompatible(
                expected As System.Collections.Generic.List(Of System.Double),
                actual As System.Collections.Generic.List(Of System.Double),
                tolerance As System.Double
            ) As System.Boolean
                If expected.Count <> actual.Count Then
                    Return False
                End If

                For index As System.Int32 = 0 To expected.Count - 1
                    If System.Math.Abs(expected(index) - actual(index)) > tolerance Then
                        Return False
                    End If
                Next
                Return True
            End Function

            Private Shared Sub MergeCompatibleSemanticElements(
                elements As System.Collections.Generic.List(Of SemanticElement),
                model As PdfDocumentModel,
                options As PdfMarkdownOptions
            )
                If elements.Count < 2 Then
                    Return
                End If

                Dim index As System.Int32 = 0
                While index < elements.Count - 1
                    Dim current As SemanticElement = elements(index)
                    Dim nextElement As SemanticElement = elements(index + 1)

                    If current.Kind = PdfSemanticKind.Heading AndAlso nextElement.Kind = PdfSemanticKind.Heading AndAlso
                        current.HeadingLevel = nextElement.HeadingLevel AndAlso
                        options.JoinAcrossPageBreaks AndAlso nextElement.PageNumber = current.PageNumber + 1 AndAlso
                        IsLikelyCrossPageHeadingContinuation(current, nextElement, model) Then

                        current.Lines.AddRange(nextElement.Lines)
                        current.Text = JoinLineModelsPlain(current.Lines)
                        elements.RemoveAt(index + 1)
                        Continue While
                    End If

                    If current.Kind = PdfSemanticKind.ListItem AndAlso nextElement.Kind = PdfSemanticKind.Paragraph AndAlso
                        options.JoinAcrossPageBreaks AndAlso nextElement.PageNumber = current.PageNumber + 1 AndAlso
                        IsLikelyCrossPageContinuation(current, nextElement, model) Then

                        current.Lines.AddRange(nextElement.Lines)
                        current.Text = JoinLineModelsPlainAfterListPrefix(current.Lines)
                        elements.RemoveAt(index + 1)
                        Continue While
                    End If

                    If current.Kind = PdfSemanticKind.Paragraph AndAlso nextElement.Kind = PdfSemanticKind.Paragraph AndAlso
                        options.JoinAcrossPageBreaks AndAlso nextElement.PageNumber = current.PageNumber + 1 AndAlso
                        IsLikelyCrossPageContinuation(current, nextElement, model) Then

                        current.Lines.AddRange(nextElement.Lines)
                        current.Text = JoinLineModelsPlain(current.Lines)
                        elements.RemoveAt(index + 1)
                        Continue While
                    End If

                    If current.Kind = PdfSemanticKind.Table AndAlso nextElement.Kind = PdfSemanticKind.Table AndAlso
                        nextElement.PageNumber = current.PageNumber + 1 AndAlso TablesHaveSameHeader(current.Table, nextElement.Table) Then

                        For rowIndex As System.Int32 = 1 To nextElement.Table.Rows.Count - 1
                            current.Table.Rows.Add(nextElement.Table.Rows(rowIndex))
                        Next
                        current.Lines.AddRange(nextElement.Lines)
                        elements.RemoveAt(index + 1)
                        Continue While
                    End If

                    index += 1
                End While
            End Sub

            Private Shared Function IsLikelyCrossPageHeadingContinuation(
                first As SemanticElement,
                second As SemanticElement,
                model As PdfDocumentModel
            ) As System.Boolean
                If first.Lines.Count = 0 OrElse second.Lines.Count = 0 Then
                    Return False
                End If
                Dim firstLine As PdfLineModel = first.Lines(first.Lines.Count - 1)
                Dim secondLine As PdfLineModel = second.Lines(0)
                Return AreCompatibleHeadingLines(firstLine, secondLine, model.BodyFontSize, 6) AndAlso
                    first.Text.Length < 240 AndAlso Not EndsWithSentenceTerminal(first.Text)
            End Function

            Private Shared Function IsLikelyCrossPageContinuation(
                first As SemanticElement,
                second As SemanticElement,
                model As PdfDocumentModel
            ) As System.Boolean
                If first.Lines.Count = 0 OrElse second.Lines.Count = 0 Then
                    Return False
                End If

                Dim firstLine As PdfLineModel = first.Lines(first.Lines.Count - 1)
                Dim secondLine As PdfLineModel = second.Lines(0)
                Dim firstPage As PdfPageModel = model.Pages(firstLine.PageNumber - 1)
                Dim secondPage As PdfPageModel = model.Pages(secondLine.PageNumber - 1)

                Dim nearBottom As System.Boolean = firstLine.Bottom <= firstPage.Height * 0.18R
                Dim nearTop As System.Boolean = secondLine.Top >= secondPage.Height * 0.82R
                Dim similarLeft As System.Boolean = System.Math.Abs(firstLine.Left - secondLine.Left) <= System.Math.Max(model.BodyFontSize * 2.0R, firstPage.Width * 0.04R)
                Return nearBottom AndAlso nearTop AndAlso similarLeft AndAlso Not EndsWithStrongParagraphBoundary(first.Text)
            End Function

            Private Shared Function TablesHaveSameHeader(first As TableCandidate, second As TableCandidate) As System.Boolean
                If first Is Nothing OrElse second Is Nothing OrElse first.Rows.Count = 0 OrElse second.Rows.Count = 0 Then
                    Return False
                End If
                If first.Rows(0).Count <> second.Rows(0).Count Then
                    Return False
                End If
                For index As System.Int32 = 0 To first.Rows(0).Count - 1
                    If Not System.String.Equals(
                        NormalizeBasicText(first.Rows(0)(index)).Trim(),
                        NormalizeBasicText(second.Rows(0)(index)).Trim(),
                        System.StringComparison.OrdinalIgnoreCase
                    ) Then
                        Return False
                    End If
                Next
                Return True
            End Function

            Private Shared Function EndsWithSentenceTerminal(text As System.String) As System.Boolean
                Dim trimmed As System.String = text.TrimEnd()
                Return trimmed.EndsWith(".", System.StringComparison.Ordinal) OrElse
                    trimmed.EndsWith("?", System.StringComparison.Ordinal) OrElse
                    trimmed.EndsWith("!", System.StringComparison.Ordinal) OrElse
                    trimmed.EndsWith(":", System.StringComparison.Ordinal)
            End Function

            Private Shared Function EndsWithStrongParagraphBoundary(text As System.String) As System.Boolean
                Dim trimmed As System.String = text.TrimEnd()
                Return trimmed.EndsWith(".", System.StringComparison.Ordinal) OrElse
                    trimmed.EndsWith("?", System.StringComparison.Ordinal) OrElse
                    trimmed.EndsWith("!", System.StringComparison.Ordinal)
            End Function

            Private Shared Function RenderSemanticElements(
                elements As System.Collections.Generic.List(Of SemanticElement),
                model As PdfDocumentModel,
                options As PdfMarkdownOptions,
                footnoteIds As System.Collections.Generic.Dictionary(Of System.String, System.String)
            ) As System.String
                Dim output As New System.Text.StringBuilder(16384)
                Dim footnotes As New System.Collections.Generic.List(Of SemanticElement)()
                Dim lastRenderedPage As System.Int32 = 0

                For Each element As SemanticElement In elements
                    If options.IncludePageBreakComments AndAlso lastRenderedPage > 0 AndAlso element.PageNumber > lastRenderedPage Then
                        AppendBlankLine(output)
                        output.AppendLine("<!-- page " & element.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) & " -->")
                        output.AppendLine()
                    End If
                    If element.PageNumber > 0 Then
                        lastRenderedPage = System.Math.Max(lastRenderedPage, element.PageNumber)
                    End If

                    If element.Kind = PdfSemanticKind.ListItem AndAlso System.String.IsNullOrWhiteSpace(element.Text) Then
                        Continue For
                    End If

                    Select Case element.Kind
                        Case PdfSemanticKind.Heading
                            AppendBlankLine(output)
                            output.AppendLine(
                                New System.String("#"c, System.Math.Max(1, System.Math.Min(options.MaximumHeadingLevel, element.HeadingLevel))) &
                                " " & RenderElementText(element, options, footnoteIds, True)
                            )
                            output.AppendLine()

                        Case PdfSemanticKind.ListItem
                            AppendBlankLine(output)
                            output.Append(New System.String(" "c, System.Math.Max(0, element.ListIndentLevel) * 4))
                            output.Append(If(System.String.IsNullOrWhiteSpace(element.ListMarker), "- ", element.ListMarker))
                            output.AppendLine(RenderListItemText(element, options, footnoteIds))

                        Case PdfSemanticKind.Table
                            AppendBlankLine(output)
                            RenderTable(element.Table, output)
                            output.AppendLine()

                        Case PdfSemanticKind.Caption
                            AppendBlankLine(output)
                            Dim captionText As System.String = RenderElementText(element, options, footnoteIds, False)
                            output.AppendLine("*" & captionText & "*")
                            output.AppendLine()

                        Case PdfSemanticKind.Footnote
                            If options.PreserveFootnotes Then
                                footnotes.Add(element)
                            End If

                        Case PdfSemanticKind.Paragraph
                            AppendBlankLine(output)
                            output.AppendLine(RenderElementText(element, options, footnoteIds, False))
                            output.AppendLine()
                    End Select
                Next

                If options.PreserveFootnotes AndAlso footnotes.Count > 0 Then
                    AppendBlankLine(output)
                    For Each footnote As SemanticElement In footnotes
                        Dim text As System.String = EscapeMarkdownInline(footnote.Text)
                        If options.EmitMarkdownFootnotes AndAlso Not System.String.IsNullOrWhiteSpace(footnote.FootnoteId) Then
                            output.AppendLine("[^" & footnote.FootnoteId & "]: " & text)
                        Else
                            Dim label As System.String = If(System.String.IsNullOrWhiteSpace(footnote.FootnoteMarker), "Footnote", footnote.FootnoteMarker)
                            output.AppendLine("- " & EscapeMarkdownInline(label) & ": " & text)
                        End If
                    Next
                    output.AppendLine()
                End If

                Return output.ToString()
            End Function

            Private Shared Function RenderElementText(
                element As SemanticElement,
                options As PdfMarkdownOptions,
                footnoteIds As System.Collections.Generic.Dictionary(Of System.String, System.String),
                suppressTypography As System.Boolean
            ) As System.String
                If element Is Nothing OrElse element.Lines.Count = 0 Then
                    Return EscapeMarkdownInline(element.Text)
                End If

                Dim renderedLines As New System.Collections.Generic.List(Of System.String)()
                Dim plainLines As New System.Collections.Generic.List(Of System.String)()

                For Each line As PdfLineModel In element.Lines
                    plainLines.Add(line.Text)
                    renderedLines.Add(RenderLineWithTypography(line, options, footnoteIds, suppressTypography))
                Next

                Return JoinRenderedLines(plainLines, renderedLines)
            End Function

            Private Shared Function RenderListItemText(
                element As SemanticElement,
                options As PdfMarkdownOptions,
                footnoteIds As System.Collections.Generic.Dictionary(Of System.String, System.String)
            ) As System.String
                If element Is Nothing OrElse element.Lines.Count = 0 Then
                    Return EscapeMarkdownInline(element.Text)
                End If

                Dim renderedLines As New System.Collections.Generic.List(Of System.String)()
                Dim plainLines As New System.Collections.Generic.List(Of System.String)()
                Dim firstMarker As ListMarkerInfo = Nothing
                TryGetListMarker(element.Lines(0).Text, firstMarker)

                For lineIndex As System.Int32 = 0 To element.Lines.Count - 1
                    Dim line As PdfLineModel = element.Lines(lineIndex)
                    Dim plain As System.String = line.Text
                    Dim rendered As System.String = RenderLineWithTypography(line, options, footnoteIds, False)
                    If lineIndex = 0 AndAlso firstMarker IsNot Nothing Then
                        plain = RemoveListPrefixPlain(plain, firstMarker)
                        rendered = RemoveListPrefixRendered(rendered, line.Text, firstMarker)
                    End If
                    plainLines.Add(plain)
                    renderedLines.Add(rendered)
                Next

                Return JoinRenderedLines(plainLines, renderedLines)
            End Function

            Private Shared Function RemoveListPrefixPlain(text As System.String, markerInfo As ListMarkerInfo) As System.String
                Dim prefixLength As System.Int32 = System.Math.Min(markerInfo.PrefixLength, text.Length)
                Return text.Substring(prefixLength).TrimStart()
            End Function

            Private Shared Function RemoveListPrefixRendered(rendered As System.String, original As System.String, markerInfo As ListMarkerInfo) As System.String
                Dim plainRemainder As System.String = RemoveListPrefixPlain(original, markerInfo)
                If plainRemainder.Length = 0 Then
                    Return System.String.Empty
                End If

                Dim position As System.Int32 = rendered.IndexOf(EscapeMarkdownInline(plainRemainder), System.StringComparison.Ordinal)
                If position >= 0 Then
                    Return rendered.Substring(position)
                End If

                Return EscapeMarkdownInline(plainRemainder)
            End Function

            Private Shared Function RenderLineWithTypography(
                line As PdfLineModel,
                options As PdfMarkdownOptions,
                footnoteIds As System.Collections.Generic.Dictionary(Of System.String, System.String),
                suppressTypography As System.Boolean
            ) As System.String
                If line Is Nothing OrElse line.Words.Count = 0 Then
                    Return EscapeMarkdownInline(line.Text)
                End If

                Dim builder As New System.Text.StringBuilder()
                Dim currentBold As System.Boolean = False
                Dim currentItalic As System.Boolean = False
                Dim previousWord As PdfWordModel = Nothing

                For Each word As PdfWordModel In line.Words
                    Dim gapRequiresSpace As System.Boolean = False
                    If previousWord IsNot Nothing Then
                        Dim gap As System.Double = word.Left - previousWord.Right
                        gapRequiresSpace = gap > System.Math.Max(0.5R, line.FontSize * 0.08R)
                    End If

                    Dim superscriptMarker As System.String = System.String.Empty
                    If word.IsSuperscript AndAlso TryGetFootnoteReferenceId(line.PageNumber, word.Text, footnoteIds, superscriptMarker) Then
                        CloseTypography(builder, currentBold, currentItalic)
                        If gapRequiresSpace Then
                            builder.Append(" ")
                        End If
                        builder.Append("[^")
                        builder.Append(superscriptMarker)
                        builder.Append("]")
                        previousWord = word
                        Continue For
                    End If

                    Dim desiredBold As System.Boolean = options.PreserveBold AndAlso word.IsBold AndAlso Not suppressTypography
                    Dim desiredItalic As System.Boolean = options.PreserveItalic AndAlso word.IsItalic AndAlso Not suppressTypography

                    If previousWord Is Nothing Then
                        AdjustTypography(builder, currentBold, currentItalic, desiredBold, desiredItalic)
                    ElseIf currentBold <> desiredBold OrElse currentItalic <> desiredItalic Then
                        CloseTypography(builder, currentBold, currentItalic)
                        If gapRequiresSpace Then
                            builder.Append(" ")
                        End If
                        AdjustTypography(builder, currentBold, currentItalic, desiredBold, desiredItalic)
                    ElseIf gapRequiresSpace Then
                        builder.Append(" ")
                    End If

                    builder.Append(EscapeMarkdownInline(word.Text))
                    previousWord = word
                Next

                CloseTypography(builder, currentBold, currentItalic)
                Return builder.ToString().Trim()
            End Function

            Private Shared Function TryGetFootnoteReferenceId(
                pageNumber As System.Int32,
                markerText As System.String,
                footnoteIds As System.Collections.Generic.Dictionary(Of System.String, System.String),
                ByRef footnoteId As System.String
            ) As System.Boolean
                footnoteId = System.String.Empty
                Dim marker As System.String = markerText.Trim().Trim("["c, "]"c, "("c, ")"c, "."c)
                If marker.Length = 0 Then
                    Return False
                End If

                Dim key As System.String = FootnoteMapKey(pageNumber, marker)
                If footnoteIds.TryGetValue(key, footnoteId) Then
                    Return True
                End If
                Return False
            End Function

            Private Shared Sub AdjustTypography(
                builder As System.Text.StringBuilder,
                ByRef currentBold As System.Boolean,
                ByRef currentItalic As System.Boolean,
                desiredBold As System.Boolean,
                desiredItalic As System.Boolean
            )
                If currentBold = desiredBold AndAlso currentItalic = desiredItalic Then
                    Return
                End If

                CloseTypography(builder, currentBold, currentItalic)
                If desiredBold AndAlso desiredItalic Then
                    builder.Append("***")
                    currentBold = True
                    currentItalic = True
                ElseIf desiredBold Then
                    builder.Append("**")
                    currentBold = True
                    currentItalic = False
                ElseIf desiredItalic Then
                    builder.Append("*")
                    currentBold = False
                    currentItalic = True
                End If
            End Sub

            Private Shared Sub CloseTypography(
                builder As System.Text.StringBuilder,
                ByRef currentBold As System.Boolean,
                ByRef currentItalic As System.Boolean
            )
                If currentBold AndAlso currentItalic Then
                    builder.Append("***")
                ElseIf currentBold Then
                    builder.Append("**")
                ElseIf currentItalic Then
                    builder.Append("*")
                End If
                currentBold = False
                currentItalic = False
            End Sub

            Private Shared Function JoinRenderedLines(
                plainLines As System.Collections.Generic.List(Of System.String),
                renderedLines As System.Collections.Generic.List(Of System.String)
            ) As System.String
                Dim builder As New System.Text.StringBuilder()
                For index As System.Int32 = 0 To renderedLines.Count - 1
                    Dim rendered As System.String = renderedLines(index).Trim()
                    Dim plain As System.String = NormalizeBasicText(plainLines(index)).Trim()
                    If rendered.Length = 0 Then
                        Continue For
                    End If
                    If builder.Length = 0 Then
                        builder.Append(rendered)
                        Continue For
                    End If

                    Dim previousPlain As System.String = NormalizeBasicText(plainLines(index - 1)).Trim()
                    If ShouldJoinWithoutSpace(previousPlain, plain) Then
                        If previousPlain.EndsWith("-", System.StringComparison.Ordinal) AndAlso builder.Length > 0 Then
                            RemoveTrailingRenderedHyphen(builder)
                        End If
                        builder.Append(rendered)
                    Else
                        builder.Append(" ")
                        builder.Append(rendered)
                    End If
                Next
                Return builder.ToString().Trim()
            End Function

            Private Shared Sub RemoveTrailingRenderedHyphen(builder As System.Text.StringBuilder)
                Dim index As System.Int32 = builder.Length - 1
                While index >= 0 AndAlso (builder(index) = "*"c OrElse builder(index) = "_"c)
                    index -= 1
                End While
                If index >= 0 AndAlso builder(index) = "-"c Then
                    builder.Remove(index, 1)
                End If
            End Sub

            Private Shared Function ShouldJoinWithoutSpace(previous As System.String, current As System.String) As System.Boolean
                If System.String.IsNullOrWhiteSpace(previous) OrElse System.String.IsNullOrWhiteSpace(current) Then
                    Return False
                End If

                Dim firstCharacter As System.Char = current(0)
                If previous.EndsWith("-", System.StringComparison.Ordinal) AndAlso System.Char.IsLetterOrDigit(firstCharacter) Then
                    If System.Char.IsLower(firstCharacter) Then
                        Return True
                    End If
                    If LooksLikeBrokenIdentifier(previous, current) Then
                        Return True
                    End If
                End If

                If LooksLikeBrokenUrlOrEmail(previous, current) Then
                    Return True
                End If

                Return False
            End Function

            Private Shared Function LooksLikeBrokenIdentifier(previous As System.String, current As System.String) As System.Boolean
                Dim combined As System.String = previous & current
                Return System.Text.RegularExpressions.Regex.IsMatch(combined, "(?:[A-Za-z0-9]+[-_/\.])+[A-Za-z0-9]+")
            End Function

            Private Shared Function LooksLikeBrokenUrlOrEmail(previous As System.String, current As System.String) As System.Boolean
                Dim previousTrimmed As System.String = previous.TrimEnd()
                Dim currentTrimmed As System.String = current.TrimStart()
                If System.Text.RegularExpressions.Regex.IsMatch(previousTrimmed, "(?:https?://|www\.)\S*$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                    Return True
                End If
                If previousTrimmed.Contains("@") AndAlso Not previousTrimmed.Contains(" ") AndAlso Not previousTrimmed.EndsWith(".", System.StringComparison.Ordinal) Then
                    Return True
                End If
                If previousTrimmed.EndsWith("/", System.StringComparison.Ordinal) OrElse previousTrimmed.EndsWith("_", System.StringComparison.Ordinal) Then
                    Return True
                End If
                Return False
            End Function

            Private Shared Function JoinLineModelsPlain(lines As System.Collections.Generic.List(Of PdfLineModel)) As System.String
                Dim textLines As New System.Collections.Generic.List(Of System.String)()
                For Each line As PdfLineModel In lines
                    textLines.Add(line.Text)
                Next
                Return JoinParagraphLines(textLines)
            End Function

            Private Shared Function JoinLineModelsPlainAfterListPrefix(lines As System.Collections.Generic.List(Of PdfLineModel)) As System.String
                If lines.Count = 0 Then
                    Return System.String.Empty
                End If

                Dim markerInfo As ListMarkerInfo = Nothing
                TryGetListMarker(lines(0).Text, markerInfo)
                If markerInfo Is Nothing Then
                    Return JoinLineModelsPlain(lines)
                End If

                Dim textLines As New System.Collections.Generic.List(Of System.String)()
                For index As System.Int32 = 0 To lines.Count - 1
                    Dim text As System.String = lines(index).Text
                    If index = 0 Then
                        text = RemoveListPrefixPlain(text, markerInfo)
                    End If
                    textLines.Add(text)
                Next
                Return JoinParagraphLines(textLines)
            End Function

            Private Shared Function JoinParagraphLines(lines As System.Collections.Generic.List(Of System.String)) As System.String
                Dim builder As New System.Text.StringBuilder()

                For index As System.Int32 = 0 To lines.Count - 1
                    Dim current As System.String = NormalizeBasicText(lines(index)).Trim()
                    If current.Length = 0 Then
                        Continue For
                    End If

                    If builder.Length = 0 Then
                        builder.Append(current)
                        Continue For
                    End If

                    Dim previous As System.String = builder.ToString()
                    If ShouldJoinWithoutSpace(previous, current) Then
                        If builder(builder.Length - 1) = "-"c Then
                            builder.Length -= 1
                        End If
                        builder.Append(current)
                    Else
                        builder.Append(" ")
                        builder.Append(current)
                    End If
                Next

                Return builder.ToString().Trim()
            End Function

            Private Shared Sub RenderTable(table As TableCandidate, output As System.Text.StringBuilder)
                If table Is Nothing OrElse table.Rows.Count = 0 Then
                    Return
                End If

                Dim columnCount As System.Int32 = table.Rows(0).Count
                output.Append("|")
                For columnIndex As System.Int32 = 0 To columnCount - 1
                    output.Append(" " & EscapeMarkdownTableCell(table.Rows(0)(columnIndex)) & " |")
                Next
                output.AppendLine()

                output.Append("|")
                For columnIndex As System.Int32 = 0 To columnCount - 1
                    output.Append(" --- |")
                Next
                output.AppendLine()

                For rowIndex As System.Int32 = 1 To table.Rows.Count - 1
                    output.Append("|")
                    For columnIndex As System.Int32 = 0 To columnCount - 1
                        output.Append(" " & EscapeMarkdownTableCell(table.Rows(rowIndex)(columnIndex)) & " |")
                    Next
                    output.AppendLine()
                Next
            End Sub

            Private Shared Function NormalizeMarkdown(value As System.String, options As PdfMarkdownOptions) As System.String
                Dim result As System.String = value.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
                result = NormalizeBasicText(result)

                If options.RemoveSoftHyphens Then
                    result = result.Replace(System.Convert.ToChar(&HAD).ToString(), System.String.Empty)
                End If

                If options.NormalizeLigatures Then
                    result = result.Replace("ﬀ", "ff").Replace("ﬁ", "fi").Replace("ﬂ", "fl").Replace("ﬃ", "ffi").Replace("ﬄ", "ffl").Replace("ﬅ", "st").Replace("ﬆ", "st")
                End If

                result = System.Text.RegularExpressions.Regex.Replace(result, "[\u200B\u200C\u200D\u2060\uFEFF]", System.String.Empty)
                result = System.Text.RegularExpressions.Regex.Replace(result, "[ \t]+\n", vbLf)
                result = System.Text.RegularExpressions.Regex.Replace(result, "\n{3,}", vbLf & vbLf)
                Return result.Replace(vbLf, System.Environment.NewLine).Trim()
            End Function

            Private Shared Function NormalizeBasicText(value As System.String) As System.String
                If System.String.IsNullOrEmpty(value) Then
                    Return System.String.Empty
                End If

                Dim result As System.String = value
                result = result.Replace(System.Convert.ToChar(&HA0), " "c)
                result = result.Replace(System.Convert.ToChar(&H2002), " "c)
                result = result.Replace(System.Convert.ToChar(&H2003), " "c)
                result = result.Replace(System.Convert.ToChar(&H202F), " "c)
                ' Remove invisible format characters BEFORE semantic parsing. Previously these
                ' were removed only in NormalizeMarkdown(), which was too late for list-marker
                ' classification (for example a visually empty bullet followed by U+200B).
                result = System.Text.RegularExpressions.Regex.Replace(result, "[\u200B\u200C\u200D\u2060\uFEFF]", System.String.Empty)
                result = System.Text.RegularExpressions.Regex.Replace(result, "[\t ]+", " ")
                Return result
            End Function

            Private Shared Function Median(values As System.Collections.Generic.List(Of System.Double)) As System.Double
                If values Is Nothing OrElse values.Count = 0 Then
                    Return 0.0R
                End If

                Dim copy As New System.Collections.Generic.List(Of System.Double)(values)
                copy.Sort()
                Dim middle As System.Int32 = copy.Count \ 2
                If copy.Count Mod 2 = 1 Then
                    Return copy(middle)
                End If
                Return (copy(middle - 1) + copy(middle)) / 2.0R
            End Function

            Private Shared Function EscapeMarkdownInline(value As System.String) As System.String
                If System.String.IsNullOrEmpty(value) Then
                    Return System.String.Empty
                End If

                Dim result As System.String = value
                result = result.Replace("\", "\\")
                result = result.Replace("`", "\`")
                result = result.Replace("*", "\*")
                result = EscapeUnderscoresConservatively(result)
                result = result.Replace("[", "\[")
                result = result.Replace("]", "\]")
                result = result.Replace("|", "\|")
                Return result
            End Function

            Private Shared Function EscapeUnderscoresConservatively(value As System.String) As System.String
                If value.IndexOf("_"c) < 0 Then
                    Return value
                End If

                Dim builder As New System.Text.StringBuilder(value.Length + 8)
                For index As System.Int32 = 0 To value.Length - 1
                    Dim character As System.Char = value(index)
                    If character <> "_"c Then
                        builder.Append(character)
                        Continue For
                    End If

                    Dim previousIsAlphaNumeric As System.Boolean = index > 0 AndAlso System.Char.IsLetterOrDigit(value(index - 1))
                    Dim nextIsAlphaNumeric As System.Boolean = index < value.Length - 1 AndAlso System.Char.IsLetterOrDigit(value(index + 1))
                    If previousIsAlphaNumeric AndAlso nextIsAlphaNumeric Then
                        builder.Append("_")
                    Else
                        builder.Append("\_")
                    End If
                Next
                Return builder.ToString()
            End Function

            Private Shared Function EscapeMarkdownTableCell(value As System.String) As System.String
                Dim result As System.String = EscapeMarkdownInline(value)
                result = result.Replace(System.Environment.NewLine, "<br>")
                Return result
            End Function

            Private Shared Sub AppendBlankLine(output As System.Text.StringBuilder)
                If output.Length = 0 Then
                    Return
                End If

                Dim current As System.String = output.ToString()
                If current.EndsWith(System.Environment.NewLine & System.Environment.NewLine, System.StringComparison.Ordinal) Then
                    Return
                End If
                If current.EndsWith(System.Environment.NewLine, System.StringComparison.Ordinal) Then
                    output.AppendLine()
                Else
                    output.AppendLine()
                    output.AppendLine()
                End If
            End Sub

        End Class

    End Class
End Namespace
