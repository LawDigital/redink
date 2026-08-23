' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Tools.Office.OpenXmlVisuals.vb
' Purpose:
'   Deterministic native OOXML visual renderer for create_word_document output.
'
' Architecture / Function:
'   - Runs after the base DOCX body has been produced (by the OOXML template/generic
'     renderer or another supported path) and replaces [[visual:ID]] anchors in-package.
'   - Quantitative charts are native Office chart parts with embedded workbook/cache data;
'     diagrams use editable DrawingML shapes; supported images are embedded as package parts.
'   - Supports inline/floating placement without starting Word or Excel and preserves the
'     surrounding template/body styles; visuals intentionally do not inherit table/paragraph
'     auto-indent rules unless their own placement contract requests it.
'   - Assigns stable RedInk visual identifiers and validates requested visual objects in
'     the saved package so missing/partial rendering is not silently accepted.
' =============================================================================

Option Explicit On
Option Strict Off

Partial Public Class ThisAddIn

    Private Const OoxmlNsWord As System.String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    Private Const OoxmlNsWordDrawing As System.String = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    Private Const OoxmlNsDrawing As System.String = "http://schemas.openxmlformats.org/drawingml/2006/main"
    Private Const OoxmlNsRelationship As System.String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    Private Const OoxmlNsPackageRelationship As System.String = "http://schemas.openxmlformats.org/package/2006/relationships"
    Private Const OoxmlNsContentTypes As System.String = "http://schemas.openxmlformats.org/package/2006/content-types"
    Private Const OoxmlNsChart As System.String = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    Private Const OoxmlNsWordCanvas As System.String = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    Private Const OoxmlNsWordShape As System.String = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    Private Const OoxmlNsMarkupCompatibility As System.String = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    Private Const OoxmlNsSpreadsheet As System.String = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    Private Class AutoPilotOoxmlRect
        Public X As System.Double
        Public Y As System.Double
        Public Width As System.Double
        Public Height As System.Double
    End Class

    Private Class AutoPilotOoxmlNodeLayout
        Public Node As Newtonsoft.Json.Linq.JObject
        Public Rect As AutoPilotOoxmlRect
        Public Depth As System.Int32
        Public ParentId As System.String
    End Class

    Private Shared Function OoxmlPointsToEmu(points As System.Double) As System.Int64
        Return CLng(System.Math.Round(points * 12700.0R, System.MidpointRounding.AwayFromZero))
    End Function

    Private Shared Function OoxmlNormalizeHex(value As System.String, fallback As System.String) As System.String
        Dim raw As System.String = If(value, System.String.Empty).Trim().TrimStart("#"c)
        If System.Text.RegularExpressions.Regex.IsMatch(raw, "^[0-9A-Fa-f]{6}$") Then Return raw.ToUpperInvariant()
        Return If(fallback, "17365D").Trim().TrimStart("#"c).ToUpperInvariant()
    End Function

    Private Shared Function OoxmlLightenHex(hexValue As System.String, amount As System.Double) As System.String
        Dim normalized As System.String = OoxmlNormalizeHex(hexValue, "17365D")
        Dim r As System.Int32 = System.Convert.ToInt32(normalized.Substring(0, 2), 16)
        Dim g As System.Int32 = System.Convert.ToInt32(normalized.Substring(2, 2), 16)
        Dim b As System.Int32 = System.Convert.ToInt32(normalized.Substring(4, 2), 16)
        amount = System.Math.Max(0.0R, System.Math.Min(1.0R, amount))
        r = CInt(System.Math.Round(r + (255 - r) * amount))
        g = CInt(System.Math.Round(g + (255 - g) * amount))
        b = CInt(System.Math.Round(b + (255 - b) * amount))
        Return r.ToString("X2") & g.ToString("X2") & b.ToString("X2")
    End Function

    Private Shared Function OoxmlEscapeSheetName(name As System.String) As System.String
        Dim safe As System.String = If(name, "Data").Replace("'", "''")
        Return "'" & safe & "'"
    End Function

    Private Shared Function OoxmlExcelColumnName(columnIndexOneBased As System.Int32) As System.String
        Dim value As System.Int32 = System.Math.Max(1, columnIndexOneBased)
        Dim result As System.String = System.String.Empty
        While value > 0
            value -= 1
            result = ChrW(AscW("A"c) + (value Mod 26)) & result
            value \= 26
        End While
        Return result
    End Function

    Private Shared Function OoxmlGetVisualPlacement(visual As Newtonsoft.Json.Linq.JObject) As System.String
        Dim placement As System.String = GetVisualText(visual, "insertion_mode", "auto").ToLowerInvariant()
        If placement <> "inline" AndAlso placement <> "floating" Then placement = "auto"
        Return placement
    End Function

    Private Shared Function OoxmlIsChartType(visualType As System.String) As System.Boolean
        Select Case If(visualType, System.String.Empty).ToLowerInvariant()
            Case "bar_chart", "column_chart", "line_chart", "area_chart", "pie_chart", "doughnut_chart"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Shared Function OoxmlIsDiagramType(visualType As System.String) As System.Boolean
        Select Case If(visualType, System.String.Empty).ToLowerInvariant()
            Case "org_chart", "hierarchy", "process", "timeline", "list", "cycle", "relationship", "matrix", "pyramid", "smartart", "diagram"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Shared Sub OoxmlReplaceZipEntry(archive As System.IO.Compression.ZipArchive,
                                            entryName As System.String,
                                            document As System.Xml.Linq.XDocument)
        Dim existing As System.IO.Compression.ZipArchiveEntry = archive.GetEntry(entryName)
        If existing IsNot Nothing Then existing.Delete()
        Dim created As System.IO.Compression.ZipArchiveEntry = archive.CreateEntry(entryName, System.IO.Compression.CompressionLevel.Optimal)
        Using stream As System.IO.Stream = created.Open()
            Dim settings As New System.Xml.XmlWriterSettings() With {
                .Encoding = New System.Text.UTF8Encoding(False),
                .Indent = False,
                .OmitXmlDeclaration = False
            }
            Using writer As System.Xml.XmlWriter = System.Xml.XmlWriter.Create(stream, settings)
                document.Save(writer)
            End Using
        End Using
    End Sub

    Private Shared Sub OoxmlReplaceZipEntryBytes(archive As System.IO.Compression.ZipArchive,
                                                 entryName As System.String,
                                                 data() As System.Byte)
        Dim existing As System.IO.Compression.ZipArchiveEntry = archive.GetEntry(entryName)
        If existing IsNot Nothing Then existing.Delete()
        Dim created As System.IO.Compression.ZipArchiveEntry = archive.CreateEntry(entryName, System.IO.Compression.CompressionLevel.Optimal)
        Using stream As System.IO.Stream = created.Open()
            stream.Write(data, 0, data.Length)
        End Using
    End Sub

    Private Shared Function OoxmlLoadZipXml(archive As System.IO.Compression.ZipArchive,
                                            entryName As System.String) As System.Xml.Linq.XDocument
        Dim entry As System.IO.Compression.ZipArchiveEntry = archive.GetEntry(entryName)
        If entry Is Nothing Then Return Nothing
        Using stream As System.IO.Stream = entry.Open()
            Return System.Xml.Linq.XDocument.Load(stream, System.Xml.Linq.LoadOptions.PreserveWhitespace)
        End Using
    End Function

    Private Shared Function OoxmlEnsureDocumentRelationships(archive As System.IO.Compression.ZipArchive) As System.Xml.Linq.XDocument
        Dim rels As System.Xml.Linq.XDocument = OoxmlLoadZipXml(archive, "word/_rels/document.xml.rels")
        If rels Is Nothing Then
            Dim nsRel As System.Xml.Linq.XNamespace = OoxmlNsPackageRelationship
            rels = New System.Xml.Linq.XDocument(New System.Xml.Linq.XElement(nsRel + "Relationships"))
        End If
        Return rels
    End Function

    Private Shared Function OoxmlNextRelationshipId(rels As System.Xml.Linq.XDocument) As System.String
        Dim nsRel As System.Xml.Linq.XNamespace = OoxmlNsPackageRelationship
        Dim maxId As System.Int32 = 0
        For Each rel As System.Xml.Linq.XElement In rels.Root.Elements(nsRel + "Relationship")
            Dim id As System.String = CStr(rel.Attribute("Id"))
            Dim match As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(If(id, System.String.Empty), "^rId(\d+)$")
            If match.Success Then
                Dim parsed As System.Int32 = 0
                If System.Int32.TryParse(match.Groups(1).Value, parsed) Then maxId = System.Math.Max(maxId, parsed)
            End If
        Next
        Return "rId" & (maxId + 1).ToString(System.Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Shared Function OoxmlNextDrawingId(documentXml As System.Xml.Linq.XDocument) As System.UInt32
        Dim nsWp As System.Xml.Linq.XNamespace = OoxmlNsWordDrawing
        Dim maxId As System.UInt32 = 0UI
        For Each docPr As System.Xml.Linq.XElement In documentXml.Descendants(nsWp + "docPr")
            Dim parsed As System.UInt32 = 0UI
            If System.UInt32.TryParse(CStr(docPr.Attribute("id")), parsed) Then maxId = System.Math.Max(maxId, parsed)
        Next
        Return maxId + 1UI
    End Function

    Private Shared Sub OoxmlEnsureNamespaceDeclarations(documentXml As System.Xml.Linq.XDocument)
        If documentXml Is Nothing OrElse documentXml.Root Is Nothing Then Exit Sub
        documentXml.Root.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + "wpc", OoxmlNsWordCanvas)
        documentXml.Root.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + "wps", OoxmlNsWordShape)
        documentXml.Root.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + "a", OoxmlNsDrawing)
        documentXml.Root.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + "wp", OoxmlNsWordDrawing)
        documentXml.Root.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + "r", OoxmlNsRelationship)
        documentXml.Root.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + "c", OoxmlNsChart)
        documentXml.Root.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + "mc", OoxmlNsMarkupCompatibility)

        Dim nsMc As System.Xml.Linq.XNamespace = OoxmlNsMarkupCompatibility
        Dim current As System.String = CStr(documentXml.Root.Attribute(nsMc + "Ignorable"))
        Dim tokens As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
        For Each token As System.String In If(current, System.String.Empty).Split(New Char() {" "c}, System.StringSplitOptions.RemoveEmptyEntries)
            tokens.Add(token)
        Next
        tokens.Add("wpc")
        tokens.Add("wps")
        documentXml.Root.SetAttributeValue(nsMc + "Ignorable", System.String.Join(" ", tokens))
    End Sub

    Private Shared Function OoxmlFindPlaceholderParagraph(documentXml As System.Xml.Linq.XDocument,
                                                           visualId As System.String,
                                                           ByRef errorText As System.String) As System.Xml.Linq.XElement
        errorText = System.String.Empty
        Dim nsW As System.Xml.Linq.XNamespace = OoxmlNsWord
        Dim placeholder As System.String = "[[visual:" & visualId & "]]"
        Dim matches As New System.Collections.Generic.List(Of System.Xml.Linq.XElement)()
        For Each paragraph As System.Xml.Linq.XElement In documentXml.Descendants(nsW + "p")
            Dim textValue As System.String = System.String.Concat(paragraph.Descendants(nsW + "t").Select(Function(t) t.Value))
            If textValue.IndexOf(placeholder, System.StringComparison.Ordinal) >= 0 Then matches.Add(paragraph)
        Next
        If matches.Count <> 1 Then
            errorText = "Expected exactly one persisted Word placeholder " & placeholder & "; found " & matches.Count.ToString() & "."
            Return Nothing
        End If
        Dim fullText As System.String = System.String.Concat(matches(0).Descendants(nsW + "t").Select(Function(t) t.Value)).Trim()
        If Not System.String.Equals(fullText, placeholder, System.StringComparison.Ordinal) Then
            errorText = placeholder & " must persist as its own Word paragraph for deterministic OOXML replacement."
            Return Nothing
        End If
        Return matches(0)
    End Function

    Private Shared Sub OoxmlReplacePlaceholderParagraph(paragraph As System.Xml.Linq.XElement,
                                                         drawing As System.Xml.Linq.XElement)
        Dim nsW As System.Xml.Linq.XNamespace = OoxmlNsWord
        Dim pPr As System.Xml.Linq.XElement = paragraph.Element(nsW + "pPr")
        Dim preservePPr As System.Xml.Linq.XElement = If(pPr Is Nothing, Nothing, New System.Xml.Linq.XElement(pPr))
        paragraph.RemoveNodes()
        If preservePPr IsNot Nothing Then paragraph.Add(preservePPr)
        paragraph.Add(New System.Xml.Linq.XElement(nsW + "r",
                                                  New System.Xml.Linq.XElement(nsW + "rPr",
                                                                              New System.Xml.Linq.XElement(nsW + "noProof")),
                                                  drawing))
    End Sub

    Private Shared Function OoxmlCreateOuterDrawing(graphicData As System.Xml.Linq.XElement,
                                                     visualId As System.String,
                                                     widthPoints As System.Double,
                                                     heightPoints As System.Double,
                                                     drawingId As System.UInt32,
                                                     placement As System.String) As System.Xml.Linq.XElement
        Dim nsW As System.Xml.Linq.XNamespace = OoxmlNsWord
        Dim nsWp As System.Xml.Linq.XNamespace = OoxmlNsWordDrawing
        Dim nsA As System.Xml.Linq.XNamespace = OoxmlNsDrawing
        Dim cx As System.Int64 = OoxmlPointsToEmu(widthPoints)
        Dim cy As System.Int64 = OoxmlPointsToEmu(heightPoints)
        Dim docPr As New System.Xml.Linq.XElement(nsWp + "docPr",
                                                  New System.Xml.Linq.XAttribute("id", drawingId),
                                                  New System.Xml.Linq.XAttribute("name", "RedInk Visual " & visualId),
                                                  New System.Xml.Linq.XAttribute("descr", "Editable Red Ink Word visual " & visualId))
        Dim framePr As New System.Xml.Linq.XElement(nsWp + "cNvGraphicFramePr",
                                                    New System.Xml.Linq.XElement(nsA + "graphicFrameLocks",
                                                                                New System.Xml.Linq.XAttribute("noChangeAspect", "0")))
        Dim graphic As New System.Xml.Linq.XElement(nsA + "graphic", graphicData)
        Dim container As System.Xml.Linq.XElement
        If placement = "floating" Then
            container = New System.Xml.Linq.XElement(nsWp + "anchor",
                New System.Xml.Linq.XAttribute("distT", "0"),
                New System.Xml.Linq.XAttribute("distB", "0"),
                New System.Xml.Linq.XAttribute("distL", "0"),
                New System.Xml.Linq.XAttribute("distR", "0"),
                New System.Xml.Linq.XAttribute("simplePos", "0"),
                New System.Xml.Linq.XAttribute("relativeHeight", (251658240UI + drawingId).ToString()),
                New System.Xml.Linq.XAttribute("behindDoc", "0"),
                New System.Xml.Linq.XAttribute("locked", "0"),
                New System.Xml.Linq.XAttribute("layoutInCell", "1"),
                New System.Xml.Linq.XAttribute("allowOverlap", "0"),
                New System.Xml.Linq.XElement(nsWp + "simplePos", New System.Xml.Linq.XAttribute("x", "0"), New System.Xml.Linq.XAttribute("y", "0")),
                New System.Xml.Linq.XElement(nsWp + "positionH", New System.Xml.Linq.XAttribute("relativeFrom", "column"), New System.Xml.Linq.XElement(nsWp + "align", "center")),
                New System.Xml.Linq.XElement(nsWp + "positionV", New System.Xml.Linq.XAttribute("relativeFrom", "paragraph"), New System.Xml.Linq.XElement(nsWp + "posOffset", "0")),
                New System.Xml.Linq.XElement(nsWp + "extent", New System.Xml.Linq.XAttribute("cx", cx), New System.Xml.Linq.XAttribute("cy", cy)),
                New System.Xml.Linq.XElement(nsWp + "effectExtent", New System.Xml.Linq.XAttribute("l", "0"), New System.Xml.Linq.XAttribute("t", "0"), New System.Xml.Linq.XAttribute("r", "0"), New System.Xml.Linq.XAttribute("b", "0")),
                New System.Xml.Linq.XElement(nsWp + "wrapTopAndBottom"),
                docPr,
                framePr,
                graphic)
        Else
            container = New System.Xml.Linq.XElement(nsWp + "inline",
                New System.Xml.Linq.XAttribute("distT", "0"),
                New System.Xml.Linq.XAttribute("distB", "0"),
                New System.Xml.Linq.XAttribute("distL", "0"),
                New System.Xml.Linq.XAttribute("distR", "0"),
                New System.Xml.Linq.XElement(nsWp + "extent", New System.Xml.Linq.XAttribute("cx", cx), New System.Xml.Linq.XAttribute("cy", cy)),
                New System.Xml.Linq.XElement(nsWp + "effectExtent", New System.Xml.Linq.XAttribute("l", "0"), New System.Xml.Linq.XAttribute("t", "0"), New System.Xml.Linq.XAttribute("r", "0"), New System.Xml.Linq.XAttribute("b", "0")),
                docPr,
                framePr,
                graphic)
        End If
        Return New System.Xml.Linq.XElement(nsW + "drawing", container)
    End Function

    Private Shared Function OoxmlCreateCanvasBox(shapeId As System.UInt32,
                                                  shapeName As System.String,
                                                  rect As AutoPilotOoxmlRect,
                                                  fillHex As System.String,
                                                  lineHex As System.String,
                                                  label As System.String,
                                                  detail As System.String,
                                                  fontName As System.String,
                                                  textHex As System.String,
                                                  labelBold As System.Boolean) As System.Xml.Linq.XElement
        Dim nsW As System.Xml.Linq.XNamespace = OoxmlNsWord
        Dim nsA As System.Xml.Linq.XNamespace = OoxmlNsDrawing
        Dim nsWps As System.Xml.Linq.XNamespace = OoxmlNsWordShape
        Dim cx As System.Int64 = OoxmlPointsToEmu(rect.Width)
        Dim cy As System.Int64 = OoxmlPointsToEmu(rect.Height)
        Dim ox As System.Int64 = OoxmlPointsToEmu(rect.X)
        Dim oy As System.Int64 = OoxmlPointsToEmu(rect.Y)
        Dim shape As New System.Xml.Linq.XElement(nsWps + "wsp",
            New System.Xml.Linq.XElement(nsWps + "cNvPr", New System.Xml.Linq.XAttribute("id", shapeId), New System.Xml.Linq.XAttribute("name", shapeName)),
            New System.Xml.Linq.XElement(nsWps + "cNvSpPr", New System.Xml.Linq.XAttribute("txBox", "1"), New System.Xml.Linq.XElement(nsA + "spLocks", New System.Xml.Linq.XAttribute("noChangeArrowheads", "1"))),
            New System.Xml.Linq.XElement(nsWps + "spPr",
                New System.Xml.Linq.XElement(nsA + "xfrm",
                    New System.Xml.Linq.XElement(nsA + "off", New System.Xml.Linq.XAttribute("x", ox), New System.Xml.Linq.XAttribute("y", oy)),
                    New System.Xml.Linq.XElement(nsA + "ext", New System.Xml.Linq.XAttribute("cx", cx), New System.Xml.Linq.XAttribute("cy", cy))),
                New System.Xml.Linq.XElement(nsA + "prstGeom", New System.Xml.Linq.XAttribute("prst", "roundRect"), New System.Xml.Linq.XElement(nsA + "avLst")),
                New System.Xml.Linq.XElement(nsA + "solidFill", New System.Xml.Linq.XElement(nsA + "srgbClr", New System.Xml.Linq.XAttribute("val", fillHex))),
                New System.Xml.Linq.XElement(nsA + "ln", New System.Xml.Linq.XAttribute("w", "12700"), New System.Xml.Linq.XElement(nsA + "solidFill", New System.Xml.Linq.XElement(nsA + "srgbClr", New System.Xml.Linq.XAttribute("val", lineHex))))),
            New System.Xml.Linq.XElement(nsWps + "txbx",
                New System.Xml.Linq.XElement(nsW + "txbxContent",
                    OoxmlCreateTextBoxParagraph(label, fontName, textHex, 9.0R, labelBold),
                    If(System.String.IsNullOrWhiteSpace(detail), Nothing, OoxmlCreateTextBoxParagraph(detail, fontName, textHex, 7.5R, False)))),
            New System.Xml.Linq.XElement(nsWps + "bodyPr",
                New System.Xml.Linq.XAttribute("rot", "0"),
                New System.Xml.Linq.XAttribute("vert", "horz"),
                New System.Xml.Linq.XAttribute("wrap", "square"),
                New System.Xml.Linq.XAttribute("lIns", OoxmlPointsToEmu(5.0R)),
                New System.Xml.Linq.XAttribute("tIns", OoxmlPointsToEmu(3.0R)),
                New System.Xml.Linq.XAttribute("rIns", OoxmlPointsToEmu(5.0R)),
                New System.Xml.Linq.XAttribute("bIns", OoxmlPointsToEmu(3.0R)),
                New System.Xml.Linq.XAttribute("anchor", "ctr"),
                New System.Xml.Linq.XAttribute("anchorCtr", "1"),
                New System.Xml.Linq.XElement(nsA + "noAutofit")))
        Return shape
    End Function

    Private Shared Function OoxmlCreateTextBoxParagraph(textValue As System.String,
                                                        fontName As System.String,
                                                        colorHex As System.String,
                                                        fontSizePoints As System.Double,
                                                        bold As System.Boolean) As System.Xml.Linq.XElement
        Dim nsW As System.Xml.Linq.XNamespace = OoxmlNsWord
        Dim rPr As New System.Xml.Linq.XElement(nsW + "rPr",
            New System.Xml.Linq.XElement(nsW + "rFonts", New System.Xml.Linq.XAttribute(nsW + "ascii", fontName), New System.Xml.Linq.XAttribute(nsW + "hAnsi", fontName), New System.Xml.Linq.XAttribute(nsW + "cs", fontName)),
            If(bold, New System.Xml.Linq.XElement(nsW + "b"), Nothing),
            New System.Xml.Linq.XElement(nsW + "color", New System.Xml.Linq.XAttribute(nsW + "val", colorHex)),
            New System.Xml.Linq.XElement(nsW + "sz", New System.Xml.Linq.XAttribute(nsW + "val", CInt(System.Math.Round(fontSizePoints * 2.0R)).ToString())),
            New System.Xml.Linq.XElement(nsW + "szCs", New System.Xml.Linq.XAttribute(nsW + "val", CInt(System.Math.Round(fontSizePoints * 2.0R)).ToString())))
        Return New System.Xml.Linq.XElement(nsW + "p",
            New System.Xml.Linq.XElement(nsW + "pPr",
                New System.Xml.Linq.XElement(nsW + "spacing", New System.Xml.Linq.XAttribute(nsW + "before", "0"), New System.Xml.Linq.XAttribute(nsW + "after", "0"), New System.Xml.Linq.XAttribute(nsW + "line", "200"), New System.Xml.Linq.XAttribute(nsW + "lineRule", "atLeast")),
                New System.Xml.Linq.XElement(nsW + "jc", New System.Xml.Linq.XAttribute(nsW + "val", "center"))),
            New System.Xml.Linq.XElement(nsW + "r", rPr, New System.Xml.Linq.XElement(nsW + "t", If(textValue, System.String.Empty))))
    End Function

    Private Shared Function OoxmlCreateCanvasLine(shapeId As System.UInt32,
                                                   shapeName As System.String,
                                                   x As System.Double,
                                                   y As System.Double,
                                                   width As System.Double,
                                                   height As System.Double,
                                                   lineHex As System.String) As System.Xml.Linq.XElement
        Dim nsA As System.Xml.Linq.XNamespace = OoxmlNsDrawing
        Dim nsWps As System.Xml.Linq.XNamespace = OoxmlNsWordShape
        width = System.Math.Max(0.8R, width)
        height = System.Math.Max(0.8R, height)
        Return New System.Xml.Linq.XElement(nsWps + "wsp",
            New System.Xml.Linq.XElement(nsWps + "cNvPr", New System.Xml.Linq.XAttribute("id", shapeId), New System.Xml.Linq.XAttribute("name", shapeName)),
            New System.Xml.Linq.XElement(nsWps + "cNvSpPr", New System.Xml.Linq.XElement(nsA + "spLocks", New System.Xml.Linq.XAttribute("noChangeArrowheads", "1"))),
            New System.Xml.Linq.XElement(nsWps + "spPr",
                New System.Xml.Linq.XElement(nsA + "xfrm",
                    New System.Xml.Linq.XElement(nsA + "off", New System.Xml.Linq.XAttribute("x", OoxmlPointsToEmu(x)), New System.Xml.Linq.XAttribute("y", OoxmlPointsToEmu(y))),
                    New System.Xml.Linq.XElement(nsA + "ext", New System.Xml.Linq.XAttribute("cx", OoxmlPointsToEmu(width)), New System.Xml.Linq.XAttribute("cy", OoxmlPointsToEmu(height)))),
                New System.Xml.Linq.XElement(nsA + "prstGeom", New System.Xml.Linq.XAttribute("prst", "rect"), New System.Xml.Linq.XElement(nsA + "avLst")),
                New System.Xml.Linq.XElement(nsA + "solidFill", New System.Xml.Linq.XElement(nsA + "srgbClr", New System.Xml.Linq.XAttribute("val", lineHex))),
                New System.Xml.Linq.XElement(nsA + "ln", New System.Xml.Linq.XElement(nsA + "noFill"))),
            New System.Xml.Linq.XElement(nsWps + "bodyPr", New System.Xml.Linq.XElement(nsA + "noAutofit")))
    End Function

    Private Shared Function OoxmlCreateCanvasTriangle(shapeId As System.UInt32,
                                                       shapeName As System.String,
                                                       x As System.Double,
                                                       y As System.Double,
                                                       width As System.Double,
                                                       height As System.Double,
                                                       fillHex As System.String) As System.Xml.Linq.XElement
        Dim nsA As System.Xml.Linq.XNamespace = OoxmlNsDrawing
        Dim nsWps As System.Xml.Linq.XNamespace = OoxmlNsWordShape
        Return New System.Xml.Linq.XElement(nsWps + "wsp",
            New System.Xml.Linq.XElement(nsWps + "cNvPr", New System.Xml.Linq.XAttribute("id", shapeId), New System.Xml.Linq.XAttribute("name", shapeName)),
            New System.Xml.Linq.XElement(nsWps + "cNvSpPr", New System.Xml.Linq.XElement(nsA + "spLocks", New System.Xml.Linq.XAttribute("noChangeArrowheads", "1"))),
            New System.Xml.Linq.XElement(nsWps + "spPr",
                New System.Xml.Linq.XElement(nsA + "xfrm",
                    New System.Xml.Linq.XElement(nsA + "off", New System.Xml.Linq.XAttribute("x", OoxmlPointsToEmu(x)), New System.Xml.Linq.XAttribute("y", OoxmlPointsToEmu(y))),
                    New System.Xml.Linq.XElement(nsA + "ext", New System.Xml.Linq.XAttribute("cx", OoxmlPointsToEmu(width)), New System.Xml.Linq.XAttribute("cy", OoxmlPointsToEmu(height)))),
                New System.Xml.Linq.XElement(nsA + "prstGeom", New System.Xml.Linq.XAttribute("prst", "rtTriangle"), New System.Xml.Linq.XElement(nsA + "avLst")),
                New System.Xml.Linq.XElement(nsA + "solidFill", New System.Xml.Linq.XElement(nsA + "srgbClr", New System.Xml.Linq.XAttribute("val", fillHex))),
                New System.Xml.Linq.XElement(nsA + "ln", New System.Xml.Linq.XElement(nsA + "noFill"))),
            New System.Xml.Linq.XElement(nsWps + "bodyPr", New System.Xml.Linq.XElement(nsA + "noAutofit")))
    End Function

    Private Shared Function OoxmlBuildOrgChartCanvas(visual As Newtonsoft.Json.Linq.JObject,
                                                     fontName As System.String,
                                                     accentHex As System.String,
                                                     ByRef widthPoints As System.Double,
                                                     ByRef heightPoints As System.Double) As System.Xml.Linq.XElement
        Dim nsA As System.Xml.Linq.XNamespace = OoxmlNsDrawing
        Dim nsWpc As System.Xml.Linq.XNamespace = OoxmlNsWordCanvas
        Dim nodes As System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JObject) = GetWordOrgChartNodes(visual)
        If nodes.Count = 0 Then Return OoxmlBuildGenericDiagramCanvas(visual, "process", fontName, accentHex, widthPoints, heightPoints)

        widthPoints = System.Math.Min(468.0R, GetVisualNumber(visual, "width_inches", 6.35R, 3.5R, 7.0R) * 72.0R)
        Dim byId As New System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JObject)(System.StringComparer.Ordinal)
        Dim levels As New System.Collections.Generic.SortedDictionary(Of System.Int32, System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JObject))()
        Dim maxDepth As System.Int32 = 0
        For Each node As Newtonsoft.Json.Linq.JObject In nodes
            byId(GetVisualText(node, "id")) = node
        Next
        For Each node As Newtonsoft.Json.Linq.JObject In nodes
            Dim depth As System.Int32 = GetWordOrgChartDepth(node, byId)
            maxDepth = System.Math.Max(maxDepth, depth)
            If Not levels.ContainsKey(depth) Then levels(depth) = New System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JObject)()
            levels(depth).Add(node)
        Next

        Dim layouts As New System.Collections.Generic.Dictionary(Of System.String, AutoPilotOoxmlNodeLayout)(System.StringComparer.Ordinal)
        Dim marginX As System.Double = 8.0R
        Dim boxH As System.Double = 43.0R
        Dim smallBoxH As System.Double = 38.0R
        Dim verticalGap As System.Double = 15.0R
        Dim shapeId As System.UInt32 = 1UI

        ' Corporate hierarchies commonly have one board/root, one CEO, a fan-out of
        ' functional leaders, then several children per function. A column layout keeps
        ' this compact and avoids both box overlap and connector paths through boxes.
        Dim branchDepth As System.Int32 = If(levels.ContainsKey(2) AndAlso levels(2).Count >= 2, 2, If(levels.ContainsKey(1) AndAlso levels(1).Count >= 2, 1, -1))
        Dim branchCount As System.Int32 = If(branchDepth >= 0, levels(branchDepth).Count, 0)
        Dim useBranchColumns As System.Boolean = branchCount >= 2 AndAlso branchCount <= 6
        Dim canvas As New System.Xml.Linq.XElement(nsWpc + "wpc")

        If useBranchColumns Then
            Dim columnGap As System.Double = 7.0R
            Dim columnWidth As System.Double = (widthPoints - marginX * 2.0R - columnGap * (branchCount - 1)) / branchCount
            columnWidth = System.Math.Max(66.0R, System.Math.Min(95.0R, columnWidth))
            Dim columnsWidth As System.Double = columnWidth * branchCount + columnGap * (branchCount - 1)
            Dim columnsLeft As System.Double = (widthPoints - columnsWidth) / 2.0R
            Dim y As System.Double = 6.0R

            For depth As System.Int32 = 0 To branchDepth - 1
                If Not levels.ContainsKey(depth) Then Continue For
                Dim upperNodes As System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JObject) = levels(depth)
                Dim upperWidth As System.Double = System.Math.Min(175.0R, System.Math.Max(115.0R, widthPoints * 0.34R))
                Dim gap As System.Double = 8.0R
                Dim rowWidth As System.Double = upperNodes.Count * upperWidth + System.Math.Max(0, upperNodes.Count - 1) * gap
                Dim left As System.Double = (widthPoints - rowWidth) / 2.0R
                For i As System.Int32 = 0 To upperNodes.Count - 1
                    Dim node As Newtonsoft.Json.Linq.JObject = upperNodes(i)
                    layouts(GetVisualText(node, "id")) = New AutoPilotOoxmlNodeLayout() With {
                        .Node = node,
                        .Depth = depth,
                        .ParentId = GetVisualText(node, "parent_id"),
                        .Rect = New AutoPilotOoxmlRect() With {.X = left + i * (upperWidth + gap), .Y = y, .Width = upperWidth, .Height = boxH}
                    }
                Next
                y += boxH + verticalGap
            Next

            Dim branchY As System.Double = y + 4.0R
            Dim maxChildCount As System.Int32 = 0
            For i As System.Int32 = 0 To branchCount - 1
                Dim branch As Newtonsoft.Json.Linq.JObject = levels(branchDepth)(i)
                Dim branchId As System.String = GetVisualText(branch, "id")
                Dim x As System.Double = columnsLeft + i * (columnWidth + columnGap)
                layouts(branchId) = New AutoPilotOoxmlNodeLayout() With {
                    .Node = branch,
                    .Depth = branchDepth,
                    .ParentId = GetVisualText(branch, "parent_id"),
                    .Rect = New AutoPilotOoxmlRect() With {.X = x, .Y = branchY, .Width = columnWidth, .Height = boxH}
                }
                Dim descendants As New System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JObject)()
                For Each candidate As Newtonsoft.Json.Linq.JObject In nodes
                    If System.String.Equals(GetVisualText(candidate, "parent_id"), branchId, System.StringComparison.Ordinal) Then descendants.Add(candidate)
                Next
                maxChildCount = System.Math.Max(maxChildCount, descendants.Count)
                For childIndex As System.Int32 = 0 To descendants.Count - 1
                    Dim child As Newtonsoft.Json.Linq.JObject = descendants(childIndex)
                    layouts(GetVisualText(child, "id")) = New AutoPilotOoxmlNodeLayout() With {
                        .Node = child,
                        .Depth = branchDepth + 1,
                        .ParentId = branchId,
                        .Rect = New AutoPilotOoxmlRect() With {
                            .X = x,
                            .Y = branchY + boxH + 18.0R + childIndex * (smallBoxH + 7.0R),
                            .Width = columnWidth,
                            .Height = smallBoxH
                        }
                    }
                Next
            Next
            heightPoints = branchY + boxH + If(maxChildCount > 0, 18.0R + maxChildCount * smallBoxH + System.Math.Max(0, maxChildCount - 1) * 7.0R, 0.0R) + 8.0R

            ' Upper-level connectors.
            For Each layout As AutoPilotOoxmlNodeLayout In layouts.Values.OrderBy(Function(l As AutoPilotOoxmlNodeLayout) l.Depth)
                If System.String.IsNullOrWhiteSpace(layout.ParentId) OrElse Not layouts.ContainsKey(layout.ParentId) Then Continue For
                If layout.Depth > branchDepth Then Continue For
                Dim parentRect As AutoPilotOoxmlRect = layouts(layout.ParentId).Rect
                Dim childRect As AutoPilotOoxmlRect = layout.Rect
                Dim px As System.Double = parentRect.X + parentRect.Width / 2.0R
                Dim py As System.Double = parentRect.Y + parentRect.Height
                Dim cx As System.Double = childRect.X + childRect.Width / 2.0R
                Dim cy As System.Double = childRect.Y
                Dim midY As System.Double = py + System.Math.Max(5.0R, (cy - py) / 2.0R)
                canvas.Add(OoxmlCreateCanvasLine(shapeId, "Connector", px - 0.5R, py, 1.0R, System.Math.Max(1.0R, midY - py), "AAB2BD")) : shapeId += 1UI
                canvas.Add(OoxmlCreateCanvasLine(shapeId, "Connector", System.Math.Min(px, cx), midY - 0.5R, System.Math.Max(1.0R, System.Math.Abs(cx - px)), 1.0R, "AAB2BD")) : shapeId += 1UI
                canvas.Add(OoxmlCreateCanvasLine(shapeId, "Connector", cx - 0.5R, midY, 1.0R, System.Math.Max(1.0R, cy - midY), "AAB2BD")) : shapeId += 1UI
            Next

            ' Child connectors use a side trunk so lines never pass through stacked boxes.
            For i As System.Int32 = 0 To branchCount - 1
                Dim branch As Newtonsoft.Json.Linq.JObject = levels(branchDepth)(i)
                Dim branchId As System.String = GetVisualText(branch, "id")
                Dim childLayouts As System.Collections.Generic.List(Of AutoPilotOoxmlNodeLayout) = layouts.Values.Where(Function(l) System.String.Equals(l.ParentId, branchId, System.StringComparison.Ordinal)).OrderBy(Function(l As AutoPilotOoxmlNodeLayout) l.Rect.Y).ToList()
                If childLayouts.Count = 0 Then Continue For
                Dim parentRect As AutoPilotOoxmlRect = layouts(branchId).Rect
                Dim trunkX As System.Double = System.Math.Max(1.0R, parentRect.X - 4.5R)
                Dim trunkStartY As System.Double = parentRect.Y + parentRect.Height / 2.0R
                Dim trunkEndY As System.Double = childLayouts(childLayouts.Count - 1).Rect.Y + childLayouts(childLayouts.Count - 1).Rect.Height / 2.0R
                canvas.Add(OoxmlCreateCanvasLine(shapeId, "Connector", trunkX, trunkStartY, 1.0R, System.Math.Max(1.0R, trunkEndY - trunkStartY), "AAB2BD")) : shapeId += 1UI
                canvas.Add(OoxmlCreateCanvasLine(shapeId, "Connector", trunkX, trunkStartY, System.Math.Max(1.0R, parentRect.X - trunkX), 1.0R, "AAB2BD")) : shapeId += 1UI
                For Each childLayout As AutoPilotOoxmlNodeLayout In childLayouts
                    Dim cy As System.Double = childLayout.Rect.Y + childLayout.Rect.Height / 2.0R
                    canvas.Add(OoxmlCreateCanvasLine(shapeId, "Connector", trunkX, cy, System.Math.Max(1.0R, childLayout.Rect.X - trunkX), 1.0R, "AAB2BD")) : shapeId += 1UI
                Next
            Next
        Else
            Dim maxColumns As System.Int32 = 5
            Dim y As System.Double = 6.0R
            For Each kvp As System.Collections.Generic.KeyValuePair(Of System.Int32, System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JObject)) In levels
                Dim levelNodes As System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JObject) = kvp.Value
                Dim rows As System.Int32 = CInt(System.Math.Ceiling(levelNodes.Count / CDbl(maxColumns)))
                For row As System.Int32 = 0 To rows - 1
                    Dim count As System.Int32 = System.Math.Min(maxColumns, levelNodes.Count - row * maxColumns)
                    Dim gap As System.Double = 8.0R
                    Dim boxW As System.Double = System.Math.Min(120.0R, (widthPoints - marginX * 2.0R - gap * System.Math.Max(0, count - 1)) / count)
                    Dim rowWidth As System.Double = count * boxW + System.Math.Max(0, count - 1) * gap
                    Dim left As System.Double = (widthPoints - rowWidth) / 2.0R
                    For i As System.Int32 = 0 To count - 1
                        Dim node As Newtonsoft.Json.Linq.JObject = levelNodes(row * maxColumns + i)
                        layouts(GetVisualText(node, "id")) = New AutoPilotOoxmlNodeLayout() With {
                            .Node = node, .Depth = kvp.Key, .ParentId = GetVisualText(node, "parent_id"),
                            .Rect = New AutoPilotOoxmlRect() With {.X = left + i * (boxW + gap), .Y = y, .Width = boxW, .Height = boxH}
                        }
                    Next
                    y += boxH + 8.0R
                Next
                y += verticalGap
            Next
            heightPoints = y + 2.0R
            For Each layout As AutoPilotOoxmlNodeLayout In layouts.Values.OrderBy(Function(l As AutoPilotOoxmlNodeLayout) l.Depth)
                If System.String.IsNullOrWhiteSpace(layout.ParentId) OrElse Not layouts.ContainsKey(layout.ParentId) Then Continue For
                Dim pr As AutoPilotOoxmlRect = layouts(layout.ParentId).Rect
                Dim cr As AutoPilotOoxmlRect = layout.Rect
                Dim px As System.Double = pr.X + pr.Width / 2.0R
                Dim py As System.Double = pr.Y + pr.Height
                Dim cx As System.Double = cr.X + cr.Width / 2.0R
                Dim cy As System.Double = cr.Y
                Dim midY As System.Double = py + System.Math.Max(4.0R, (cy - py) / 2.0R)
                canvas.Add(OoxmlCreateCanvasLine(shapeId, "Connector", px - 0.5R, py, 1.0R, System.Math.Max(1.0R, midY - py), "AAB2BD")) : shapeId += 1UI
                canvas.Add(OoxmlCreateCanvasLine(shapeId, "Connector", System.Math.Min(px, cx), midY - 0.5R, System.Math.Max(1.0R, System.Math.Abs(cx - px)), 1.0R, "AAB2BD")) : shapeId += 1UI
                canvas.Add(OoxmlCreateCanvasLine(shapeId, "Connector", cx - 0.5R, midY, 1.0R, System.Math.Max(1.0R, cy - midY), "AAB2BD")) : shapeId += 1UI
            Next
        End If

        Dim lightFill As System.String = OoxmlLightenHex(accentHex, 0.91R)
        For Each layout As AutoPilotOoxmlNodeLayout In layouts.Values.OrderBy(Function(l As AutoPilotOoxmlNodeLayout) l.Depth).ThenBy(Function(l As AutoPilotOoxmlNodeLayout) l.Rect.X).ThenBy(Function(l As AutoPilotOoxmlNodeLayout) l.Rect.Y)
            Dim isRoot As System.Boolean = System.String.IsNullOrWhiteSpace(layout.ParentId)
            Dim fill As System.String = If(isRoot, accentHex, lightFill)
            Dim text As System.String = If(isRoot, "FFFFFF", "202124")
            canvas.Add(OoxmlCreateCanvasBox(shapeId,
                                           "Org " & GetVisualText(layout.Node, "id"),
                                           layout.Rect,
                                           fill,
                                           accentHex,
                                           GetVisualText(layout.Node, "label"),
                                           GetVisualText(layout.Node, "detail"),
                                           fontName,
                                           text,
                                           True))
            shapeId += 1UI
        Next
        heightPoints = System.Math.Min(390.0R, System.Math.Max(105.0R, heightPoints))
        Return canvas
    End Function

    Private Shared Function OoxmlBuildGenericDiagramCanvas(visual As Newtonsoft.Json.Linq.JObject,
                                                           visualType As System.String,
                                                           fontName As System.String,
                                                           accentHex As System.String,
                                                           ByRef widthPoints As System.Double,
                                                           ByRef heightPoints As System.Double) As System.Xml.Linq.XElement
        Dim nsWpc As System.Xml.Linq.XNamespace = OoxmlNsWordCanvas
        widthPoints = System.Math.Min(468.0R, GetVisualNumber(visual, "width_inches", 6.35R, 3.5R, 7.0R) * 72.0R)
        Dim canvas As New System.Xml.Linq.XElement(nsWpc + "wpc")
        Dim items As System.Collections.Generic.List(Of System.Tuple(Of System.String, System.String)) = GetWordVisualItems(visual)
        If items.Count = 0 Then items.Add(System.Tuple.Create(If(GetVisualText(visual, "title").Length > 0, GetVisualText(visual, "title"), "Visual"), System.String.Empty))
        If items.Count > 12 Then items = items.GetRange(0, 12)
        Dim shapeId As System.UInt32 = 1UI
        Dim fillLight As System.String = OoxmlLightenHex(accentHex, 0.91R)
        Dim lineHex As System.String = OoxmlLightenHex(accentHex, 0.35R)
        visualType = If(visualType, "process").ToLowerInvariant()

        Select Case visualType
            Case "timeline"
                heightPoints = 150.0R
                Dim left As System.Double = 26.0R
                Dim right As System.Double = widthPoints - 26.0R
                Dim centerY As System.Double = 74.0R
                canvas.Add(OoxmlCreateCanvasLine(shapeId, "Timeline", left, centerY, right - left, 1.5R, lineHex)) : shapeId += 1UI
                Dim stepX As System.Double = If(items.Count <= 1, 0.0R, (right - left) / (items.Count - 1))
                For i As System.Int32 = 0 To items.Count - 1
                    Dim cx As System.Double = left + i * stepX
                    Dim top As System.Double = If(i Mod 2 = 0, 10.0R, 88.0R)
                    Dim boxW As System.Double = System.Math.Min(92.0R, System.Math.Max(62.0R, widthPoints / System.Math.Max(4.5R, items.Count + 0.7R)))
                    Dim x As System.Double = System.Math.Max(2.0R, System.Math.Min(widthPoints - boxW - 2.0R, cx - boxW / 2.0R))
                    Dim rect As New AutoPilotOoxmlRect() With {.X = x, .Y = top, .Width = boxW, .Height = 49.0R}
                    canvas.Add(OoxmlCreateCanvasLine(shapeId, "Timeline tick", cx - 0.5R, System.Math.Min(centerY, top + 49.0R), 1.0R, System.Math.Max(1.0R, System.Math.Abs(centerY - (top + If(top < centerY, 49.0R, 0.0R)))), lineHex)) : shapeId += 1UI
                    canvas.Add(OoxmlCreateCanvasBox(shapeId, "Timeline item", rect, fillLight, accentHex, items(i).Item1, items(i).Item2, fontName, "202124", True)) : shapeId += 1UI
                Next

            Case "cycle", "relationship"
                heightPoints = 235.0R
                Dim centerX As System.Double = widthPoints / 2.0R
                Dim centerY As System.Double = heightPoints / 2.0R
                Dim radiusX As System.Double = System.Math.Min(155.0R, widthPoints * 0.34R)
                Dim radiusY As System.Double = 74.0R
                Dim boxW As System.Double = System.Math.Min(92.0R, widthPoints * 0.22R)
                Dim boxH As System.Double = 46.0R
                For i As System.Int32 = 0 To items.Count - 1
                    Dim angle As System.Double = -System.Math.PI / 2.0R + (2.0R * System.Math.PI * i / items.Count)
                    Dim x As System.Double = centerX + System.Math.Cos(angle) * radiusX - boxW / 2.0R
                    Dim y As System.Double = centerY + System.Math.Sin(angle) * radiusY - boxH / 2.0R
                    Dim rx As System.Double = centerX + System.Math.Cos(angle) * (radiusX - boxW * 0.55R)
                    Dim ry As System.Double = centerY + System.Math.Sin(angle) * (radiusY - boxH * 0.55R)
                    Dim horizontalStartX As System.Double = System.Math.Min(centerX, rx)
                    canvas.Add(OoxmlCreateCanvasLine(shapeId, "Relationship", horizontalStartX, centerY - 0.5R, System.Math.Max(1.0R, System.Math.Abs(rx - centerX)), 1.0R, lineHex)) : shapeId += 1UI
                    Dim verticalStartY As System.Double = System.Math.Min(centerY, ry)
                    canvas.Add(OoxmlCreateCanvasLine(shapeId, "Relationship", rx - 0.5R, verticalStartY, 1.0R, System.Math.Max(1.0R, System.Math.Abs(ry - centerY)), lineHex)) : shapeId += 1UI
                    canvas.Add(OoxmlCreateCanvasBox(shapeId, "Relationship item", New AutoPilotOoxmlRect() With {.X = x, .Y = y, .Width = boxW, .Height = boxH}, fillLight, accentHex, items(i).Item1, items(i).Item2, fontName, "202124", True)) : shapeId += 1UI
                Next

            Case "matrix"
                heightPoints = 215.0R
                Dim gap As System.Double = 9.0R
                Dim boxW As System.Double = (widthPoints - 24.0R - gap) / 2.0R
                Dim boxH As System.Double = 87.0R
                For i As System.Int32 = 0 To System.Math.Min(3, items.Count - 1)
                    Dim row As System.Int32 = i \ 2
                    Dim col As System.Int32 = i Mod 2
                    Dim rect As New AutoPilotOoxmlRect() With {.X = 12.0R + col * (boxW + gap), .Y = 10.0R + row * (boxH + gap), .Width = boxW, .Height = boxH}
                    canvas.Add(OoxmlCreateCanvasBox(shapeId, "Matrix item", rect, If(i = 0, accentHex, fillLight), accentHex, items(i).Item1, items(i).Item2, fontName, If(i = 0, "FFFFFF", "202124"), True)) : shapeId += 1UI
                Next

            Case "pyramid"
                heightPoints = 220.0R
                Dim levels As System.Int32 = items.Count
                Dim levelH As System.Double = System.Math.Min(48.0R, (heightPoints - 15.0R) / System.Math.Max(1, levels))
                For i As System.Int32 = 0 To levels - 1
                    Dim fraction As System.Double = 0.48R + 0.52R * ((i + 1.0R) / levels)
                    Dim boxW As System.Double = widthPoints * fraction
                    Dim rect As New AutoPilotOoxmlRect() With {.X = (widthPoints - boxW) / 2.0R, .Y = 7.0R + i * levelH, .Width = boxW, .Height = levelH - 4.0R}
                    canvas.Add(OoxmlCreateCanvasBox(shapeId, "Pyramid item", rect, If(i = levels - 1, accentHex, fillLight), accentHex, items(i).Item1, items(i).Item2, fontName, If(i = levels - 1, "FFFFFF", "202124"), True)) : shapeId += 1UI
                Next

            Case "list", "smartart", "diagram"
                Dim columns As System.Int32 = If(items.Count <= 4, 1, 2)
                Dim rows As System.Int32 = CInt(System.Math.Ceiling(items.Count / CDbl(columns)))
                Dim gap As System.Double = 8.0R
                Dim boxW As System.Double = (widthPoints - 18.0R - gap * (columns - 1)) / columns
                Dim boxH As System.Double = 50.0R
                heightPoints = rows * boxH + System.Math.Max(0, rows - 1) * gap + 16.0R
                For i As System.Int32 = 0 To items.Count - 1
                    Dim row As System.Int32 = i \ columns
                    Dim col As System.Int32 = i Mod columns
                    Dim rect As New AutoPilotOoxmlRect() With {.X = 9.0R + col * (boxW + gap), .Y = 8.0R + row * (boxH + gap), .Width = boxW, .Height = boxH}
                    canvas.Add(OoxmlCreateCanvasBox(shapeId, "List item", rect, If(i = 0, accentHex, fillLight), accentHex, items(i).Item1, items(i).Item2, fontName, If(i = 0, "FFFFFF", "202124"), True)) : shapeId += 1UI
                Next

            Case Else ' process
                Dim maxPerRow As System.Int32 = 5
                Dim rows As System.Int32 = CInt(System.Math.Ceiling(items.Count / CDbl(maxPerRow)))
                Dim rowH As System.Double = 58.0R
                heightPoints = rows * rowH + 12.0R
                For row As System.Int32 = 0 To rows - 1
                    Dim startIndex As System.Int32 = row * maxPerRow
                    Dim count As System.Int32 = System.Math.Min(maxPerRow, items.Count - startIndex)
                    Dim gap As System.Double = 18.0R
                    Dim boxW As System.Double = System.Math.Min(86.0R, (widthPoints - 16.0R - gap * System.Math.Max(0, count - 1)) / count)
                    Dim rowWidth As System.Double = count * boxW + System.Math.Max(0, count - 1) * gap
                    Dim left As System.Double = (widthPoints - rowWidth) / 2.0R
                    For i As System.Int32 = 0 To count - 1
                        Dim itemIndex As System.Int32 = startIndex + i
                        Dim rect As New AutoPilotOoxmlRect() With {.X = left + i * (boxW + gap), .Y = 7.0R + row * rowH, .Width = boxW, .Height = 43.0R}
                        canvas.Add(OoxmlCreateCanvasBox(shapeId, "Process item", rect, If(itemIndex = 0, accentHex, fillLight), accentHex, items(itemIndex).Item1, items(itemIndex).Item2, fontName, If(itemIndex = 0, "FFFFFF", "202124"), True)) : shapeId += 1UI
                        If i < count - 1 Then
                            Dim lineX As System.Double = rect.X + rect.Width + 3.0R
                            Dim lineW As System.Double = gap - 7.0R
                            Dim lineY As System.Double = rect.Y + rect.Height / 2.0R
                            canvas.Add(OoxmlCreateCanvasLine(shapeId, "Process arrow", lineX, lineY - 0.6R, lineW, 1.2R, accentHex)) : shapeId += 1UI
                            canvas.Add(OoxmlCreateCanvasTriangle(shapeId, "Process arrowhead", lineX + lineW - 2.0R, lineY - 3.0R, 6.0R, 6.0R, accentHex)) : shapeId += 1UI
                        End If
                    Next
                Next
        End Select

        heightPoints = System.Math.Min(360.0R, System.Math.Max(90.0R, heightPoints))
        Return canvas
    End Function

    Private Shared Function OoxmlCreateDiagramDrawing(visual As Newtonsoft.Json.Linq.JObject,
                                                      visualId As System.String,
                                                      visualType As System.String,
                                                      fontName As System.String,
                                                      accentHex As System.String,
                                                      drawingId As System.UInt32) As System.Xml.Linq.XElement
        Dim nsA As System.Xml.Linq.XNamespace = OoxmlNsDrawing
        Dim widthPoints As System.Double = 0.0R
        Dim heightPoints As System.Double = 0.0R
        Dim canvas As System.Xml.Linq.XElement
        If visualType = "org_chart" OrElse visualType = "hierarchy" Then
            canvas = OoxmlBuildOrgChartCanvas(visual, fontName, accentHex, widthPoints, heightPoints)
        Else
            canvas = OoxmlBuildGenericDiagramCanvas(visual, visualType, fontName, accentHex, widthPoints, heightPoints)
        End If
        Dim graphicData As New System.Xml.Linq.XElement(nsA + "graphicData",
                                                        New System.Xml.Linq.XAttribute("uri", OoxmlNsWordCanvas),
                                                        canvas)
        Dim placement As System.String = OoxmlGetVisualPlacement(visual)
        If placement = "auto" Then placement = "inline"
        Return OoxmlCreateOuterDrawing(graphicData, visualId, widthPoints, heightPoints, drawingId, placement)
    End Function

    Private Shared Function OoxmlCreateChartDrawing(visual As Newtonsoft.Json.Linq.JObject,
                                                    visualId As System.String,
                                                    relationshipId As System.String,
                                                    drawingId As System.UInt32) As System.Xml.Linq.XElement
        Dim nsA As System.Xml.Linq.XNamespace = OoxmlNsDrawing
        Dim nsC As System.Xml.Linq.XNamespace = OoxmlNsChart
        Dim nsR As System.Xml.Linq.XNamespace = OoxmlNsRelationship
        Dim widthPoints As System.Double = System.Math.Min(468.0R, GetVisualNumber(visual, "width_inches", 6.35R, 3.5R, 7.0R) * 72.0R)
        Dim heightPoints As System.Double = System.Math.Min(300.0R, GetVisualNumber(visual, "height_inches", 3.15R, 2.0R, 4.4R) * 72.0R)
        Dim chartRef As New System.Xml.Linq.XElement(nsC + "chart", New System.Xml.Linq.XAttribute(nsR + "id", relationshipId))
        Dim graphicData As New System.Xml.Linq.XElement(nsA + "graphicData",
                                                        New System.Xml.Linq.XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart"),
                                                        chartRef)
        Dim placement As System.String = OoxmlGetVisualPlacement(visual)
        If placement = "auto" Then placement = "inline"
        Return OoxmlCreateOuterDrawing(graphicData, visualId, widthPoints, heightPoints, drawingId, placement)
    End Function

    Private Shared Function OoxmlCreateChartDocument(visual As Newtonsoft.Json.Linq.JObject,
                                                     visualType As System.String,
                                                     categories As System.Collections.Generic.List(Of System.String),
                                                     series As System.Collections.Generic.List(Of System.Tuple(Of System.String, System.Collections.Generic.List(Of System.Double))),
                                                     accentHex As System.String) As System.Xml.Linq.XDocument
        Dim nsC As System.Xml.Linq.XNamespace = OoxmlNsChart
        Dim nsA As System.Xml.Linq.XNamespace = OoxmlNsDrawing
        Dim nsR As System.Xml.Linq.XNamespace = OoxmlNsRelationship
        Dim sheetFormulaName As System.String = OoxmlEscapeSheetName("Data")
        Dim categoryFormula As System.String = sheetFormulaName & "!$A$2:$A$" & (categories.Count + 1).ToString()
        Dim chartElement As System.Xml.Linq.XElement
        Dim plotArea As New System.Xml.Linq.XElement(nsC + "plotArea",
                                                     New System.Xml.Linq.XElement(nsC + "layout"))
        Dim axisIdCategory As System.UInt32 = 48650112UI
        Dim axisIdValue As System.UInt32 = 48672768UI

        Select Case visualType
            Case "line_chart"
                chartElement = New System.Xml.Linq.XElement(nsC + "lineChart",
                    New System.Xml.Linq.XElement(nsC + "grouping", New System.Xml.Linq.XAttribute("val", "standard")),
                    New System.Xml.Linq.XElement(nsC + "varyColors", New System.Xml.Linq.XAttribute("val", "0")))
            Case "area_chart"
                chartElement = New System.Xml.Linq.XElement(nsC + "areaChart",
                    New System.Xml.Linq.XElement(nsC + "grouping", New System.Xml.Linq.XAttribute("val", "standard")),
                    New System.Xml.Linq.XElement(nsC + "varyColors", New System.Xml.Linq.XAttribute("val", "0")))
            Case "pie_chart"
                chartElement = New System.Xml.Linq.XElement(nsC + "pieChart",
                    New System.Xml.Linq.XElement(nsC + "varyColors", New System.Xml.Linq.XAttribute("val", "1")))
            Case "doughnut_chart"
                chartElement = New System.Xml.Linq.XElement(nsC + "doughnutChart",
                    New System.Xml.Linq.XElement(nsC + "varyColors", New System.Xml.Linq.XAttribute("val", "1")),
                    New System.Xml.Linq.XElement(nsC + "holeSize", New System.Xml.Linq.XAttribute("val", "55")))
            Case Else
                chartElement = New System.Xml.Linq.XElement(nsC + "barChart",
                    New System.Xml.Linq.XElement(nsC + "barDir", New System.Xml.Linq.XAttribute("val", If(visualType = "bar_chart", "col", "col"))),
                    New System.Xml.Linq.XElement(nsC + "grouping", New System.Xml.Linq.XAttribute("val", "clustered")),
                    New System.Xml.Linq.XElement(nsC + "varyColors", New System.Xml.Linq.XAttribute("val", "0")))
        End Select

        For s As System.Int32 = 0 To series.Count - 1
            Dim values As System.Collections.Generic.List(Of System.Double) = series(s).Item2
            Dim count As System.Int32 = System.Math.Min(categories.Count, values.Count)
            Dim valueFormula As System.String = sheetFormulaName & "!$" & OoxmlExcelColumnName(s + 2) & "$2:$" & OoxmlExcelColumnName(s + 2) & "$" & (count + 1).ToString()
            Dim seriesNameFormula As System.String = sheetFormulaName & "!$" & OoxmlExcelColumnName(s + 2) & "$1"
            Dim ser As New System.Xml.Linq.XElement(nsC + "ser",
                New System.Xml.Linq.XElement(nsC + "idx", New System.Xml.Linq.XAttribute("val", s.ToString())),
                New System.Xml.Linq.XElement(nsC + "order", New System.Xml.Linq.XAttribute("val", s.ToString())),
                New System.Xml.Linq.XElement(nsC + "tx",
                    New System.Xml.Linq.XElement(nsC + "strRef",
                        New System.Xml.Linq.XElement(nsC + "f", seriesNameFormula),
                        New System.Xml.Linq.XElement(nsC + "strCache",
                            New System.Xml.Linq.XElement(nsC + "ptCount", New System.Xml.Linq.XAttribute("val", "1")),
                            New System.Xml.Linq.XElement(nsC + "pt", New System.Xml.Linq.XAttribute("idx", "0"), New System.Xml.Linq.XElement(nsC + "v", series(s).Item1))))),
                New System.Xml.Linq.XElement(nsC + "cat",
                    New System.Xml.Linq.XElement(nsC + "strRef",
                        New System.Xml.Linq.XElement(nsC + "f", categoryFormula),
                        OoxmlCreateChartStringCache(nsC, categories, count))),
                New System.Xml.Linq.XElement(nsC + "val",
                    New System.Xml.Linq.XElement(nsC + "numRef",
                        New System.Xml.Linq.XElement(nsC + "f", valueFormula),
                        OoxmlCreateChartNumberCache(nsC, values, count))))
            If visualType = "line_chart" Then
                ser.Add(New System.Xml.Linq.XElement(nsC + "marker", New System.Xml.Linq.XElement(nsC + "symbol", New System.Xml.Linq.XAttribute("val", "circle")), New System.Xml.Linq.XElement(nsC + "size", New System.Xml.Linq.XAttribute("val", "5"))))
                ser.Add(New System.Xml.Linq.XElement(nsC + "smooth", New System.Xml.Linq.XAttribute("val", "0")))
            End If
            chartElement.Add(ser)
        Next

        If visualType = "bar_chart" OrElse visualType = "column_chart" Then
            chartElement.Add(New System.Xml.Linq.XElement(nsC + "gapWidth", New System.Xml.Linq.XAttribute("val", "85")))
        End If
        If visualType <> "pie_chart" AndAlso visualType <> "doughnut_chart" Then
            chartElement.Add(New System.Xml.Linq.XElement(nsC + "axId", New System.Xml.Linq.XAttribute("val", axisIdCategory.ToString())))
            chartElement.Add(New System.Xml.Linq.XElement(nsC + "axId", New System.Xml.Linq.XAttribute("val", axisIdValue.ToString())))
        End If
        plotArea.Add(chartElement)

        If visualType <> "pie_chart" AndAlso visualType <> "doughnut_chart" Then
            plotArea.Add(New System.Xml.Linq.XElement(nsC + "catAx",
                New System.Xml.Linq.XElement(nsC + "axId", New System.Xml.Linq.XAttribute("val", axisIdCategory.ToString())),
                New System.Xml.Linq.XElement(nsC + "scaling", New System.Xml.Linq.XElement(nsC + "orientation", New System.Xml.Linq.XAttribute("val", "minMax"))),
                New System.Xml.Linq.XElement(nsC + "delete", New System.Xml.Linq.XAttribute("val", "0")),
                New System.Xml.Linq.XElement(nsC + "axPos", New System.Xml.Linq.XAttribute("val", "b")),
                New System.Xml.Linq.XElement(nsC + "tickLblPos", New System.Xml.Linq.XAttribute("val", "nextTo")),
                New System.Xml.Linq.XElement(nsC + "crossAx", New System.Xml.Linq.XAttribute("val", axisIdValue.ToString())),
                New System.Xml.Linq.XElement(nsC + "crosses", New System.Xml.Linq.XAttribute("val", "autoZero")),
                New System.Xml.Linq.XElement(nsC + "auto", New System.Xml.Linq.XAttribute("val", "1")),
                New System.Xml.Linq.XElement(nsC + "lblAlgn", New System.Xml.Linq.XAttribute("val", "ctr")),
                New System.Xml.Linq.XElement(nsC + "lblOffset", New System.Xml.Linq.XAttribute("val", "100"))))
            plotArea.Add(New System.Xml.Linq.XElement(nsC + "valAx",
                New System.Xml.Linq.XElement(nsC + "axId", New System.Xml.Linq.XAttribute("val", axisIdValue.ToString())),
                New System.Xml.Linq.XElement(nsC + "scaling", New System.Xml.Linq.XElement(nsC + "orientation", New System.Xml.Linq.XAttribute("val", "minMax"))),
                New System.Xml.Linq.XElement(nsC + "delete", New System.Xml.Linq.XAttribute("val", "0")),
                New System.Xml.Linq.XElement(nsC + "axPos", New System.Xml.Linq.XAttribute("val", "l")),
                New System.Xml.Linq.XElement(nsC + "majorGridlines"),
                New System.Xml.Linq.XElement(nsC + "numFmt", New System.Xml.Linq.XAttribute("formatCode", "General"), New System.Xml.Linq.XAttribute("sourceLinked", "1")),
                New System.Xml.Linq.XElement(nsC + "tickLblPos", New System.Xml.Linq.XAttribute("val", "nextTo")),
                New System.Xml.Linq.XElement(nsC + "crossAx", New System.Xml.Linq.XAttribute("val", axisIdCategory.ToString())),
                New System.Xml.Linq.XElement(nsC + "crosses", New System.Xml.Linq.XAttribute("val", "autoZero")),
                New System.Xml.Linq.XElement(nsC + "crossBetween", New System.Xml.Linq.XAttribute("val", "between"))))
        End If

        Dim chart As New System.Xml.Linq.XElement(nsC + "chart")
        Dim title As System.String = GetVisualText(visual, "title")
        If Not System.String.IsNullOrWhiteSpace(title) Then chart.Add(OoxmlCreateChartTitle(nsC, nsA, title))
        chart.Add(New System.Xml.Linq.XElement(nsC + "autoTitleDeleted", New System.Xml.Linq.XAttribute("val", If(System.String.IsNullOrWhiteSpace(title), "1", "0"))))
        chart.Add(plotArea)
        If series.Count > 1 Then
            chart.Add(New System.Xml.Linq.XElement(nsC + "legend",
                New System.Xml.Linq.XElement(nsC + "legendPos", New System.Xml.Linq.XAttribute("val", "b")),
                New System.Xml.Linq.XElement(nsC + "layout"),
                New System.Xml.Linq.XElement(nsC + "overlay", New System.Xml.Linq.XAttribute("val", "0"))))
        End If
        chart.Add(New System.Xml.Linq.XElement(nsC + "plotVisOnly", New System.Xml.Linq.XAttribute("val", "1")))
        chart.Add(New System.Xml.Linq.XElement(nsC + "dispBlanksAs", New System.Xml.Linq.XAttribute("val", "gap")))

        Dim chartSpace As New System.Xml.Linq.XElement(nsC + "chartSpace",
            New System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xmlns + "c", OoxmlNsChart),
            New System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xmlns + "a", OoxmlNsDrawing),
            New System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xmlns + "r", OoxmlNsRelationship),
            New System.Xml.Linq.XElement(nsC + "date1904", New System.Xml.Linq.XAttribute("val", "0")),
            New System.Xml.Linq.XElement(nsC + "lang", New System.Xml.Linq.XAttribute("val", "de-CH")),
            New System.Xml.Linq.XElement(nsC + "roundedCorners", New System.Xml.Linq.XAttribute("val", "0")),
            New System.Xml.Linq.XElement(nsC + "style", New System.Xml.Linq.XAttribute("val", "10")),
            chart,
            New System.Xml.Linq.XElement(nsC + "externalData", New System.Xml.Linq.XAttribute(nsR + "id", "rId1"), New System.Xml.Linq.XElement(nsC + "autoUpdate", New System.Xml.Linq.XAttribute("val", "0"))))
        Return New System.Xml.Linq.XDocument(New System.Xml.Linq.XDeclaration("1.0", "UTF-8", "yes"), chartSpace)
    End Function

    Private Shared Function OoxmlCreateChartStringCache(nsC As System.Xml.Linq.XNamespace,
                                                        values As System.Collections.Generic.List(Of System.String),
                                                        count As System.Int32) As System.Xml.Linq.XElement
        Dim cache As New System.Xml.Linq.XElement(nsC + "strCache", New System.Xml.Linq.XElement(nsC + "ptCount", New System.Xml.Linq.XAttribute("val", count.ToString())))
        For i As System.Int32 = 0 To count - 1
            cache.Add(New System.Xml.Linq.XElement(nsC + "pt", New System.Xml.Linq.XAttribute("idx", i.ToString()), New System.Xml.Linq.XElement(nsC + "v", values(i))))
        Next
        Return cache
    End Function

    Private Shared Function OoxmlCreateChartNumberCache(nsC As System.Xml.Linq.XNamespace,
                                                        values As System.Collections.Generic.List(Of System.Double),
                                                        count As System.Int32) As System.Xml.Linq.XElement
        Dim cache As New System.Xml.Linq.XElement(nsC + "numCache",
            New System.Xml.Linq.XElement(nsC + "formatCode", "General"),
            New System.Xml.Linq.XElement(nsC + "ptCount", New System.Xml.Linq.XAttribute("val", count.ToString())))
        For i As System.Int32 = 0 To count - 1
            cache.Add(New System.Xml.Linq.XElement(nsC + "pt", New System.Xml.Linq.XAttribute("idx", i.ToString()), New System.Xml.Linq.XElement(nsC + "v", values(i).ToString(System.Globalization.CultureInfo.InvariantCulture))))
        Next
        Return cache
    End Function

    Private Shared Function OoxmlCreateChartTitle(nsC As System.Xml.Linq.XNamespace,
                                                  nsA As System.Xml.Linq.XNamespace,
                                                  title As System.String) As System.Xml.Linq.XElement
        Return New System.Xml.Linq.XElement(nsC + "title",
            New System.Xml.Linq.XElement(nsC + "tx",
                New System.Xml.Linq.XElement(nsC + "rich",
                    New System.Xml.Linq.XElement(nsA + "bodyPr"),
                    New System.Xml.Linq.XElement(nsA + "lstStyle"),
                    New System.Xml.Linq.XElement(nsA + "p",
                        New System.Xml.Linq.XElement(nsA + "r",
                            New System.Xml.Linq.XElement(nsA + "rPr", New System.Xml.Linq.XAttribute("lang", "de-CH"), New System.Xml.Linq.XAttribute("sz", "1400"), New System.Xml.Linq.XAttribute("b", "1")),
                            New System.Xml.Linq.XElement(nsA + "t", title))))),
            New System.Xml.Linq.XElement(nsC + "layout"),
            New System.Xml.Linq.XElement(nsC + "overlay", New System.Xml.Linq.XAttribute("val", "0")))
    End Function

    Private Shared Function OoxmlCreateEmbeddedWorkbook(categories As System.Collections.Generic.List(Of System.String),
                                                        series As System.Collections.Generic.List(Of System.Tuple(Of System.String, System.Collections.Generic.List(Of System.Double)))) As System.Byte()
        Dim nsX As System.Xml.Linq.XNamespace = OoxmlNsSpreadsheet
        Dim nsR As System.Xml.Linq.XNamespace = OoxmlNsRelationship
        Dim nsRel As System.Xml.Linq.XNamespace = OoxmlNsPackageRelationship
        Dim nsCt As System.Xml.Linq.XNamespace = OoxmlNsContentTypes
        Using memory As New System.IO.MemoryStream()
            Using zip As New System.IO.Compression.ZipArchive(memory, System.IO.Compression.ZipArchiveMode.Create, True)
                Dim contentTypes As New System.Xml.Linq.XDocument(New System.Xml.Linq.XDeclaration("1.0", "UTF-8", "yes"),
                    New System.Xml.Linq.XElement(nsCt + "Types",
                        New System.Xml.Linq.XElement(nsCt + "Default", New System.Xml.Linq.XAttribute("Extension", "rels"), New System.Xml.Linq.XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                        New System.Xml.Linq.XElement(nsCt + "Default", New System.Xml.Linq.XAttribute("Extension", "xml"), New System.Xml.Linq.XAttribute("ContentType", "application/xml")),
                        New System.Xml.Linq.XElement(nsCt + "Override", New System.Xml.Linq.XAttribute("PartName", "/xl/workbook.xml"), New System.Xml.Linq.XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")),
                        New System.Xml.Linq.XElement(nsCt + "Override", New System.Xml.Linq.XAttribute("PartName", "/xl/worksheets/sheet1.xml"), New System.Xml.Linq.XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")),
                        New System.Xml.Linq.XElement(nsCt + "Override", New System.Xml.Linq.XAttribute("PartName", "/xl/styles.xml"), New System.Xml.Linq.XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"))))
                OoxmlWriteNewZipXml(zip, "[Content_Types].xml", contentTypes)
                Dim packageRels As New System.Xml.Linq.XDocument(New System.Xml.Linq.XDeclaration("1.0", "UTF-8", "yes"),
                    New System.Xml.Linq.XElement(nsRel + "Relationships",
                        New System.Xml.Linq.XElement(nsRel + "Relationship", New System.Xml.Linq.XAttribute("Id", "rId1"), New System.Xml.Linq.XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"), New System.Xml.Linq.XAttribute("Target", "xl/workbook.xml"))))
                OoxmlWriteNewZipXml(zip, "_rels/.rels", packageRels)
                Dim workbook As New System.Xml.Linq.XDocument(New System.Xml.Linq.XDeclaration("1.0", "UTF-8", "yes"),
                    New System.Xml.Linq.XElement(nsX + "workbook",
                        New System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xmlns + "r", OoxmlNsRelationship),
                        New System.Xml.Linq.XElement(nsX + "sheets",
                            New System.Xml.Linq.XElement(nsX + "sheet", New System.Xml.Linq.XAttribute("name", "Data"), New System.Xml.Linq.XAttribute("sheetId", "1"), New System.Xml.Linq.XAttribute(nsR + "id", "rId1")))))
                OoxmlWriteNewZipXml(zip, "xl/workbook.xml", workbook)
                Dim workbookRels As New System.Xml.Linq.XDocument(New System.Xml.Linq.XDeclaration("1.0", "UTF-8", "yes"),
                    New System.Xml.Linq.XElement(nsRel + "Relationships",
                        New System.Xml.Linq.XElement(nsRel + "Relationship", New System.Xml.Linq.XAttribute("Id", "rId1"), New System.Xml.Linq.XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"), New System.Xml.Linq.XAttribute("Target", "worksheets/sheet1.xml")),
                        New System.Xml.Linq.XElement(nsRel + "Relationship", New System.Xml.Linq.XAttribute("Id", "rId2"), New System.Xml.Linq.XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"), New System.Xml.Linq.XAttribute("Target", "styles.xml"))))
                OoxmlWriteNewZipXml(zip, "xl/_rels/workbook.xml.rels", workbookRels)

                Dim sheetData As New System.Xml.Linq.XElement(nsX + "sheetData")
                Dim header As New System.Xml.Linq.XElement(nsX + "row", New System.Xml.Linq.XAttribute("r", "1"))
                header.Add(OoxmlCreateInlineStringCell(nsX, "A1", "Category"))
                For s As System.Int32 = 0 To series.Count - 1
                    header.Add(OoxmlCreateInlineStringCell(nsX, OoxmlExcelColumnName(s + 2) & "1", series(s).Item1))
                Next
                sheetData.Add(header)
                For i As System.Int32 = 0 To categories.Count - 1
                    Dim rowIndex As System.Int32 = i + 2
                    Dim row As New System.Xml.Linq.XElement(nsX + "row", New System.Xml.Linq.XAttribute("r", rowIndex.ToString()))
                    row.Add(OoxmlCreateInlineStringCell(nsX, "A" & rowIndex.ToString(), categories(i)))
                    For s As System.Int32 = 0 To series.Count - 1
                        Dim value As System.Double = If(i < series(s).Item2.Count, series(s).Item2.Item(i), 0.0R)
                        row.Add(New System.Xml.Linq.XElement(nsX + "c", New System.Xml.Linq.XAttribute("r", OoxmlExcelColumnName(s + 2) & rowIndex.ToString()), New System.Xml.Linq.XElement(nsX + "v", value.ToString(System.Globalization.CultureInfo.InvariantCulture))))
                    Next
                    sheetData.Add(row)
                Next
                Dim worksheet As New System.Xml.Linq.XDocument(New System.Xml.Linq.XDeclaration("1.0", "UTF-8", "yes"),
                    New System.Xml.Linq.XElement(nsX + "worksheet",
                        New System.Xml.Linq.XElement(nsX + "dimension", New System.Xml.Linq.XAttribute("ref", "A1:" & OoxmlExcelColumnName(series.Count + 1) & (categories.Count + 1).ToString())),
                        New System.Xml.Linq.XElement(nsX + "sheetViews", New System.Xml.Linq.XElement(nsX + "sheetView", New System.Xml.Linq.XAttribute("workbookViewId", "0"))),
                        New System.Xml.Linq.XElement(nsX + "sheetFormatPr", New System.Xml.Linq.XAttribute("defaultRowHeight", "15")),
                        sheetData))
                OoxmlWriteNewZipXml(zip, "xl/worksheets/sheet1.xml", worksheet)
                Dim styles As New System.Xml.Linq.XDocument(New System.Xml.Linq.XDeclaration("1.0", "UTF-8", "yes"),
                    New System.Xml.Linq.XElement(nsX + "styleSheet",
                        New System.Xml.Linq.XElement(nsX + "fonts", New System.Xml.Linq.XAttribute("count", "1"), New System.Xml.Linq.XElement(nsX + "font", New System.Xml.Linq.XElement(nsX + "sz", New System.Xml.Linq.XAttribute("val", "11")), New System.Xml.Linq.XElement(nsX + "name", New System.Xml.Linq.XAttribute("val", "Aptos")))),
                        New System.Xml.Linq.XElement(nsX + "fills", New System.Xml.Linq.XAttribute("count", "2"), New System.Xml.Linq.XElement(nsX + "fill", New System.Xml.Linq.XElement(nsX + "patternFill", New System.Xml.Linq.XAttribute("patternType", "none"))), New System.Xml.Linq.XElement(nsX + "fill", New System.Xml.Linq.XElement(nsX + "patternFill", New System.Xml.Linq.XAttribute("patternType", "gray125")))),
                        New System.Xml.Linq.XElement(nsX + "borders", New System.Xml.Linq.XAttribute("count", "1"), New System.Xml.Linq.XElement(nsX + "border", New System.Xml.Linq.XElement(nsX + "left"), New System.Xml.Linq.XElement(nsX + "right"), New System.Xml.Linq.XElement(nsX + "top"), New System.Xml.Linq.XElement(nsX + "bottom"), New System.Xml.Linq.XElement(nsX + "diagonal"))),
                        New System.Xml.Linq.XElement(nsX + "cellStyleXfs", New System.Xml.Linq.XAttribute("count", "1"), New System.Xml.Linq.XElement(nsX + "xf", New System.Xml.Linq.XAttribute("numFmtId", "0"), New System.Xml.Linq.XAttribute("fontId", "0"), New System.Xml.Linq.XAttribute("fillId", "0"), New System.Xml.Linq.XAttribute("borderId", "0"))),
                        New System.Xml.Linq.XElement(nsX + "cellXfs", New System.Xml.Linq.XAttribute("count", "1"), New System.Xml.Linq.XElement(nsX + "xf", New System.Xml.Linq.XAttribute("numFmtId", "0"), New System.Xml.Linq.XAttribute("fontId", "0"), New System.Xml.Linq.XAttribute("fillId", "0"), New System.Xml.Linq.XAttribute("borderId", "0"), New System.Xml.Linq.XAttribute("xfId", "0")))))
                OoxmlWriteNewZipXml(zip, "xl/styles.xml", styles)
            End Using
            Return memory.ToArray()
        End Using
    End Function

    Private Shared Function OoxmlCreateInlineStringCell(nsX As System.Xml.Linq.XNamespace,
                                                        reference As System.String,
                                                        value As System.String) As System.Xml.Linq.XElement
        Return New System.Xml.Linq.XElement(nsX + "c",
            New System.Xml.Linq.XAttribute("r", reference),
            New System.Xml.Linq.XAttribute("t", "inlineStr"),
            New System.Xml.Linq.XElement(nsX + "is", New System.Xml.Linq.XElement(nsX + "t", If(value, System.String.Empty))))
    End Function

    Private Shared Sub OoxmlWriteNewZipXml(zip As System.IO.Compression.ZipArchive,
                                           entryName As System.String,
                                           document As System.Xml.Linq.XDocument)
        Dim entry As System.IO.Compression.ZipArchiveEntry = zip.CreateEntry(entryName, System.IO.Compression.CompressionLevel.Optimal)
        Using stream As System.IO.Stream = entry.Open()
            Dim settings As New System.Xml.XmlWriterSettings() With {.Encoding = New System.Text.UTF8Encoding(False), .Indent = False, .OmitXmlDeclaration = False}
            Using writer As System.Xml.XmlWriter = System.Xml.XmlWriter.Create(stream, settings)
                document.Save(writer)
            End Using
        End Using
    End Sub

    Private Shared Sub OoxmlEnsureContentTypeOverride(contentTypes As System.Xml.Linq.XDocument,
                                                       partName As System.String,
                                                       contentType As System.String)
        Dim nsCt As System.Xml.Linq.XNamespace = OoxmlNsContentTypes
        Dim existing As System.Xml.Linq.XElement = contentTypes.Root.Elements(nsCt + "Override").FirstOrDefault(Function(e) System.String.Equals(CStr(e.Attribute("PartName")), partName, System.StringComparison.OrdinalIgnoreCase))
        If existing Is Nothing Then
            contentTypes.Root.Add(New System.Xml.Linq.XElement(nsCt + "Override", New System.Xml.Linq.XAttribute("PartName", partName), New System.Xml.Linq.XAttribute("ContentType", contentType)))
        Else
            existing.SetAttributeValue("ContentType", contentType)
        End If
    End Sub

    Private Shared Function InsertAutoPilotWordVisualsOpenXml(outputPath As System.String,
                                                              visuals As Newtonsoft.Json.Linq.JArray,
                                                              fontName As System.String,
                                                              accentHexRaw As System.String,
                                                              ByRef embeddedCount As System.Int32,
                                                              ByRef warnings As System.Collections.Generic.List(Of System.String)) As System.Boolean
        embeddedCount = 0
        If warnings Is Nothing Then warnings = New System.Collections.Generic.List(Of System.String)()
        If visuals Is Nothing OrElse visuals.Count = 0 Then Return True
        If System.String.IsNullOrWhiteSpace(outputPath) OrElse Not System.IO.File.Exists(outputPath) Then
            warnings.Add("OOXML visual renderer could not find the saved Word document.")
            Return False
        End If

        Dim accentHex As System.String = OoxmlNormalizeHex(accentHexRaw, "17365D")
        If System.String.IsNullOrWhiteSpace(fontName) Then fontName = "Aptos"
        Try
            Using fileStream As New System.IO.FileStream(outputPath, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
                Using archive As New System.IO.Compression.ZipArchive(fileStream, System.IO.Compression.ZipArchiveMode.Update, True)
                    Dim documentXml As System.Xml.Linq.XDocument = OoxmlLoadZipXml(archive, "word/document.xml")
                    If documentXml Is Nothing OrElse documentXml.Root Is Nothing Then
                        warnings.Add("DOCX has no word/document.xml part.")
                        Return False
                    End If
                    Dim rels As System.Xml.Linq.XDocument = OoxmlEnsureDocumentRelationships(archive)
                    Dim contentTypes As System.Xml.Linq.XDocument = OoxmlLoadZipXml(archive, "[Content_Types].xml")
                    If contentTypes Is Nothing OrElse contentTypes.Root Is Nothing Then
                        warnings.Add("DOCX has no [Content_Types].xml part.")
                        Return False
                    End If
                    OoxmlEnsureNamespaceDeclarations(documentXml)
                    Dim drawingId As System.UInt32 = OoxmlNextDrawingId(documentXml)
                    Dim existingChartIndex As System.Int32 = System.Linq.Enumerable.Count(Of System.IO.Compression.ZipArchiveEntry)(archive.Entries, Function(e As System.IO.Compression.ZipArchiveEntry) System.Text.RegularExpressions.Regex.IsMatch(e.FullName, "^word/charts/chart\d+\.xml$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    Dim chartIndex As System.Int32 = existingChartIndex + 1
                    Dim nsRel As System.Xml.Linq.XNamespace = OoxmlNsPackageRelationship

                    For Each token As Newtonsoft.Json.Linq.JToken In visuals
                        If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then
                            warnings.Add("Invalid visual entry during OOXML rendering.")
                            Return False
                        End If
                        Dim visual As Newtonsoft.Json.Linq.JObject = DirectCast(token, Newtonsoft.Json.Linq.JObject)
                        Dim visualId As System.String = GetVisualText(visual, "id")
                        Dim visualType As System.String = GetVisualText(visual, "type", "process").ToLowerInvariant()
                        Dim placeholderError As System.String = System.String.Empty
                        Dim paragraph As System.Xml.Linq.XElement = OoxmlFindPlaceholderParagraph(documentXml, visualId, placeholderError)
                        If paragraph Is Nothing Then
                            warnings.Add(placeholderError)
                            Return False
                        End If

                        If OoxmlIsChartType(visualType) Then
                            Dim categories As System.Collections.Generic.List(Of System.String) = GetWordVisualCategories(visual)
                            Dim series As System.Collections.Generic.List(Of System.Tuple(Of System.String, System.Collections.Generic.List(Of System.Double))) = GetWordVisualSeries(visual)
                            If categories.Count = 0 OrElse series.Count = 0 Then
                                warnings.Add("Chart visual '" & visualId & "' requires categories and numeric series.")
                                Return False
                            End If
                            Dim usableCount As System.Int32 = categories.Count
                            For Each s As System.Tuple(Of System.String, System.Collections.Generic.List(Of System.Double)) In series
                                usableCount = System.Math.Min(usableCount, s.Item2.Count)
                            Next
                            If usableCount <= 0 Then
                                warnings.Add("Chart visual '" & visualId & "' has no aligned category/value rows.")
                                Return False
                            End If
                            If usableCount < categories.Count Then categories = categories.GetRange(0, usableCount)
                            For i As System.Int32 = 0 To series.Count - 1
                                If series(i).Item2.Count > usableCount Then series(i) = New System.Tuple(Of System.String, System.Collections.Generic.List(Of System.Double))(series(i).Item1, series(i).Item2.GetRange(0, usableCount))
                            Next

                            Dim relId As System.String = OoxmlNextRelationshipId(rels)
                            Dim chartName As System.String = "chart" & chartIndex.ToString() & ".xml"
                            Dim chartPartName As System.String = "word/charts/" & chartName
                            Dim workbookName As System.String = "Microsoft_Excel_Worksheet" & chartIndex.ToString() & ".xlsx"
                            Dim workbookPartName As System.String = "word/embeddings/" & workbookName
                            rels.Root.Add(New System.Xml.Linq.XElement(nsRel + "Relationship",
                                New System.Xml.Linq.XAttribute("Id", relId),
                                New System.Xml.Linq.XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"),
                                New System.Xml.Linq.XAttribute("Target", "charts/" & chartName)))
                            OoxmlEnsureContentTypeOverride(contentTypes, "/" & chartPartName, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
                            OoxmlEnsureContentTypeOverride(contentTypes, "/" & workbookPartName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            OoxmlReplaceZipEntry(archive, chartPartName, OoxmlCreateChartDocument(visual, visualType, categories, series, accentHex))
                            Dim chartRels As New System.Xml.Linq.XDocument(New System.Xml.Linq.XDeclaration("1.0", "UTF-8", "yes"),
                                New System.Xml.Linq.XElement(nsRel + "Relationships",
                                    New System.Xml.Linq.XElement(nsRel + "Relationship",
                                        New System.Xml.Linq.XAttribute("Id", "rId1"),
                                        New System.Xml.Linq.XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"),
                                        New System.Xml.Linq.XAttribute("Target", "../embeddings/" & workbookName))))
                            OoxmlReplaceZipEntry(archive, "word/charts/_rels/" & chartName & ".rels", chartRels)
                            OoxmlReplaceZipEntryBytes(archive, workbookPartName, OoxmlCreateEmbeddedWorkbook(categories, series))
                            OoxmlReplacePlaceholderParagraph(paragraph, OoxmlCreateChartDrawing(visual, visualId, relId, drawingId))
                            chartIndex += 1
                        ElseIf OoxmlIsDiagramType(visualType) Then
                            OoxmlReplacePlaceholderParagraph(paragraph, OoxmlCreateDiagramDrawing(visual, visualId, visualType, fontName, accentHex, drawingId))
                        Else
                            warnings.Add("Unsupported native OOXML Word visual type '" & visualType & "' for '" & visualId & "'.")
                            Return False
                        End If
                        drawingId += 1UI
                        embeddedCount += 1
                    Next

                    OoxmlReplaceZipEntry(archive, "word/document.xml", documentXml)
                    OoxmlReplaceZipEntry(archive, "word/_rels/document.xml.rels", rels)
                    OoxmlReplaceZipEntry(archive, "[Content_Types].xml", contentTypes)
                End Using
            End Using
        Catch ex As System.Exception
            warnings.Add("Native OOXML visual insertion failed: " & ex.Message)
            Return False
        End Try
        Return embeddedCount = visuals.Count
    End Function

    Private Shared Function ValidateSavedAutoPilotWordVisualPersistenceOpenXml(outputPath As System.String,
                                                                               visuals As Newtonsoft.Json.Linq.JArray,
                                                                               ByRef validationError As System.String) As System.Boolean
        validationError = System.String.Empty
        If visuals Is Nothing OrElse visuals.Count = 0 Then Return True
        If System.String.IsNullOrWhiteSpace(outputPath) OrElse Not System.IO.File.Exists(outputPath) Then
            validationError = "Saved Word document is missing."
            Return False
        End If
        Try
            Using archive As System.IO.Compression.ZipArchive = System.IO.Compression.ZipFile.OpenRead(outputPath)
                Dim documentXml As System.Xml.Linq.XDocument = OoxmlLoadZipXml(archive, "word/document.xml")
                If documentXml Is Nothing Then
                    validationError = "Saved DOCX contains no word/document.xml."
                    Return False
                End If
                Dim nsWp As System.Xml.Linq.XNamespace = OoxmlNsWordDrawing
                Dim markers As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                For Each docPr As System.Xml.Linq.XElement In documentXml.Descendants(nsWp + "docPr")
                    Dim marker As System.String = CStr(docPr.Attribute("name"))
                    If Not System.String.IsNullOrWhiteSpace(marker) Then markers.Add(marker)
                Next
                Dim chartCount As System.Int32 = System.Linq.Enumerable.Count(Of System.IO.Compression.ZipArchiveEntry)(archive.Entries, Function(e As System.IO.Compression.ZipArchiveEntry) System.Text.RegularExpressions.Regex.IsMatch(e.FullName, "^word/charts/chart\d+\.xml$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                Dim workbookCount As System.Int32 = System.Linq.Enumerable.Count(Of System.IO.Compression.ZipArchiveEntry)(archive.Entries, Function(e As System.IO.Compression.ZipArchiveEntry) System.Text.RegularExpressions.Regex.IsMatch(e.FullName, "^word/embeddings/.+\.xlsx$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                Dim nsA As System.Xml.Linq.XNamespace = OoxmlNsDrawing
                Dim canvasCount As System.Int32 = System.Linq.Enumerable.Count(Of System.Xml.Linq.XElement)(documentXml.Descendants(nsA + "graphicData"), Function(g As System.Xml.Linq.XElement) System.String.Equals(CStr(g.Attribute("uri")), OoxmlNsWordCanvas, System.StringComparison.Ordinal))
                Dim expectedCharts As System.Int32 = 0
                Dim expectedDiagrams As System.Int32 = 0
                For Each token As Newtonsoft.Json.Linq.JToken In visuals
                    Dim visual As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
                    If visual Is Nothing Then Continue For
                    Dim id As System.String = GetVisualText(visual, "id")
                    Dim visualType As System.String = GetVisualText(visual, "type", "process").ToLowerInvariant()
                    If Not markers.Contains("RedInk Visual " & id) Then
                        validationError = "Saved DOCX does not contain the exact native visual marker for '" & id & "'."
                        Return False
                    End If
                    If OoxmlIsChartType(visualType) Then
                        expectedCharts += 1
                    ElseIf OoxmlIsDiagramType(visualType) Then
                        expectedDiagrams += 1
                    End If
                Next
                If chartCount < expectedCharts OrElse workbookCount < expectedCharts Then
                    validationError = "Saved DOCX persistence check failed: expected at least " & expectedCharts.ToString() & " native chart part(s) and embedded workbook(s); found charts=" & chartCount.ToString() & ", workbooks=" & workbookCount.ToString() & "."
                    Return False
                End If
                If canvasCount < expectedDiagrams Then
                    validationError = "Saved DOCX persistence check failed: expected at least " & expectedDiagrams.ToString() & " editable DrawingML diagram canvas(es); found " & canvasCount.ToString() & "."
                    Return False
                End If
                Dim fullText As System.String = System.String.Concat(documentXml.Descendants(System.Xml.Linq.XName.Get("t", OoxmlNsWord)).Select(Function(t) t.Value))
                If System.Text.RegularExpressions.Regex.IsMatch(fullText, "\[\[visual:[A-Za-z0-9_.-]{1,64}\]\]") Then
                    validationError = "Saved DOCX still contains an unresolved [[visual:ID]] placeholder."
                    Return False
                End If
            End Using
            Return True
        Catch ex As System.Exception
            validationError = "Saved DOCX native visual validation failed: " & ex.Message
            Return False
        End Try
    End Function

End Class
