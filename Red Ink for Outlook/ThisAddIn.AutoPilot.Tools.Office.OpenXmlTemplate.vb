' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Tools.Office.OpenXmlTemplate.vb
' Purpose:
'   Deterministic OOXML-only renderer for slot-bound create_word_document
'   templates. This path never starts Word and never performs COM automation.
' =============================================================================

Option Explicit On
Option Strict Off

Partial Public Class ThisAddIn

    Private Shared ReadOnly AutoPilotWordMainNs As System.Xml.Linq.XNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    Private Shared ReadOnly AutoPilotWordXmlNs As System.Xml.Linq.XNamespace = "http://www.w3.org/XML/1998/namespace"

    Private Shared Function GetAutoPilotWordOpenXmlParagraphText(paragraph As System.Xml.Linq.XElement) As System.String
        If paragraph Is Nothing Then Return System.String.Empty
        Dim builder As New System.Text.StringBuilder()
        For Each node As System.Xml.Linq.XElement In paragraph.Descendants()
            If node.Name = AutoPilotWordMainNs + "t" Then
                builder.Append(If(node.Value, System.String.Empty))
            ElseIf node.Name = AutoPilotWordMainNs + "br" Then
                builder.AppendLine()
            ElseIf node.Name = AutoPilotWordMainNs + "tab" Then
                builder.Append(ControlChars.Tab)
            End If
        Next
        Return builder.ToString()
    End Function

    Private Shared Sub SetAutoPilotWordOpenXmlTextNodeValue(textNode As System.Xml.Linq.XElement,
                                                               value As System.String)
        If textNode Is Nothing Then Return
        Dim normalized As System.String = If(value, System.String.Empty).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        Dim parts() As System.String = normalized.Split(ControlChars.Lf)
        textNode.Value = If(parts.Length > 0, parts(0), System.String.Empty)
        If textNode.Value.Length > 0 AndAlso
           (System.Char.IsWhiteSpace(textNode.Value(0)) OrElse System.Char.IsWhiteSpace(textNode.Value(textNode.Value.Length - 1))) Then
            textNode.SetAttributeValue(AutoPilotWordXmlNs + "space", "preserve")
        Else
            textNode.SetAttributeValue(AutoPilotWordXmlNs + "space", Nothing)
        End If
        If parts.Length <= 1 Then Return

        Dim anchor As System.Xml.Linq.XElement = textNode
        For i As System.Int32 = 1 To parts.Length - 1
            Dim lineBreak As New System.Xml.Linq.XElement(AutoPilotWordMainNs + "br")
            anchor.AddAfterSelf(lineBreak)
            Dim nextText As New System.Xml.Linq.XElement(AutoPilotWordMainNs + "t", parts(i))
            If parts(i).Length > 0 AndAlso
               (System.Char.IsWhiteSpace(parts(i)(0)) OrElse System.Char.IsWhiteSpace(parts(i)(parts(i).Length - 1))) Then
                nextText.SetAttributeValue(AutoPilotWordXmlNs + "space", "preserve")
            End If
            lineBreak.AddAfterSelf(nextText)
            anchor = nextText
        Next
    End Sub

    Private Shared Function ReplaceAutoPilotWordOpenXmlPlaceholderInParagraph(
            paragraph As System.Xml.Linq.XElement,
            placeholder As System.String,
            replacement As System.String) As System.Int32

        If paragraph Is Nothing OrElse System.String.IsNullOrEmpty(placeholder) Then Return 0
        Dim replacementValue As System.String = If(replacement, System.String.Empty)
        Dim replaced As System.Int32 = 0

        Do
            Dim textNodes As System.Collections.Generic.List(Of System.Xml.Linq.XElement) = paragraph.Descendants(AutoPilotWordMainNs + "t").ToList()
            If textNodes.Count = 0 Then Exit Do

            Dim fullText As New System.Text.StringBuilder()
            Dim starts As New System.Collections.Generic.List(Of System.Int32)()
            For Each textNode As System.Xml.Linq.XElement In textNodes
                starts.Add(fullText.Length)
                fullText.Append(If(textNode.Value, System.String.Empty))
            Next

            Dim whole As System.String = fullText.ToString()
            Dim matchStart As System.Int32 = whole.IndexOf(placeholder, System.StringComparison.OrdinalIgnoreCase)
            If matchStart < 0 Then Exit Do
            Dim matchEndExclusive As System.Int32 = matchStart + placeholder.Length

            Dim firstIndex As System.Int32 = -1
            Dim lastIndex As System.Int32 = -1
            For index As System.Int32 = 0 To textNodes.Count - 1
                Dim nodeStart As System.Int32 = starts(index)
                Dim nodeEnd As System.Int32 = nodeStart + If(textNodes(index).Value, System.String.Empty).Length
                If firstIndex < 0 AndAlso matchStart < nodeEnd AndAlso matchEndExclusive > nodeStart Then firstIndex = index
                If matchEndExclusive > nodeStart AndAlso matchStart < nodeEnd Then lastIndex = index
            Next
            If firstIndex < 0 OrElse lastIndex < 0 Then Exit Do

            Dim firstNode As System.Xml.Linq.XElement = textNodes(firstIndex)
            Dim lastNode As System.Xml.Linq.XElement = textNodes(lastIndex)
            Dim firstStart As System.Int32 = starts(firstIndex)
            Dim lastStart As System.Int32 = starts(lastIndex)
            Dim firstValue As System.String = If(firstNode.Value, System.String.Empty)
            Dim lastValue As System.String = If(lastNode.Value, System.String.Empty)
            Dim prefixLength As System.Int32 = System.Math.Max(0, matchStart - firstStart)
            Dim suffixOffset As System.Int32 = System.Math.Max(0, matchEndExclusive - lastStart)
            Dim prefix As System.String = firstValue.Substring(0, System.Math.Min(prefixLength, firstValue.Length))
            Dim suffix As System.String = If(suffixOffset <= lastValue.Length, lastValue.Substring(suffixOffset), System.String.Empty)

            If firstIndex = lastIndex Then
                SetAutoPilotWordOpenXmlTextNodeValue(firstNode, prefix & replacementValue & suffix)
            Else
                SetAutoPilotWordOpenXmlTextNodeValue(firstNode, prefix & replacementValue)
                For index As System.Int32 = firstIndex + 1 To lastIndex - 1
                    textNodes(index).Value = System.String.Empty
                Next
                lastNode.Value = suffix
            End If
            replaced += 1
        Loop

        Return replaced
    End Function

    Private Shared Function GetAutoPilotWordOpenXmlStoryEntryNames(archive As System.IO.Compression.ZipArchive) As System.Collections.Generic.List(Of System.String)
        Dim result As New System.Collections.Generic.List(Of System.String)()
        If archive Is Nothing Then Return result

        For Each entry As System.IO.Compression.ZipArchiveEntry In archive.Entries
            Dim name As System.String = If(entry.FullName, System.String.Empty).Replace("\", "/")
            If System.String.Equals(name, "word/document.xml", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.Text.RegularExpressions.Regex.IsMatch(name, "^word/header[0-9]+\.xml$", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant) OrElse
               System.Text.RegularExpressions.Regex.IsMatch(name, "^word/footer[0-9]+\.xml$", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant) OrElse
               System.String.Equals(name, "word/footnotes.xml", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(name, "word/endnotes.xml", System.StringComparison.OrdinalIgnoreCase) Then
                result.Add(entry.FullName)
            End If
        Next
        Return result
    End Function

    Private Shared Function LoadAutoPilotWordOpenXmlEntry(archive As System.IO.Compression.ZipArchive, entryName As System.String) As System.Xml.Linq.XDocument
        Dim entry As System.IO.Compression.ZipArchiveEntry = archive.GetEntry(entryName)
        If entry Is Nothing Then Return Nothing
        Using input As System.IO.Stream = entry.Open()
            Return System.Xml.Linq.XDocument.Load(input, System.Xml.Linq.LoadOptions.PreserveWhitespace)
        End Using
    End Function

    Private Shared Sub SaveAutoPilotWordOpenXmlEntry(archive As System.IO.Compression.ZipArchive, entryName As System.String, xml As System.Xml.Linq.XDocument)
        Dim existing As System.IO.Compression.ZipArchiveEntry = archive.GetEntry(entryName)
        If existing IsNot Nothing Then existing.Delete()
        Dim replacement As System.IO.Compression.ZipArchiveEntry = archive.CreateEntry(entryName, System.IO.Compression.CompressionLevel.Optimal)
        Using output As System.IO.Stream = replacement.Open()
            xml.Save(output, System.Xml.Linq.SaveOptions.DisableFormatting)
        End Using
    End Sub

    Private Shared Function BuildAutoPilotWordOpenXmlStyleIdMap(
            archive As System.IO.Compression.ZipArchive,
            ByRef styleMap As System.Collections.Generic.Dictionary(Of System.String, System.String),
            ByRef validationError As System.String) As System.Boolean

        styleMap = New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
        validationError = System.String.Empty
        Dim stylesXml As System.Xml.Linq.XDocument = LoadAutoPilotWordOpenXmlEntry(archive, "word/styles.xml")
        If stylesXml Is Nothing OrElse stylesXml.Root Is Nothing Then
            validationError = "The selected Word template has no readable word/styles.xml part."
            Return False
        End If

        For Each style As System.Xml.Linq.XElement In stylesXml.Descendants(AutoPilotWordMainNs + "style")
            Dim typeAttribute As System.Xml.Linq.XAttribute = style.Attribute(AutoPilotWordMainNs + "type")
            If typeAttribute IsNot Nothing AndAlso
               Not System.String.Equals(typeAttribute.Value, "paragraph", System.StringComparison.OrdinalIgnoreCase) AndAlso
               Not System.String.Equals(typeAttribute.Value, "table", System.StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            Dim styleIdAttribute As System.Xml.Linq.XAttribute = style.Attribute(AutoPilotWordMainNs + "styleId")
            If styleIdAttribute Is Nothing OrElse System.String.IsNullOrWhiteSpace(styleIdAttribute.Value) Then Continue For
            Dim styleId As System.String = styleIdAttribute.Value
            If Not styleMap.ContainsKey(styleId) Then styleMap(styleId) = styleId

            Dim nameElement As System.Xml.Linq.XElement = style.Element(AutoPilotWordMainNs + "name")
            If nameElement Is Nothing Then Continue For
            Dim nameAttribute As System.Xml.Linq.XAttribute = nameElement.Attribute(AutoPilotWordMainNs + "val")
            If nameAttribute Is Nothing OrElse System.String.IsNullOrWhiteSpace(nameAttribute.Value) Then Continue For
            styleMap(nameAttribute.Value) = styleId
        Next
        Return True
    End Function

    Private Shared Function BuildAutoPilotWordOpenXmlStyleLeftIndentMap(
            archive As System.IO.Compression.ZipArchive,
            ByRef indentByStyleId As System.Collections.Generic.Dictionary(Of System.String, System.Int32),
            ByRef validationError As System.String) As System.Boolean

        ' Keep the recursive resolver on a local dictionary. Capturing the ByRef output
        ' parameter inside the lambda is rejected by the VB compiler (BC36639).
        Dim resolvedIndentByStyleId As New System.Collections.Generic.Dictionary(Of System.String, System.Int32)(System.StringComparer.OrdinalIgnoreCase)
        indentByStyleId = resolvedIndentByStyleId
        validationError = System.String.Empty
        Dim stylesXml As System.Xml.Linq.XDocument = LoadAutoPilotWordOpenXmlEntry(archive, "word/styles.xml")
        If stylesXml Is Nothing OrElse stylesXml.Root Is Nothing Then Return True
        Dim numberingXml As System.Xml.Linq.XDocument = LoadAutoPilotWordOpenXmlEntry(archive, "word/numbering.xml")

        Dim stylesById As New System.Collections.Generic.Dictionary(Of System.String, System.Xml.Linq.XElement)(System.StringComparer.OrdinalIgnoreCase)
        For Each style As System.Xml.Linq.XElement In stylesXml.Descendants(AutoPilotWordMainNs + "style")
            Dim typeAttribute As System.Xml.Linq.XAttribute = style.Attribute(AutoPilotWordMainNs + "type")
            If typeAttribute Is Nothing OrElse Not System.String.Equals(typeAttribute.Value, "paragraph", System.StringComparison.OrdinalIgnoreCase) Then Continue For
            Dim idAttribute As System.Xml.Linq.XAttribute = style.Attribute(AutoPilotWordMainNs + "styleId")
            If idAttribute Is Nothing OrElse System.String.IsNullOrWhiteSpace(idAttribute.Value) Then Continue For
            stylesById(idAttribute.Value) = style
        Next

        Dim resolving As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        Dim resolveIndent As System.Func(Of System.String, System.Int32) = Nothing
        resolveIndent = Function(styleId As System.String) As System.Int32
                            If System.String.IsNullOrWhiteSpace(styleId) Then Return 0
                            If resolvedIndentByStyleId.ContainsKey(styleId) Then Return resolvedIndentByStyleId(styleId)
                            If resolving.Contains(styleId) Then Return 0
                            resolving.Add(styleId)
                            Try
                                If Not stylesById.ContainsKey(styleId) Then Return 0
                                Dim style As System.Xml.Linq.XElement = stylesById(styleId)
                                Dim pPr As System.Xml.Linq.XElement = style.Element(AutoPilotWordMainNs + "pPr")
                                If pPr IsNot Nothing Then
                                    Dim ind As System.Xml.Linq.XElement = pPr.Element(AutoPilotWordMainNs + "ind")
                                    If ind IsNot Nothing Then
                                        Dim leftAttribute As System.Xml.Linq.XAttribute = ind.Attribute(AutoPilotWordMainNs + "left")
                                        Dim parsedLeft As System.Int32 = 0
                                        If leftAttribute IsNot Nothing AndAlso System.Int32.TryParse(leftAttribute.Value, parsedLeft) Then
                                            resolvedIndentByStyleId(styleId) = parsedLeft
                                            Return parsedLeft
                                        End If
                                    End If

                                    Dim numPr As System.Xml.Linq.XElement = pPr.Element(AutoPilotWordMainNs + "numPr")
                                    If numPr IsNot Nothing AndAlso numberingXml IsNot Nothing Then
                                        Dim numIdElement As System.Xml.Linq.XElement = numPr.Element(AutoPilotWordMainNs + "numId")
                                        Dim numIdAttribute As System.Xml.Linq.XAttribute = If(numIdElement Is Nothing, Nothing, numIdElement.Attribute(AutoPilotWordMainNs + "val"))
                                        If numIdAttribute IsNot Nothing Then
                                            Dim level As System.Int32 = 0
                                            Dim levelElement As System.Xml.Linq.XElement = numPr.Element(AutoPilotWordMainNs + "ilvl")
                                            Dim levelAttribute As System.Xml.Linq.XAttribute = If(levelElement Is Nothing, Nothing, levelElement.Attribute(AutoPilotWordMainNs + "val"))
                                            If levelAttribute IsNot Nothing Then System.Int32.TryParse(levelAttribute.Value, level)
                                            Dim num As System.Xml.Linq.XElement = numberingXml.Descendants(AutoPilotWordMainNs + "num").FirstOrDefault(
                                                Function(candidate As System.Xml.Linq.XElement) System.String.Equals(If(candidate.Attribute(AutoPilotWordMainNs + "numId"), New System.Xml.Linq.XAttribute("x", "")).Value, numIdAttribute.Value, System.StringComparison.OrdinalIgnoreCase))
                                            If num IsNot Nothing Then
                                                Dim abstractIdElement As System.Xml.Linq.XElement = num.Element(AutoPilotWordMainNs + "abstractNumId")
                                                Dim abstractIdAttribute As System.Xml.Linq.XAttribute = If(abstractIdElement Is Nothing, Nothing, abstractIdElement.Attribute(AutoPilotWordMainNs + "val"))
                                                If abstractIdAttribute IsNot Nothing Then
                                                    Dim abstractNum As System.Xml.Linq.XElement = numberingXml.Descendants(AutoPilotWordMainNs + "abstractNum").FirstOrDefault(
                                                        Function(candidate As System.Xml.Linq.XElement) System.String.Equals(If(candidate.Attribute(AutoPilotWordMainNs + "abstractNumId"), New System.Xml.Linq.XAttribute("x", "")).Value, abstractIdAttribute.Value, System.StringComparison.OrdinalIgnoreCase))
                                                    If abstractNum IsNot Nothing Then
                                                        Dim lvl As System.Xml.Linq.XElement = abstractNum.Elements(AutoPilotWordMainNs + "lvl").FirstOrDefault(
                                                            Function(candidate As System.Xml.Linq.XElement) System.String.Equals(If(candidate.Attribute(AutoPilotWordMainNs + "ilvl"), New System.Xml.Linq.XAttribute("x", "")).Value, level.ToString(System.Globalization.CultureInfo.InvariantCulture), System.StringComparison.OrdinalIgnoreCase))
                                                        Dim numInd As System.Xml.Linq.XElement = If(lvl Is Nothing OrElse lvl.Element(AutoPilotWordMainNs + "pPr") Is Nothing, Nothing, lvl.Element(AutoPilotWordMainNs + "pPr").Element(AutoPilotWordMainNs + "ind"))
                                                        Dim numLeftAttribute As System.Xml.Linq.XAttribute = If(numInd Is Nothing, Nothing, numInd.Attribute(AutoPilotWordMainNs + "left"))
                                                        Dim parsedNumLeft As System.Int32 = 0
                                                        If numLeftAttribute IsNot Nothing AndAlso System.Int32.TryParse(numLeftAttribute.Value, parsedNumLeft) Then
                                                            resolvedIndentByStyleId(styleId) = parsedNumLeft
                                                            Return parsedNumLeft
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If

                                Dim basedOn As System.Xml.Linq.XElement = style.Element(AutoPilotWordMainNs + "basedOn")
                                Dim basedOnAttribute As System.Xml.Linq.XAttribute = If(basedOn Is Nothing, Nothing, basedOn.Attribute(AutoPilotWordMainNs + "val"))
                                Dim inherited As System.Int32 = If(basedOnAttribute Is Nothing, 0, resolveIndent(basedOnAttribute.Value))
                                resolvedIndentByStyleId(styleId) = inherited
                                Return inherited
                            Finally
                                resolving.Remove(styleId)
                            End Try
                        End Function

        For Each styleId As System.String In stylesById.Keys.ToList()
            resolveIndent(styleId)
        Next
        Return True
    End Function

    Private Shared Function GetAutoPilotWordRenderedParagraphLeftIndent(
            output As System.Collections.Generic.IList(Of System.Xml.Linq.XElement),
            indentByStyleId As System.Collections.Generic.IDictionary(Of System.String, System.Int32)) As System.Int32
        If output Is Nothing Then Return 0
        For index As System.Int32 = output.Count - 1 To 0 Step -1
            Dim element As System.Xml.Linq.XElement = output(index)
            If element Is Nothing OrElse element.Name <> AutoPilotWordMainNs + "p" Then Continue For
            Dim pPr As System.Xml.Linq.XElement = element.Element(AutoPilotWordMainNs + "pPr")
            If pPr IsNot Nothing Then
                Dim ind As System.Xml.Linq.XElement = pPr.Element(AutoPilotWordMainNs + "ind")
                Dim leftAttribute As System.Xml.Linq.XAttribute = If(ind Is Nothing, Nothing, ind.Attribute(AutoPilotWordMainNs + "left"))
                Dim directLeft As System.Int32 = 0
                If leftAttribute IsNot Nothing AndAlso System.Int32.TryParse(leftAttribute.Value, directLeft) Then Return System.Math.Max(0, directLeft)
                Dim pStyle As System.Xml.Linq.XElement = pPr.Element(AutoPilotWordMainNs + "pStyle")
                Dim styleAttribute As System.Xml.Linq.XAttribute = If(pStyle Is Nothing, Nothing, pStyle.Attribute(AutoPilotWordMainNs + "val"))
                If styleAttribute IsNot Nothing AndAlso indentByStyleId IsNot Nothing AndAlso indentByStyleId.ContainsKey(styleAttribute.Value) Then
                    Return System.Math.Max(0, indentByStyleId(styleAttribute.Value))
                End If
            End If
            Return 0
        Next
        Return 0
    End Function

    Private Shared Function CreateAutoPilotWordOpenXmlRun(
            value As System.String,
            bold As System.Boolean,
            italic As System.Boolean,
            strike As System.Boolean,
            underline As System.Boolean) As System.Xml.Linq.XElement

        Dim run As New System.Xml.Linq.XElement(AutoPilotWordMainNs + "r")
        If bold OrElse italic OrElse strike OrElse underline Then
            Dim runProperties As New System.Xml.Linq.XElement(AutoPilotWordMainNs + "rPr")
            If bold Then runProperties.Add(New System.Xml.Linq.XElement(AutoPilotWordMainNs + "b"))
            If italic Then runProperties.Add(New System.Xml.Linq.XElement(AutoPilotWordMainNs + "i"))
            If strike Then runProperties.Add(New System.Xml.Linq.XElement(AutoPilotWordMainNs + "strike"))
            If underline Then runProperties.Add(New System.Xml.Linq.XElement(AutoPilotWordMainNs + "u", New System.Xml.Linq.XAttribute(AutoPilotWordMainNs + "val", "single")))
            run.Add(runProperties)
        End If

        Dim text As New System.Xml.Linq.XElement(AutoPilotWordMainNs + "t", If(value, System.String.Empty))
        If Not System.String.IsNullOrEmpty(value) AndAlso
           (System.Char.IsWhiteSpace(value(0)) OrElse System.Char.IsWhiteSpace(value(value.Length - 1))) Then
            text.SetAttributeValue(AutoPilotWordXmlNs + "space", "preserve")
        End If
        run.Add(text)
        Return run
    End Function

    Private Shared Sub AppendAutoPilotWordOpenXmlInline(
            target As System.Xml.Linq.XElement,
            node As HtmlAgilityPack.HtmlNode,
            bold As System.Boolean,
            italic As System.Boolean,
            strike As System.Boolean,
            underline As System.Boolean)

        If target Is Nothing OrElse node Is Nothing Then Return
        If node.NodeType = HtmlAgilityPack.HtmlNodeType.Text Then
            Dim value As System.String = HtmlAgilityPack.HtmlEntity.DeEntitize(node.InnerText)
            If value <> System.String.Empty Then target.Add(CreateAutoPilotWordOpenXmlRun(value, bold, italic, strike, underline))
            Return
        End If
        If node.NodeType <> HtmlAgilityPack.HtmlNodeType.Element Then Return

        Dim name As System.String = node.Name.ToLowerInvariant()
        If name = "ul" OrElse name = "ol" OrElse name = "table" Then Return
        If name = "br" Then
            target.Add(New System.Xml.Linq.XElement(AutoPilotWordMainNs + "r", New System.Xml.Linq.XElement(AutoPilotWordMainNs + "br")))
            Return
        End If

        Dim nextBold As System.Boolean = bold OrElse name = "strong" OrElse name = "b"
        Dim nextItalic As System.Boolean = italic OrElse name = "em" OrElse name = "i"
        Dim nextStrike As System.Boolean = strike OrElse name = "del" OrElse name = "s"
        Dim nextUnderline As System.Boolean = underline OrElse name = "a" OrElse name = "u"
        For Each child As HtmlAgilityPack.HtmlNode In node.ChildNodes
            AppendAutoPilotWordOpenXmlInline(target, child, nextBold, nextItalic, nextStrike, nextUnderline)
        Next
    End Sub

    Private Shared Sub NormalizeAutoPilotWordOpenXmlInlineBoundaries(paragraph As System.Xml.Linq.XElement)
        If paragraph Is Nothing Then Return
        Dim textNodes As System.Collections.Generic.List(Of System.Xml.Linq.XElement) = paragraph.Descendants(AutoPilotWordMainNs + "t").ToList()
        If textNodes.Count = 0 Then Return

        Dim first As System.Xml.Linq.XElement = textNodes.FirstOrDefault(Function(item As System.Xml.Linq.XElement) Not System.String.IsNullOrEmpty(item.Value))
        If first IsNot Nothing Then
            first.Value = first.Value.TrimStart()
            If first.Value.Length > 0 AndAlso System.Char.IsWhiteSpace(first.Value(first.Value.Length - 1)) Then
                first.SetAttributeValue(AutoPilotWordXmlNs + "space", "preserve")
            Else
                first.SetAttributeValue(AutoPilotWordXmlNs + "space", Nothing)
            End If
        End If

        Dim last As System.Xml.Linq.XElement = textNodes.LastOrDefault(Function(item As System.Xml.Linq.XElement) Not System.String.IsNullOrEmpty(item.Value))
        If last IsNot Nothing Then
            last.Value = last.Value.TrimEnd()
            If last.Value.Length > 0 AndAlso System.Char.IsWhiteSpace(last.Value(0)) Then
                last.SetAttributeValue(AutoPilotWordXmlNs + "space", "preserve")
            Else
                last.SetAttributeValue(AutoPilotWordXmlNs + "space", Nothing)
            End If
        End If

        Dim firstMeaningfulIndex As System.Int32 = textNodes.FindIndex(Function(item As System.Xml.Linq.XElement) Not System.String.IsNullOrWhiteSpace(item.Value))
        Dim lastMeaningfulIndex As System.Int32 = textNodes.FindLastIndex(Function(item As System.Xml.Linq.XElement) Not System.String.IsNullOrWhiteSpace(item.Value))
        For index As System.Int32 = 0 To textNodes.Count - 1
            If index >= firstMeaningfulIndex AndAlso index <= lastMeaningfulIndex Then Continue For
            Dim boundaryWhitespace As System.Xml.Linq.XElement = textNodes(index)
            If Not System.String.IsNullOrWhiteSpace(boundaryWhitespace.Value) Then Continue For
            Dim run As System.Xml.Linq.XElement = boundaryWhitespace.Ancestors(AutoPilotWordMainNs + "r").FirstOrDefault()
            If run IsNot Nothing AndAlso run.Parent Is paragraph AndAlso run.Elements().All(Function(child As System.Xml.Linq.XElement) child.Name = AutoPilotWordMainNs + "rPr" OrElse child.Name = AutoPilotWordMainNs + "t") Then run.Remove()
        Next
    End Sub

    Private Shared Sub RemoveAutoPilotWordHeadingPrefix(node As HtmlAgilityPack.HtmlNode)
        If node Is Nothing Then Return
        Dim visibleText As System.String = HtmlAgilityPack.HtmlEntity.DeEntitize(node.InnerText)
        Dim match As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(
            visibleText,
            "^\s*(?:(?:(?:[IVXLCDM]+|[A-Z]|\d+(?:\.\d+)*)[\.\)]|\([ivxlcdm]+\))\s+)+",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant)
        If Not match.Success OrElse match.Length <= 0 Then Return

        Dim remaining As System.Int32 = match.Length
        For Each textNode As HtmlAgilityPack.HtmlNode In node.DescendantsAndSelf().Where(
            Function(candidate As HtmlAgilityPack.HtmlNode) candidate.NodeType = HtmlAgilityPack.HtmlNodeType.Text).ToList()
            If remaining <= 0 Then Exit For
            Dim decoded As System.String = HtmlAgilityPack.HtmlEntity.DeEntitize(textNode.InnerText)
            If decoded.Length <= remaining Then
                remaining -= decoded.Length
                textNode.InnerHtml = System.String.Empty
            Else
                textNode.InnerHtml = HtmlAgilityPack.HtmlEntity.Entitize(decoded.Substring(remaining))
                remaining = 0
            End If
        Next
    End Sub

    Private Shared Function CreateAutoPilotWordOpenXmlParagraph(
            htmlNode As HtmlAgilityPack.HtmlNode,
            semantic As System.String,
            semanticStyleIds As System.Collections.Generic.IDictionary(Of System.String, System.String),
            forcePlainMarker As System.Boolean) As System.Xml.Linq.XElement

        Dim paragraph As New System.Xml.Linq.XElement(AutoPilotWordMainNs + "p")
        If Not forcePlainMarker AndAlso semanticStyleIds IsNot Nothing AndAlso
           Not System.String.IsNullOrWhiteSpace(semantic) AndAlso semanticStyleIds.ContainsKey(semantic) Then
            paragraph.Add(New System.Xml.Linq.XElement(
                AutoPilotWordMainNs + "pPr",
                New System.Xml.Linq.XElement(AutoPilotWordMainNs + "pStyle", New System.Xml.Linq.XAttribute(AutoPilotWordMainNs + "val", semanticStyleIds(semantic)))))
        End If

        If htmlNode IsNot Nothing Then
            For Each child As HtmlAgilityPack.HtmlNode In htmlNode.ChildNodes
                AppendAutoPilotWordOpenXmlInline(paragraph, child, False, False, False, False)
            Next
        End If
        NormalizeAutoPilotWordOpenXmlInlineBoundaries(paragraph)
        If Not paragraph.Elements(AutoPilotWordMainNs + "r").Any() Then
            paragraph.Add(CreateAutoPilotWordOpenXmlRun(System.String.Empty, False, False, False, False))
        End If
        Return paragraph
    End Function

    Private Shared Sub AppendAutoPilotWordOpenXmlList(
            output As System.Collections.Generic.List(Of System.Xml.Linq.XElement),
            listNode As HtmlAgilityPack.HtmlNode,
            level As System.Int32,
            semanticStyleIds As System.Collections.Generic.IDictionary(Of System.String, System.String))

        If output Is Nothing OrElse listNode Is Nothing Then Return
        Dim prefix As System.String = If(System.String.Equals(listNode.Name, "ol", System.StringComparison.OrdinalIgnoreCase), "numbered", "bullet")

        Dim itemOrdinal As System.Int32 = 0
        For Each item As HtmlAgilityPack.HtmlNode In listNode.ChildNodes.Where(
            Function(candidate As HtmlAgilityPack.HtmlNode) candidate.NodeType = HtmlAgilityPack.HtmlNodeType.Element AndAlso System.String.Equals(candidate.Name, "li", System.StringComparison.OrdinalIgnoreCase))

            itemOrdinal += 1
            Dim inlineHost As HtmlAgilityPack.HtmlNode = HtmlAgilityPack.HtmlNode.CreateNode("<span></span>")
            If System.String.Equals(prefix, "numbered", System.StringComparison.OrdinalIgnoreCase) Then
                Dim numberPrefix As HtmlAgilityPack.HtmlNode = HtmlAgilityPack.HtmlNode.CreateNode("<span></span>")
                numberPrefix.InnerHtml = HtmlAgilityPack.HtmlEntity.Entitize(itemOrdinal.ToString(System.Globalization.CultureInfo.InvariantCulture) & ". ")
                inlineHost.AppendChild(numberPrefix)
            End If
            For Each child As HtmlAgilityPack.HtmlNode In item.ChildNodes
                If child.NodeType = HtmlAgilityPack.HtmlNodeType.Element AndAlso
                   (System.String.Equals(child.Name, "ul", System.StringComparison.OrdinalIgnoreCase) OrElse System.String.Equals(child.Name, "ol", System.StringComparison.OrdinalIgnoreCase)) Then
                    Continue For
                End If

                If child.NodeType = HtmlAgilityPack.HtmlNodeType.Element AndAlso System.String.Equals(child.Name, "p", System.StringComparison.OrdinalIgnoreCase) Then
                    For Each grandChild As HtmlAgilityPack.HtmlNode In child.ChildNodes
                        inlineHost.AppendChild(grandChild.CloneNode(True))
                    Next
                Else
                    inlineHost.AppendChild(child.CloneNode(True))
                End If
            Next

            output.Add(CreateAutoPilotWordOpenXmlParagraph(
                inlineHost,
                prefix & level.ToString(System.Globalization.CultureInfo.InvariantCulture),
                semanticStyleIds,
                False))

            For Each nested As HtmlAgilityPack.HtmlNode In item.ChildNodes.Where(
                Function(candidate As HtmlAgilityPack.HtmlNode) candidate.NodeType = HtmlAgilityPack.HtmlNodeType.Element AndAlso (System.String.Equals(candidate.Name, "ul", System.StringComparison.OrdinalIgnoreCase) OrElse System.String.Equals(candidate.Name, "ol", System.StringComparison.OrdinalIgnoreCase)))
                AppendAutoPilotWordOpenXmlList(output, nested, level + 1, semanticStyleIds)
            Next
        Next
    End Sub

    Private Shared Function CreateAutoPilotWordOpenXmlTable(
            tableNode As HtmlAgilityPack.HtmlNode,
            tableStyleId As System.String,
            leftIndentTwips As System.Int32) As System.Xml.Linq.XElement

        Dim table As New System.Xml.Linq.XElement(AutoPilotWordMainNs + "tbl")
        Dim tableProperties As New System.Xml.Linq.XElement(AutoPilotWordMainNs + "tblPr")
        If Not System.String.IsNullOrWhiteSpace(tableStyleId) Then
            tableProperties.Add(New System.Xml.Linq.XElement(AutoPilotWordMainNs + "tblStyle", New System.Xml.Linq.XAttribute(AutoPilotWordMainNs + "val", tableStyleId)))
        Else
            Dim borders As New System.Xml.Linq.XElement(AutoPilotWordMainNs + "tblBorders")
            For Each edge As System.String In New System.String() {"top", "left", "bottom", "right", "insideH", "insideV"}
                borders.Add(New System.Xml.Linq.XElement(AutoPilotWordMainNs + edge,
                    New System.Xml.Linq.XAttribute(AutoPilotWordMainNs + "val", "single"),
                    New System.Xml.Linq.XAttribute(AutoPilotWordMainNs + "sz", "4"),
                    New System.Xml.Linq.XAttribute(AutoPilotWordMainNs + "space", "0"),
                    New System.Xml.Linq.XAttribute(AutoPilotWordMainNs + "color", "D9D9D9")))
            Next
            tableProperties.Add(borders)
        End If
        If leftIndentTwips > 0 Then
            tableProperties.Add(New System.Xml.Linq.XElement(
                AutoPilotWordMainNs + "tblInd",
                New System.Xml.Linq.XAttribute(AutoPilotWordMainNs + "w", leftIndentTwips.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                New System.Xml.Linq.XAttribute(AutoPilotWordMainNs + "type", "dxa")))
        End If
        tableProperties.Add(New System.Xml.Linq.XElement(
            AutoPilotWordMainNs + "tblW",
            New System.Xml.Linq.XAttribute(AutoPilotWordMainNs + "w", "0"),
            New System.Xml.Linq.XAttribute(AutoPilotWordMainNs + "type", "auto")))
        table.Add(tableProperties)

        For Each rowNode As HtmlAgilityPack.HtmlNode In tableNode.Descendants("tr")
            Dim row As New System.Xml.Linq.XElement(AutoPilotWordMainNs + "tr")
            For Each cellNode As HtmlAgilityPack.HtmlNode In rowNode.ChildNodes.Where(
                Function(candidate As HtmlAgilityPack.HtmlNode) candidate.NodeType = HtmlAgilityPack.HtmlNodeType.Element AndAlso (System.String.Equals(candidate.Name, "td", System.StringComparison.OrdinalIgnoreCase) OrElse System.String.Equals(candidate.Name, "th", System.StringComparison.OrdinalIgnoreCase)))
                Dim cell As New System.Xml.Linq.XElement(AutoPilotWordMainNs + "tc")
                cell.Add(New System.Xml.Linq.XElement(AutoPilotWordMainNs + "tcPr"))
                Dim paragraph As New System.Xml.Linq.XElement(AutoPilotWordMainNs + "p")
                For Each child As HtmlAgilityPack.HtmlNode In cellNode.ChildNodes
                    AppendAutoPilotWordOpenXmlInline(paragraph, child, System.String.Equals(cellNode.Name, "th", System.StringComparison.OrdinalIgnoreCase), False, False, False)
                Next
                If Not paragraph.Elements(AutoPilotWordMainNs + "r").Any() Then paragraph.Add(CreateAutoPilotWordOpenXmlRun(System.String.Empty, False, False, False, False))
                cell.Add(paragraph)
                row.Add(cell)
            Next
            If row.Elements(AutoPilotWordMainNs + "tc").Any() Then table.Add(row)
        Next
        Return table
    End Function

    Private Shared Function ValidateAutoPilotWordOpenXmlMarkdownSemantics(
            htmlDoc As HtmlAgilityPack.HtmlDocument,
            semanticStyleIds As System.Collections.Generic.IDictionary(Of System.String, System.String),
            ByRef validationError As System.String) As System.Boolean

        validationError = System.String.Empty
        If htmlDoc Is Nothing OrElse semanticStyleIds Is Nothing OrElse semanticStyleIds.Count = 0 Then Return True

        For Each node As HtmlAgilityPack.HtmlNode In htmlDoc.DocumentNode.Descendants()
            If node.NodeType <> HtmlAgilityPack.HtmlNodeType.Element Then Continue For
            Dim name As System.String = node.Name.ToLowerInvariant()
            If System.Text.RegularExpressions.Regex.IsMatch(name, "^h[1-6]$") Then
                Dim semantic As System.String = "heading" & name.Substring(1)
                If Not semanticStyleIds.ContainsKey(semantic) Then
                    validationError = "The selected Word style policy does not permit Markdown " & semantic & ". Use only the heading levels declared by the design."
                    Return False
                End If
            ElseIf name = "ul" OrElse name = "ol" Then
                Dim depth As System.Int32 = 1
                Dim ancestor As HtmlAgilityPack.HtmlNode = node.ParentNode
                Do While ancestor IsNot Nothing
                    If ancestor.NodeType = HtmlAgilityPack.HtmlNodeType.Element AndAlso
                       (System.String.Equals(ancestor.Name, "ul", System.StringComparison.OrdinalIgnoreCase) OrElse System.String.Equals(ancestor.Name, "ol", System.StringComparison.OrdinalIgnoreCase)) Then
                        depth += 1
                    End If
                    ancestor = ancestor.ParentNode
                Loop
                Dim semantic As System.String = If(name = "ol", "numbered", "bullet") & depth.ToString(System.Globalization.CultureInfo.InvariantCulture)
                If Not semanticStyleIds.ContainsKey(semantic) Then
                    validationError = "The selected Word style policy does not permit Markdown " & semantic & ". Use only the list levels declared by the design."
                    Return False
                End If
            End If
        Next
        Return True
    End Function

    Private Shared Function RenderAutoPilotWordMarkdownOpenXml(
            markdownContent As System.String,
            semanticStyleIds As System.Collections.Generic.IDictionary(Of System.String, System.String),
            tableStyleId As System.String,
            headingNumberingMode As System.String,
            ByRef renderingError As System.String,
            Optional styleLeftIndentById As System.Collections.Generic.IDictionary(Of System.String, System.Int32) = Nothing) As System.Collections.Generic.List(Of System.Xml.Linq.XElement)

        renderingError = System.String.Empty
        Dim output As New System.Collections.Generic.List(Of System.Xml.Linq.XElement)()

        Try
            Dim pipeline As Markdig.MarkdownPipeline = SharedLibrary.SharedLibrary.SharedMethods.CreateMarkdownHtmlPipeline(useSoftlineBreakAsHardlineBreak:=True)
            Dim html As System.String = Markdig.Markdown.ToHtml(SharedLibrary.SharedLibrary.SharedMethods.NormalizeMarkdownForHtmlDisplay(If(markdownContent, System.String.Empty)), pipeline)
            Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()
            htmlDoc.LoadHtml(html)
            SharedLibrary.SharedLibrary.SharedMethods.NormalizeMarkdigHtmlBlockBoundaryWhitespace(htmlDoc)
            If Not ValidateAutoPilotWordOpenXmlMarkdownSemantics(htmlDoc, semanticStyleIds, renderingError) Then Return Nothing

            For Each node As HtmlAgilityPack.HtmlNode In htmlDoc.DocumentNode.ChildNodes
                If node.NodeType <> HtmlAgilityPack.HtmlNodeType.Element Then Continue For
                Dim name As System.String = node.Name.ToLowerInvariant()
                If System.Text.RegularExpressions.Regex.IsMatch(name, "^h[1-6]$") Then
                    Dim level As System.Int32 = System.Int32.Parse(name.Substring(1), System.Globalization.CultureInfo.InvariantCulture)
                    If System.String.Equals(headingNumberingMode, "native", System.StringComparison.OrdinalIgnoreCase) Then RemoveAutoPilotWordHeadingPrefix(node)
                    output.Add(CreateAutoPilotWordOpenXmlParagraph(node, "heading" & level.ToString(System.Globalization.CultureInfo.InvariantCulture), semanticStyleIds, False))
                ElseIf name = "p" Then
                    Dim visibleText As System.String = HtmlAgilityPack.HtmlEntity.DeEntitize(node.InnerText).Trim()
                    Dim isVisualMarker As System.Boolean = System.Text.RegularExpressions.Regex.IsMatch(visibleText, "^\[\[visual:[^\]]+\]\]$", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)
                    output.Add(CreateAutoPilotWordOpenXmlParagraph(node, "paragraph", semanticStyleIds, isVisualMarker))
                ElseIf name = "ul" OrElse name = "ol" Then
                    AppendAutoPilotWordOpenXmlList(output, node, 1, semanticStyleIds)
                ElseIf name = "table" Then
                    Dim tableIndent As System.Int32 = GetAutoPilotWordRenderedParagraphLeftIndent(output, styleLeftIndentById)
                    output.Add(CreateAutoPilotWordOpenXmlTable(node, tableStyleId, tableIndent))
                ElseIf name = "blockquote" Then
                    For Each child As HtmlAgilityPack.HtmlNode In node.ChildNodes.Where(Function(candidate As HtmlAgilityPack.HtmlNode) candidate.NodeType = HtmlAgilityPack.HtmlNodeType.Element)
                        output.Add(CreateAutoPilotWordOpenXmlParagraph(child, "paragraph", semanticStyleIds, False))
                    Next
                ElseIf name = "pre" Then
                    output.Add(CreateAutoPilotWordOpenXmlParagraph(node, "paragraph", semanticStyleIds, False))
                End If
            Next

            If output.Count = 0 Then output.Add(CreateAutoPilotWordOpenXmlParagraph(Nothing, "paragraph", semanticStyleIds, False))
            Return output
        Catch ex As System.Exception
            renderingError = "OOXML Markdown rendering failed: " & ex.Message
            Return Nothing
        End Try
    End Function

    Private Shared Sub NormalizeAutoPilotWordOpenXmlFieldUpdateState(
            archive As System.IO.Compression.ZipArchive,
            storyXml As System.Collections.Generic.IDictionary(Of System.String, System.Xml.Linq.XDocument))

        ' OOXML-only creation must never force Word to update fields on open.
        ' Forcing updateFields/dirty causes Word to display link/field update prompts and can
        ' make the document appear blank behind a modal prompt. Preserve the cached field
        ' result already stored in the template and leave any later refresh to an explicit
        ' Word operation requested by the caller.
        If storyXml IsNot Nothing Then
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Xml.Linq.XDocument) In storyXml
                If pair.Value Is Nothing Then Continue For
                For Each fieldChar As System.Xml.Linq.XElement In pair.Value.Descendants(AutoPilotWordMainNs + "fldChar")
                    fieldChar.SetAttributeValue(AutoPilotWordMainNs + "dirty", Nothing)
                Next
            Next
        End If

        Dim settingsXml As System.Xml.Linq.XDocument = LoadAutoPilotWordOpenXmlEntry(archive, "word/settings.xml")
        If settingsXml Is Nothing OrElse settingsXml.Root Is Nothing Then Return
        For Each updateFields As System.Xml.Linq.XElement In settingsXml.Root.Elements(AutoPilotWordMainNs + "updateFields").ToList()
            updateFields.Remove()
        Next
        SaveAutoPilotWordOpenXmlEntry(archive, "word/settings.xml", settingsXml)
    End Sub

    Private Shared Function NormalizeAutoPilotWordOpenXmlValidationText(value As System.String) As System.String
        Dim normalized As System.String = If(value, System.String.Empty).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        normalized = System.Text.RegularExpressions.Regex.Replace(normalized, "\s+", " ").Trim()
        Return normalized
    End Function

    Private Shared Function ValidateAutoPilotStructuredWordOpenXmlOutput(
            archive As System.IO.Compression.ZipArchive,
            contract As SharedLibrary.Agents.WordTemplateBindingContract,
            templateFields As Newtonsoft.Json.Linq.JObject,
            markdownContent As System.String,
            ByRef validationError As System.String) As System.Boolean

        validationError = System.String.Empty
        If archive Is Nothing Then
            validationError = "The generated Word package could not be reopened for validation."
            Return False
        End If

        Dim mainXml As System.Xml.Linq.XDocument = LoadAutoPilotWordOpenXmlEntry(archive, "word/document.xml")
        If mainXml Is Nothing OrElse mainXml.Root Is Nothing Then
            validationError = "The generated Word package has no readable main document part."
            Return False
        End If

        Dim visibleText As System.String = System.String.Join(
            System.Environment.NewLine,
            mainXml.Descendants(AutoPilotWordMainNs + "p").Select(
                Function(paragraph As System.Xml.Linq.XElement) GetAutoPilotWordOpenXmlParagraphText(paragraph)).Where(
                Function(value As System.String) Not System.String.IsNullOrWhiteSpace(value)))

        If System.String.IsNullOrWhiteSpace(visibleText) Then
            validationError = "The generated Word document contains no visible main-document text. Output was rejected."
            Return False
        End If

        If System.Text.RegularExpressions.Regex.IsMatch(
                visibleText,
                "\[\[RI:[\p{L}\p{N}_.-]{1,64}\]\]",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant) Then
            validationError = "The generated Word document still contains one or more unresolved [[RI:...]] placeholders. Output was rejected."
            Return False
        End If

        If contract IsNot Nothing Then
            For Each slot As SharedLibrary.Agents.WordTemplateSlotDefinition In contract.Slots
                If slot Is Nothing OrElse Not slot.Required Then Continue For
                Dim expectedValue As System.String = GetAutoPilotWordTemplateSlotValue(slot, templateFields, markdownContent)
                If System.String.IsNullOrWhiteSpace(expectedValue) Then Continue For

                If System.String.Equals(slot.ContentMode, "text", System.StringComparison.OrdinalIgnoreCase) Then
                    Dim normalizedVisible As System.String = NormalizeAutoPilotWordOpenXmlValidationText(visibleText)
                    Dim normalizedExpected As System.String = NormalizeAutoPilotWordOpenXmlValidationText(expectedValue)
                    If normalizedExpected <> "" AndAlso normalizedVisible.IndexOf(normalizedExpected, System.StringComparison.OrdinalIgnoreCase) < 0 Then
                        validationError = "Required Word template slot '" & slot.Placeholder & "' was bound but its value is not present in the generated main document. Output was rejected."
                        Return False
                    End If
                ElseIf System.String.Equals(slot.ContentMode, "markdown", System.StringComparison.OrdinalIgnoreCase) Then
                    Dim plainProbe As System.String = System.Text.RegularExpressions.Regex.Replace(expectedValue, "[#*_`>\[\]()~-]", System.String.Empty)
                    plainProbe = System.Text.RegularExpressions.Regex.Replace(plainProbe, "\s+", " ").Trim()
                    If plainProbe.Length > 40 Then plainProbe = plainProbe.Substring(0, 40).Trim()
                    If plainProbe.Length >= 12 AndAlso visibleText.IndexOf(plainProbe, System.StringComparison.OrdinalIgnoreCase) < 0 Then
                        ' Markdown syntax can change visible spacing, so this is only a bounded
                        ' sanity probe. Structural validation above remains authoritative.
                        Dim paragraphCount As System.Int32 = mainXml.Descendants(AutoPilotWordMainNs + "p").Count(
                            Function(paragraph As System.Xml.Linq.XElement) Not System.String.IsNullOrWhiteSpace(GetAutoPilotWordOpenXmlParagraphText(paragraph)))
                        If paragraphCount < 2 Then
                            validationError = "The generated Word document did not retain the required Markdown body content. Output was rejected."
                            Return False
                        End If
                    End If
                End If
            Next
        End If

        Dim settingsXml As System.Xml.Linq.XDocument = LoadAutoPilotWordOpenXmlEntry(archive, "word/settings.xml")
        If settingsXml IsNot Nothing AndAlso settingsXml.Root IsNot Nothing AndAlso settingsXml.Root.Elements(AutoPilotWordMainNs + "updateFields").Any() Then
            validationError = "The generated Word package still requests automatic field updates on open. Output was rejected."
            Return False
        End If

        Return True
    End Function

    Private Shared Sub WriteAutoPilotGenericWordTextEntry(archive As System.IO.Compression.ZipArchive,
                                                            entryName As System.String,
                                                            content As System.String)
        Dim entry As System.IO.Compression.ZipArchiveEntry = archive.CreateEntry(entryName, System.IO.Compression.CompressionLevel.Optimal)
        Using writer As New System.IO.StreamWriter(entry.Open(), New System.Text.UTF8Encoding(False))
            writer.Write(If(content, System.String.Empty))
        End Using
    End Sub

    Private Shared Function NormalizeAutoPilotWordHexColor(value As System.String, fallback As System.String) As System.String
        Dim raw As System.String = If(value, System.String.Empty).Trim().TrimStart("#"c)
        If System.Text.RegularExpressions.Regex.IsMatch(raw, "^[0-9A-Fa-f]{6}$", System.Text.RegularExpressions.RegexOptions.CultureInvariant) Then Return raw.ToUpperInvariant()
        Dim fb As System.String = If(fallback, "000000").Trim().TrimStart("#"c)
        If System.Text.RegularExpressions.Regex.IsMatch(fb, "^[0-9A-Fa-f]{6}$", System.Text.RegularExpressions.RegexOptions.CultureInvariant) Then Return fb.ToUpperInvariant()
        Return "000000"
    End Function

    Private Shared Function BuildAutoPilotGenericWordStyles(fontName As System.String,
                                                             baseFontSize As System.Double,
                                                             accentHex As System.String,
                                                             textHex As System.String) As System.Xml.Linq.XDocument
        Dim w As System.Xml.Linq.XNamespace = AutoPilotWordMainNs
        Dim effectiveFont As System.String = If(System.String.IsNullOrWhiteSpace(fontName), "Aptos", fontName.Trim())
        Dim halfPoints As System.Int32 = CInt(System.Math.Round(System.Math.Max(8.0R, System.Math.Min(18.0R, baseFontSize)) * 2.0R))
        Dim accent As System.String = NormalizeAutoPilotWordHexColor(accentHex, "17365D")
        Dim textColor As System.String = NormalizeAutoPilotWordHexColor(textHex, "202124")

        Dim styles As New System.Xml.Linq.XElement(w + "styles",
            New System.Xml.Linq.XElement(w + "docDefaults",
                New System.Xml.Linq.XElement(w + "rPrDefault",
                    New System.Xml.Linq.XElement(w + "rPr",
                        New System.Xml.Linq.XElement(w + "rFonts", New System.Xml.Linq.XAttribute(w + "ascii", effectiveFont), New System.Xml.Linq.XAttribute(w + "hAnsi", effectiveFont), New System.Xml.Linq.XAttribute(w + "eastAsia", effectiveFont)),
                        New System.Xml.Linq.XElement(w + "color", New System.Xml.Linq.XAttribute(w + "val", textColor)),
                        New System.Xml.Linq.XElement(w + "sz", New System.Xml.Linq.XAttribute(w + "val", halfPoints.ToString(System.Globalization.CultureInfo.InvariantCulture))),
                        New System.Xml.Linq.XElement(w + "szCs", New System.Xml.Linq.XAttribute(w + "val", halfPoints.ToString(System.Globalization.CultureInfo.InvariantCulture))))),
                New System.Xml.Linq.XElement(w + "pPrDefault",
                    New System.Xml.Linq.XElement(w + "pPr",
                        New System.Xml.Linq.XElement(w + "spacing", New System.Xml.Linq.XAttribute(w + "after", "120"), New System.Xml.Linq.XAttribute(w + "line", "276"), New System.Xml.Linq.XAttribute(w + "lineRule", "auto"))))),
            New System.Xml.Linq.XElement(w + "style", New System.Xml.Linq.XAttribute(w + "type", "paragraph"), New System.Xml.Linq.XAttribute(w + "default", "1"), New System.Xml.Linq.XAttribute(w + "styleId", "Normal"),
                New System.Xml.Linq.XElement(w + "name", New System.Xml.Linq.XAttribute(w + "val", "Normal")),
                New System.Xml.Linq.XElement(w + "qFormat")))

        Dim headingSizes() As System.Int32 = {32, 28, 26, 24, 22, 20}
        For level As System.Int32 = 1 To 6
            styles.Add(New System.Xml.Linq.XElement(w + "style", New System.Xml.Linq.XAttribute(w + "type", "paragraph"), New System.Xml.Linq.XAttribute(w + "styleId", "Heading" & level.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                New System.Xml.Linq.XElement(w + "name", New System.Xml.Linq.XAttribute(w + "val", "Heading " & level.ToString(System.Globalization.CultureInfo.InvariantCulture))),
                New System.Xml.Linq.XElement(w + "basedOn", New System.Xml.Linq.XAttribute(w + "val", "Normal")),
                New System.Xml.Linq.XElement(w + "next", New System.Xml.Linq.XAttribute(w + "val", "Normal")),
                New System.Xml.Linq.XElement(w + "qFormat"),
                New System.Xml.Linq.XElement(w + "pPr",
                    New System.Xml.Linq.XElement(w + "keepNext"),
                    New System.Xml.Linq.XElement(w + "keepLines"),
                    New System.Xml.Linq.XElement(w + "spacing", New System.Xml.Linq.XAttribute(w + "before", If(level = 1, "300", "220")), New System.Xml.Linq.XAttribute(w + "after", "100")),
                    New System.Xml.Linq.XElement(w + "outlineLvl", New System.Xml.Linq.XAttribute(w + "val", (level - 1).ToString(System.Globalization.CultureInfo.InvariantCulture)))),
                New System.Xml.Linq.XElement(w + "rPr",
                    New System.Xml.Linq.XElement(w + "b"),
                    New System.Xml.Linq.XElement(w + "color", New System.Xml.Linq.XAttribute(w + "val", accent)),
                    New System.Xml.Linq.XElement(w + "sz", New System.Xml.Linq.XAttribute(w + "val", headingSizes(level - 1).ToString(System.Globalization.CultureInfo.InvariantCulture))),
                    New System.Xml.Linq.XElement(w + "szCs", New System.Xml.Linq.XAttribute(w + "val", headingSizes(level - 1).ToString(System.Globalization.CultureInfo.InvariantCulture))))))
        Next

        styles.Add(New System.Xml.Linq.XElement(w + "style", New System.Xml.Linq.XAttribute(w + "type", "paragraph"), New System.Xml.Linq.XAttribute(w + "styleId", "Title"),
            New System.Xml.Linq.XElement(w + "name", New System.Xml.Linq.XAttribute(w + "val", "Title")),
            New System.Xml.Linq.XElement(w + "basedOn", New System.Xml.Linq.XAttribute(w + "val", "Normal")),
            New System.Xml.Linq.XElement(w + "next", New System.Xml.Linq.XAttribute(w + "val", "Normal")),
            New System.Xml.Linq.XElement(w + "rPr", New System.Xml.Linq.XElement(w + "b"), New System.Xml.Linq.XElement(w + "color", New System.Xml.Linq.XAttribute(w + "val", accent)), New System.Xml.Linq.XElement(w + "sz", New System.Xml.Linq.XAttribute(w + "val", "40")))))

        For level As System.Int32 = 1 To 3
            styles.Add(New System.Xml.Linq.XElement(w + "style", New System.Xml.Linq.XAttribute(w + "type", "paragraph"), New System.Xml.Linq.XAttribute(w + "styleId", "GenericBullet" & level.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                New System.Xml.Linq.XElement(w + "name", New System.Xml.Linq.XAttribute(w + "val", "Generic Bullet " & level.ToString(System.Globalization.CultureInfo.InvariantCulture))),
                New System.Xml.Linq.XElement(w + "basedOn", New System.Xml.Linq.XAttribute(w + "val", "Normal")),
                New System.Xml.Linq.XElement(w + "pPr", New System.Xml.Linq.XElement(w + "numPr", New System.Xml.Linq.XElement(w + "ilvl", New System.Xml.Linq.XAttribute(w + "val", (level - 1).ToString(System.Globalization.CultureInfo.InvariantCulture))), New System.Xml.Linq.XElement(w + "numId", New System.Xml.Linq.XAttribute(w + "val", "1"))))))
            styles.Add(New System.Xml.Linq.XElement(w + "style", New System.Xml.Linq.XAttribute(w + "type", "paragraph"), New System.Xml.Linq.XAttribute(w + "styleId", "GenericNumbered" & level.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                New System.Xml.Linq.XElement(w + "name", New System.Xml.Linq.XAttribute(w + "val", "Generic Numbered " & level.ToString(System.Globalization.CultureInfo.InvariantCulture))),
                New System.Xml.Linq.XElement(w + "basedOn", New System.Xml.Linq.XAttribute(w + "val", "Normal")),
                New System.Xml.Linq.XElement(w + "pPr", New System.Xml.Linq.XElement(w + "numPr", New System.Xml.Linq.XElement(w + "ilvl", New System.Xml.Linq.XAttribute(w + "val", (level - 1).ToString(System.Globalization.CultureInfo.InvariantCulture))), New System.Xml.Linq.XElement(w + "numId", New System.Xml.Linq.XAttribute(w + "val", "2"))))))
        Next

        Return New System.Xml.Linq.XDocument(New System.Xml.Linq.XDeclaration("1.0", "UTF-8", "yes"), styles)
    End Function

    Private Shared Function BuildAutoPilotGenericWordNumbering() As System.Xml.Linq.XDocument
        Dim w As System.Xml.Linq.XNamespace = AutoPilotWordMainNs
        Dim numbering As New System.Xml.Linq.XElement(w + "numbering")
        Dim bullet As New System.Xml.Linq.XElement(w + "abstractNum", New System.Xml.Linq.XAttribute(w + "abstractNumId", "1"), New System.Xml.Linq.XElement(w + "multiLevelType", New System.Xml.Linq.XAttribute(w + "val", "multilevel")))
        Dim numbered As New System.Xml.Linq.XElement(w + "abstractNum", New System.Xml.Linq.XAttribute(w + "abstractNumId", "2"), New System.Xml.Linq.XElement(w + "multiLevelType", New System.Xml.Linq.XAttribute(w + "val", "multilevel")))
        For level As System.Int32 = 0 To 2
            Dim left As System.Int32 = 720 + level * 360
            bullet.Add(New System.Xml.Linq.XElement(w + "lvl", New System.Xml.Linq.XAttribute(w + "ilvl", level.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                New System.Xml.Linq.XElement(w + "start", New System.Xml.Linq.XAttribute(w + "val", "1")),
                New System.Xml.Linq.XElement(w + "numFmt", New System.Xml.Linq.XAttribute(w + "val", "bullet")),
                New System.Xml.Linq.XElement(w + "lvlText", New System.Xml.Linq.XAttribute(w + "val", If(level Mod 2 = 0, "•", "–"))),
                New System.Xml.Linq.XElement(w + "pPr", New System.Xml.Linq.XElement(w + "ind", New System.Xml.Linq.XAttribute(w + "left", left.ToString(System.Globalization.CultureInfo.InvariantCulture)), New System.Xml.Linq.XAttribute(w + "hanging", "360")))))
            numbered.Add(New System.Xml.Linq.XElement(w + "lvl", New System.Xml.Linq.XAttribute(w + "ilvl", level.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                New System.Xml.Linq.XElement(w + "start", New System.Xml.Linq.XAttribute(w + "val", "1")),
                New System.Xml.Linq.XElement(w + "numFmt", New System.Xml.Linq.XAttribute(w + "val", "decimal")),
                New System.Xml.Linq.XElement(w + "lvlText", New System.Xml.Linq.XAttribute(w + "val", "%" & (level + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) & ".")),
                New System.Xml.Linq.XElement(w + "pPr", New System.Xml.Linq.XElement(w + "ind", New System.Xml.Linq.XAttribute(w + "left", left.ToString(System.Globalization.CultureInfo.InvariantCulture)), New System.Xml.Linq.XAttribute(w + "hanging", "360")))))
        Next
        numbering.Add(bullet)
        numbering.Add(numbered)
        numbering.Add(New System.Xml.Linq.XElement(w + "num", New System.Xml.Linq.XAttribute(w + "numId", "1"), New System.Xml.Linq.XElement(w + "abstractNumId", New System.Xml.Linq.XAttribute(w + "val", "1"))))
        numbering.Add(New System.Xml.Linq.XElement(w + "num", New System.Xml.Linq.XAttribute(w + "numId", "2"), New System.Xml.Linq.XElement(w + "abstractNumId", New System.Xml.Linq.XAttribute(w + "val", "2"))))
        Return New System.Xml.Linq.XDocument(New System.Xml.Linq.XDeclaration("1.0", "UTF-8", "yes"), numbering)
    End Function

    Private Shared Function TryCreateAutoPilotGenericWordDocumentOpenXml(outputPath As System.String,
                                                                         markdownContent As System.String,
                                                                         arguments As System.Collections.Generic.Dictionary(Of System.String, System.Object),
                                                                         ByRef creationSummary As System.String,
                                                                         ByRef creationError As System.String) As System.Boolean
        creationSummary = System.String.Empty
        creationError = System.String.Empty
        Try
            If System.IO.File.Exists(outputPath) Then System.IO.File.Delete(outputPath)
            Dim fontName As System.String = GetArgString(arguments, "base_font_name")
            If System.String.IsNullOrWhiteSpace(fontName) Then fontName = "Aptos"
            Dim fontSize As System.Double = 11.0R
            Dim rawFontSize As System.String = GetArgString(arguments, "base_font_size")
            If Not System.String.IsNullOrWhiteSpace(rawFontSize) Then System.Double.TryParse(rawFontSize, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, fontSize)

            Dim semanticStyleIds As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase) From {
                {"paragraph", "Normal"}, {"heading1", "Heading1"}, {"heading2", "Heading2"}, {"heading3", "Heading3"}, {"heading4", "Heading4"}, {"heading5", "Heading5"}, {"heading6", "Heading6"},
                {"bullet1", "GenericBullet1"}, {"bullet2", "GenericBullet2"}, {"bullet3", "GenericBullet3"}, {"numbered1", "GenericNumbered1"}, {"numbered2", "GenericNumbered2"}, {"numbered3", "GenericNumbered3"}
            }
            Dim genericIndentByStyleId As New System.Collections.Generic.Dictionary(Of System.String, System.Int32)(System.StringComparer.OrdinalIgnoreCase) From {
                {"Normal", 0}, {"Heading1", 0}, {"Heading2", 0}, {"Heading3", 0}, {"Heading4", 0}, {"Heading5", 0}, {"Heading6", 0},
                {"GenericBullet1", 720}, {"GenericBullet2", 1080}, {"GenericBullet3", 1440}, {"GenericNumbered1", 720}, {"GenericNumbered2", 1080}, {"GenericNumbered3", 1440}
            }
            Dim renderingError As System.String = System.String.Empty
            Dim rendered As System.Collections.Generic.List(Of System.Xml.Linq.XElement) = RenderAutoPilotWordMarkdownOpenXml(markdownContent, semanticStyleIds, System.String.Empty, "preserve", renderingError, genericIndentByStyleId)
            If rendered Is Nothing Then
                creationError = renderingError
                Return False
            End If

            Dim w As System.Xml.Linq.XNamespace = AutoPilotWordMainNs
            Dim body As New System.Xml.Linq.XElement(w + "body")
            If GetArgBool(arguments, "include_cover", False) Then
                Dim coverTitle As System.String = GetArgString(arguments, "cover_title")
                If System.String.IsNullOrWhiteSpace(coverTitle) Then coverTitle = GetArgString(arguments, "document_title")
                If Not System.String.IsNullOrWhiteSpace(coverTitle) Then
                    body.Add(New System.Xml.Linq.XElement(w + "p", New System.Xml.Linq.XElement(w + "pPr", New System.Xml.Linq.XElement(w + "pStyle", New System.Xml.Linq.XAttribute(w + "val", "Title"))), CreateAutoPilotWordOpenXmlRun(coverTitle, False, False, False, False)))
                    Dim subtitle As System.String = GetArgString(arguments, "cover_subtitle")
                    If Not System.String.IsNullOrWhiteSpace(subtitle) Then body.Add(New System.Xml.Linq.XElement(w + "p", CreateAutoPilotWordOpenXmlRun(subtitle, False, False, False, False)))
                    body.Add(New System.Xml.Linq.XElement(w + "p", New System.Xml.Linq.XElement(w + "r", New System.Xml.Linq.XElement(w + "br", New System.Xml.Linq.XAttribute(w + "type", "page")))))
                End If
            End If
            body.Add(rendered)

            Dim landscape As System.Boolean = System.String.Equals(GetArgString(arguments, "page_orientation"), "landscape", System.StringComparison.OrdinalIgnoreCase)
            Dim pageWidth As System.String = If(landscape, "16838", "11906")
            Dim pageHeight As System.String = If(landscape, "11906", "16838")
            body.Add(New System.Xml.Linq.XElement(w + "sectPr",
                New System.Xml.Linq.XElement(w + "pgSz", New System.Xml.Linq.XAttribute(w + "w", pageWidth), New System.Xml.Linq.XAttribute(w + "h", pageHeight), If(landscape, New System.Xml.Linq.XAttribute(w + "orient", "landscape"), Nothing)),
                New System.Xml.Linq.XElement(w + "pgMar", New System.Xml.Linq.XAttribute(w + "top", "1440"), New System.Xml.Linq.XAttribute(w + "right", "1440"), New System.Xml.Linq.XAttribute(w + "bottom", "1440"), New System.Xml.Linq.XAttribute(w + "left", "1440"), New System.Xml.Linq.XAttribute(w + "header", "720"), New System.Xml.Linq.XAttribute(w + "footer", "720"), New System.Xml.Linq.XAttribute(w + "gutter", "0"))))
            Dim documentXml As New System.Xml.Linq.XDocument(New System.Xml.Linq.XDeclaration("1.0", "UTF-8", "yes"), New System.Xml.Linq.XElement(w + "document", body))

            Using fs As New System.IO.FileStream(outputPath, System.IO.FileMode.CreateNew, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
                Using archive As New System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Create, leaveOpen:=False)
                    WriteAutoPilotGenericWordTextEntry(archive, "[Content_Types].xml", "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/><Default Extension=""xml"" ContentType=""application/xml""/><Override PartName=""/word/document.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml""/><Override PartName=""/word/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml""/><Override PartName=""/word/numbering.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml""/><Override PartName=""/word/settings.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml""/></Types>")
                    WriteAutoPilotGenericWordTextEntry(archive, "_rels/.rels", "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""word/document.xml""/></Relationships>")
                    WriteAutoPilotGenericWordTextEntry(archive, "word/_rels/document.xml.rels", "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/><Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"" Target=""numbering.xml""/><Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"" Target=""settings.xml""/></Relationships>")
                    WriteAutoPilotGenericWordTextEntry(archive, "word/document.xml", documentXml.ToString(System.Xml.Linq.SaveOptions.DisableFormatting))
                    WriteAutoPilotGenericWordTextEntry(archive, "word/styles.xml", BuildAutoPilotGenericWordStyles(fontName, fontSize, GetArgString(arguments, "accent_color"), GetArgString(arguments, "text_color")).ToString(System.Xml.Linq.SaveOptions.DisableFormatting))
                    WriteAutoPilotGenericWordTextEntry(archive, "word/numbering.xml", BuildAutoPilotGenericWordNumbering().ToString(System.Xml.Linq.SaveOptions.DisableFormatting))
                    WriteAutoPilotGenericWordTextEntry(archive, "word/settings.xml", "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><w:settings xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""/>")
                End Using
            End Using

            If Not System.IO.File.Exists(outputPath) OrElse New System.IO.FileInfo(outputPath).Length <= 0 Then
                creationError = "Generic OOXML Word creation produced no output file."
                Return False
            End If

            Using validationStream As New System.IO.FileStream(outputPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                Using validationArchive As New System.IO.Compression.ZipArchive(validationStream, System.IO.Compression.ZipArchiveMode.Read, leaveOpen:=False)
                    If Not ValidateAutoPilotStructuredWordOpenXmlOutput(validationArchive, Nothing, Nothing, markdownContent, creationError) Then
                        Return False
                    End If
                End Using
            End Using

            creationSummary = " Created by the generic OOXML-only renderer; Word/COM was not started."
            Return True
        Catch ex As System.Exception
            creationError = "Generic OOXML Word creation failed: " & ex.Message
            Try
                If System.IO.File.Exists(outputPath) Then System.IO.File.Delete(outputPath)
            Catch cleanupEx As System.Exception
            End Try
            Return False
        End Try
    End Function

    Private Shared Function TryCreateAutoPilotStructuredWordDocumentOpenXml(
            templatePath As System.String,
            outputPath As System.String,
            contract As SharedLibrary.Agents.WordTemplateBindingContract,
            templateFields As Newtonsoft.Json.Linq.JObject,
            markdownContent As System.String,
            tableStyleName As System.String,
            ByRef bindingSummary As System.String,
            ByRef creationError As System.String) As System.Boolean

        bindingSummary = System.String.Empty
        creationError = System.String.Empty
        If contract Is Nothing OrElse Not contract.HasSlots Then
            creationError = "The OOXML structured Word renderer requires a slot-bound template contract."
            Return False
        End If
        If System.String.IsNullOrWhiteSpace(templatePath) OrElse Not System.IO.File.Exists(templatePath) Then
            creationError = "The selected structured Word template carrier was not found."
            Return False
        End If
        If templateFields Is Nothing Then templateFields = New Newtonsoft.Json.Linq.JObject()

        Try
            Dim extension As System.String = System.IO.Path.GetExtension(templatePath).ToLowerInvariant()
            Dim carrierSummary As System.String = System.String.Empty
            If extension = ".dotx" Then
                If Not TryMaterializeAutoPilotSlotBoundDotxAsDocx(templatePath, outputPath, carrierSummary, creationError) Then Return False
            ElseIf extension = ".docx" Then
                System.IO.File.Copy(templatePath, outputPath, overwrite:=False)
                carrierSummary = " Cloned the structured DOCX carrier without starting Word."
            ElseIf extension = ".dotm" Then
                creationError = "Slot-bound .dotm templates are not rendered through Word/COM. Use a macro-free .dotx/.docx carrier for deterministic OOXML creation."
                Return False
            Else
                creationError = "Unsupported structured Word template carrier: " & extension
                Return False
            End If

            Using fileStream As New System.IO.FileStream(outputPath, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
                Using archive As New System.IO.Compression.ZipArchive(fileStream, System.IO.Compression.ZipArchiveMode.Update, leaveOpen:=False)
                    Dim styleIdByName As System.Collections.Generic.Dictionary(Of System.String, System.String) = Nothing
                    If Not BuildAutoPilotWordOpenXmlStyleIdMap(archive, styleIdByName, creationError) Then Return False

                    Dim styleLeftIndentById As System.Collections.Generic.Dictionary(Of System.String, System.Int32) = Nothing
                    Dim indentMapError As System.String = System.String.Empty
                    If Not BuildAutoPilotWordOpenXmlStyleLeftIndentMap(archive, styleLeftIndentById, indentMapError) Then
                        creationError = indentMapError
                        Return False
                    End If

                    Dim semanticStyleIds As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
                    For Each definition As SharedLibrary.Agents.WordTemplateBodyStyleDefinition In contract.BodyStyles
                        If definition Is Nothing Then Continue For
                        If Not styleIdByName.ContainsKey(definition.StyleName) Then
                            creationError = "Word style '" & definition.StyleName & "' declared for semantic '" & definition.Semantic & "' was not found in the selected template."
                            Return False
                        End If
                        semanticStyleIds(definition.Semantic) = styleIdByName(definition.StyleName)
                    Next
                    For Each definition As SharedLibrary.Agents.WordTemplateBodyStyleDefinition In contract.NativeStyles
                        If definition Is Nothing Then Continue For
                        If Not styleIdByName.ContainsKey(definition.StyleName) Then
                            creationError = "Word native style '" & definition.StyleName & "' declared for policy key '" & definition.Semantic & "' was not found in the selected template."
                            Return False
                        End If
                    Next

                    Dim tableStyleId As System.String = System.String.Empty
                    If Not System.String.IsNullOrWhiteSpace(tableStyleName) AndAlso styleIdByName.ContainsKey(tableStyleName) Then tableStyleId = styleIdByName(tableStyleName)

                    Dim storyNames As System.Collections.Generic.List(Of System.String) = GetAutoPilotWordOpenXmlStoryEntryNames(archive)
                    Dim storyXml As New System.Collections.Generic.Dictionary(Of System.String, System.Xml.Linq.XDocument)(System.StringComparer.OrdinalIgnoreCase)
                    For Each storyName As System.String In storyNames
                        Dim xml As System.Xml.Linq.XDocument = LoadAutoPilotWordOpenXmlEntry(archive, storyName)
                        If xml IsNot Nothing Then storyXml(storyName) = xml
                    Next
                    If Not storyXml.ContainsKey("word/document.xml") Then
                        creationError = "The structured Word carrier has no readable main document part."
                        Return False
                    End If

                    Dim boundCount As System.Int32 = 0
                    For Each slot As SharedLibrary.Agents.WordTemplateSlotDefinition In contract.Slots
                        If slot Is Nothing OrElse Not System.String.Equals(slot.ContentMode, "text", System.StringComparison.OrdinalIgnoreCase) Then Continue For
                        Dim value As System.String = GetAutoPilotWordTemplateSlotValue(slot, templateFields, markdownContent)
                        Dim replacedForSlot As System.Int32 = 0
                        For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Xml.Linq.XDocument) In storyXml
                            For Each paragraph As System.Xml.Linq.XElement In pair.Value.Descendants(AutoPilotWordMainNs + "p")
                                replacedForSlot += ReplaceAutoPilotWordOpenXmlPlaceholderInParagraph(paragraph, slot.Placeholder, value)
                            Next
                        Next
                        If replacedForSlot = 0 Then
                            creationError = "Word template placeholder " & slot.Placeholder & " disappeared before it could be filled."
                            Return False
                        End If
                        boundCount += replacedForSlot
                    Next

                    For Each slot As SharedLibrary.Agents.WordTemplateSlotDefinition In contract.Slots
                        If slot Is Nothing OrElse Not System.String.Equals(slot.ContentMode, "markdown", System.StringComparison.OrdinalIgnoreCase) Then Continue For
                        Dim matches As System.Collections.Generic.List(Of System.Xml.Linq.XElement) = storyXml("word/document.xml").Descendants(AutoPilotWordMainNs + "p").Where(
                            Function(paragraph As System.Xml.Linq.XElement) System.String.Equals(GetAutoPilotWordOpenXmlParagraphText(paragraph).Trim(), slot.Placeholder, System.StringComparison.OrdinalIgnoreCase)).ToList()
                        If matches.Count <> 1 Then
                            creationError = "Markdown Word template placeholder " & slot.Placeholder & " must occur exactly once as the only visible content of a main-document paragraph; found " & matches.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & "."
                            Return False
                        End If

                        Dim renderingError As System.String = System.String.Empty
                        Dim rendered As System.Collections.Generic.List(Of System.Xml.Linq.XElement) = RenderAutoPilotWordMarkdownOpenXml(
                            GetAutoPilotWordTemplateSlotValue(slot, templateFields, markdownContent), semanticStyleIds, tableStyleId, contract.HeadingNumberingMode, renderingError, styleLeftIndentById)
                        If rendered Is Nothing Then
                            creationError = renderingError
                            Return False
                        End If

                        Dim markerParagraph As System.Xml.Linq.XElement = matches(0)
                        markerParagraph.AddBeforeSelf(rendered)
                        markerParagraph.Remove()
                        boundCount += 1
                    Next

                    For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Xml.Linq.XDocument) In storyXml
                        Dim unresolved As System.Boolean = pair.Value.Descendants(AutoPilotWordMainNs + "p").Any(
                            Function(paragraph As System.Xml.Linq.XElement) System.Text.RegularExpressions.Regex.IsMatch(
                                GetAutoPilotWordOpenXmlParagraphText(paragraph),
                                "\[\[RI:[\p{L}\p{N}_.-]{1,64}\]\]",
                                System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant))
                        If unresolved Then
                            creationError = "The generated Word document still contains one or more unresolved [[RI:...]] template placeholders. Output was rejected."
                            Return False
                        End If
                    Next

                    NormalizeAutoPilotWordOpenXmlFieldUpdateState(archive, storyXml)
                    For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Xml.Linq.XDocument) In storyXml
                        SaveAutoPilotWordOpenXmlEntry(archive, pair.Key, pair.Value)
                    Next

                    Dim outputValidationError As System.String = System.String.Empty
                    If Not ValidateAutoPilotStructuredWordOpenXmlOutput(archive, contract, templateFields, markdownContent, outputValidationError) Then
                        creationError = outputValidationError
                        Return False
                    End If

                    bindingSummary = carrierSummary & " Bound " & boundCount.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                     " template placeholder occurrence(s) using " & System.IO.Path.GetFileName(contract.GuidancePath) &
                                     " with the OOXML-only structured renderer; Word/COM was not started."
                End Using
            End Using

            Return System.IO.File.Exists(outputPath) AndAlso New System.IO.FileInfo(outputPath).Length > 0
        Catch ex As System.Exception
            creationError = "Structured OOXML Word creation failed: " & ex.Message
            Try
                If System.IO.File.Exists(outputPath) Then System.IO.File.Delete(outputPath)
            Catch cleanupEx As System.Exception
            End Try
            Return False
        End Try
    End Function

End Class
