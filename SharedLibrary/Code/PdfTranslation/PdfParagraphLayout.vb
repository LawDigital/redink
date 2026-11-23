Option Explicit On
Option Strict On
Option Infer On

Imports System.Linq
Imports Glyph = UglyToad.PdfPig.Content.Letter

Namespace Agents
    Partial Public NotInheritable Class PdfFileTranslation
        Private Class SourceLine
            Public Text As New System.Text.StringBuilder()
            Public Letters As New System.Collections.Generic.List(Of Glyph)()
            Public ReadOnly Property Left As System.Double
                Get
                    Return Letters.Min(Function(g) g.StartBaseLine.X)
                End Get
            End Property
            Public ReadOnly Property Right As System.Double
                Get
                    Return Letters.Max(Function(g) g.EndBaseLine.X)
                End Get
            End Property
        End Class

        Private Class TranslationChunk
            Public ParagraphNumber As System.Int32
            Public Anchor As System.Int32
            Public Epoch As System.Int32
            Public Link As System.Int32
            Public Lines As New System.Collections.Generic.List(Of SourceLine)()
            Public Painted As New System.Collections.Generic.List(Of PaintedLine)()
            Public Commands As System.Byte()
            Public NoteText As System.String
            Public Output As System.String
            Public Source As System.String
            Public Left As System.Double
            Public Right As System.Double
            Public Alignment As System.String
            Public Reflow As System.Boolean
            Public CanRetainOriginal As System.Boolean
            Public RetainOriginal As System.Boolean
            Public GrowthAllowance As System.Double
            Public OriginalEncoding As SourceEncoding
        End Class

        Private Shared Function SameFace(a As Glyph, b As Glyph) As System.Boolean
            Return a.RenderingMode = b.RenderingMode AndAlso a.FontName = b.FontName AndAlso System.Math.Abs(a.PointSize - b.PointSize) < 0.1 AndAlso a.FillColor.ToRGBValues().Equals(b.FillColor.ToRGBValues())
        End Function

        ' A justified PDF line can arrive as several extraction fragments, sometimes
        ' in content-stream order. Reconstruct physical rows within one styled chunk.
        Private Shared Sub NormalizeRows(chunk As TranslationChunk, starts As System.Collections.Generic.HashSet(Of Glyph),
                pageLetters As System.Collections.Generic.IReadOnlyList(Of Glyph), rules As System.Collections.Generic.List(Of UglyToad.PdfPig.Core.PdfRectangle))
            Dim letters = chunk.Lines.SelectMany(Function(l) l.Letters).OrderByDescending(Function(g) g.StartBaseLine.Y).ToList()
            If letters.Count = 0 Then Return
            Dim own As New System.Collections.Generic.HashSet(Of Glyph)(letters)
            Dim rows As New System.Collections.Generic.List(Of SourceLine)()
            Dim tolerance = System.Math.Min(1.0, letters(0).PointSize * 0.08)
            For Each g In letters
                If rows.Count = 0 OrElse System.Math.Abs(rows.Last().Letters(0).StartBaseLine.Y - g.StartBaseLine.Y) > tolerance Then rows.Add(New SourceLine())
                rows.Last().Letters.Add(g)
            Next
            For Each row In rows
                Dim baseline = row.Letters(0).StartBaseLine.Y
                Dim neighbors = pageLetters.Where(Function(other) Not own.Contains(other) AndAlso Not System.String.IsNullOrWhiteSpace(other.Value) AndAlso System.Math.Abs(other.StartBaseLine.Y - baseline) <= tolerance).ToList()
                row.Letters = row.Letters.OrderBy(Function(g) g.StartBaseLine.X).ToList()
                Dim prior As Glyph = Nothing
                For Each g In row.Letters
                    If prior IsNot Nothing Then
                        Dim x1 = prior.EndBaseLine.X
                        Dim x2 = g.StartBaseLine.X
                        Dim y = g.StartBaseLine.Y
                        Dim gap = x2 - x1
                        If gap > 2.2 * g.PointSize Then Return
                        If gap > 0.5 AndAlso (rules.Any(Function(r) r.Bottom <= y AndAlso r.Top >= y AndAlso r.Left > x1 AndAlso r.Right < x2) OrElse
                            neighbors.Any(Function(other) other.StartBaseLine.X >= x1 AndAlso other.StartBaseLine.X < x2)) Then Return
                        If starts.Contains(g) OrElse gap > g.PointSize * 0.2 Then row.Text.Append(" "c)
                    End If
                    row.Text.Append(g.Value)
                    prior = g
                Next
            Next
            chunk.Lines = rows
        End Sub

        ' A fill or clip can be a background panel. Only an axis-aligned four-corner
        ' path has a uniform interior that can be used up to its right edge.
        Private Shared Function AxisAlignedPanel(part As UglyToad.PdfPig.Core.PdfSubpath) As System.Boolean
            Dim points As New System.Collections.Generic.List(Of UglyToad.PdfPig.Core.PdfPoint)()
            For Each pathCommand As UglyToad.PdfPig.Core.PdfSubpath.IPathCommand In part.Commands
                If TypeOf pathCommand Is UglyToad.PdfPig.Core.PdfSubpath.Move AndAlso points.Count = 0 Then
                    points.Add(DirectCast(pathCommand, UglyToad.PdfPig.Core.PdfSubpath.Move).Location)
                ElseIf TypeOf pathCommand Is UglyToad.PdfPig.Core.PdfSubpath.Line Then
                    points.Add(DirectCast(pathCommand, UglyToad.PdfPig.Core.PdfSubpath.Line).[To])
                ElseIf Not TypeOf pathCommand Is UglyToad.PdfPig.Core.PdfSubpath.Close Then
                    Return False
                End If
            Next
            If points.Count = 5 AndAlso points(0).Equals(points(4)) Then points.RemoveAt(4)
            If points.Count <> 4 OrElse points.Distinct().Count() <> 4 Then Return False
            For i As System.Int32 = 0 To 3
                Dim other = points((i + 1) Mod 4)
                If points(i).X <> other.X AndAlso points(i).Y <> other.Y Then Return False
            Next
            Return True
        End Function

        Private Shared Function AddAnnotationObstacles(page As PdfSharp.Pdf.PdfPage, obstacles As System.Collections.Generic.List(Of UglyToad.PdfPig.Core.PdfRectangle),
                trace As System.Action(Of System.String)) As System.Boolean
            If Not page.Elements.ContainsKey("/Annots") Then Return True
            Try
                Dim entries = page.Elements.GetArray("/Annots")
                If entries IsNot Nothing Then
                    For i As System.Int32 = 0 To entries.Elements.Count - 1
                        Dim annotation = entries.Elements.GetDictionary(i)
                        If annotation Is Nothing OrElse Not annotation.Elements.ContainsKey("/Rect") Then Return False
                        Dim r = annotation.Elements.GetRectangle("/Rect")
                        If r Is Nothing OrElse r.X1 = r.X2 OrElse r.Y1 = r.Y2 Then Return False
                        obstacles.Add(New UglyToad.PdfPig.Core.PdfRectangle(System.Math.Min(r.X1, r.X2), System.Math.Min(r.Y1, r.Y2),
                            System.Math.Max(r.X1, r.X2), System.Math.Max(r.Y1, r.Y2)))
                    Next
                    Return True
                End If
            Catch ex As System.Exception When TypeOf ex Is System.ArgumentException OrElse TypeOf ex Is System.InvalidOperationException OrElse
                TypeOf ex Is System.InvalidCastException OrElse TypeOf ex Is System.FormatException OrElse TypeOf ex Is System.IndexOutOfRangeException
                trace("Annotation bounds unavailable: " & ex.Message)
            End Try
            Return False
        End Function

        Private Shared Function Paragraphs(page As UglyToad.PdfPig.Content.Page, marked As UglyToad.PdfPig.Content.Page, raw As OperationSource, targetPage As PdfSharp.Pdf.PdfPage, trace As System.Action(Of System.String)) As System.Collections.Generic.List(Of TranslationChunk)
            Dim ids = ObjectIds(page, marked, trace)
            Dim encodings = OriginalFonts(page, raw, ids)
            Dim epochs = PaintGroups(page)
            Dim links = page.GetHyperlinks()
            Dim words = page.GetWords(UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor.NearestNeighbourWordExtractor.Instance).ToList()
            Dim blocks = UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter.DocstrumBoundingBoxes.Instance.GetBlocks(words)
            Dim chunks As New System.Collections.Generic.List(Of TranslationChunk)()
            Dim verticalRules As New System.Collections.Generic.List(Of UglyToad.PdfPig.Core.PdfRectangle)()
            Dim obstacles As New System.Collections.Generic.List(Of UglyToad.PdfPig.Core.PdfRectangle)()
            Dim panels As New System.Collections.Generic.List(Of UglyToad.PdfPig.Core.PdfRectangle)()
            Dim allowGrowth = AddAnnotationObstacles(targetPage, obstacles, trace)
            For Each path In page.Paths
                For Each subpath In path
                    Dim rect = subpath.GetBoundingRectangle()
                    If rect.HasValue Then
                        If path.Count = 1 AndAlso AxisAlignedPanel(subpath) Then
                            panels.Add(rect.Value)
                        Else
                            obstacles.Add(rect.Value)
                        End If
                    End If
                    If rect.HasValue AndAlso rect.Value.Width <= 1 AndAlso rect.Value.Height > 1 Then verticalRules.Add(rect.Value)
                Next
            Next
            Dim imageBounds = page.GetImages().Select(Function(image) image.Bounds).ToArray()
            obstacles.AddRange(imageBounds)
            Dim textMargin = page.Letters.Where(Function(g) Not System.String.IsNullOrWhiteSpace(g.Value)).Select(Function(g) g.EndBaseLine.X).DefaultIfEmpty(0.0).Max()
            Dim seen As New System.Collections.Generic.HashSet(Of Glyph)()
            Dim paragraphNumber As System.Int32 = 0
            For Each block In blocks
                paragraphNumber += 1
                Dim chunk As TranslationChunk = Nothing
                For Each line In block.TextLines
                    Dim part As SourceLine = Nothing
                    For Each word In line.Words
                        Dim newWord As System.Boolean = True
                        For Each letter In word.Letters
                            If Not seen.Add(letter) Then Throw New System.IO.InvalidDataException("Duplicate source glyph.")
                            If letter.TextOrientation.ToString() <> "Horizontal" OrElse (CInt(letter.RenderingMode) <> 0 AndAlso CInt(letter.RenderingMode) <> 2) Then Throw New System.IO.InvalidDataException("Unsupported text orientation or paint mode: " & letter.TextOrientation.ToString() & " / " & CInt(letter.RenderingMode).ToString(Invariant) & ".")
                            Dim id = ids(letter)
                            Dim linkIndex As System.Int32 = -1
                            Dim box = letter.GlyphRectangle
                            Dim x = (box.Left + box.Right) / 2
                            Dim y = (box.Top + box.Bottom) / 2
                            For i As System.Int32 = 0 To links.Count - 1
                                Dim area = links(i).Bounds
                                If x >= area.Left AndAlso x <= area.Right AndAlso y >= area.Bottom AndAlso y <= area.Top Then
                                    linkIndex = i
                                    Exit For
                                End If
                            Next
                            Dim cellBoundary As System.Boolean = False
                            If part IsNot Nothing AndAlso newWord Then
                                Dim prior = part.Letters.Last()
                                Dim gap = letter.StartBaseLine.X - prior.EndBaseLine.X
                                cellBoundary = gap > 2.2 * letter.PointSize OrElse verticalRules.Any(Function(r) r.Bottom <= y AndAlso r.Top >= y AndAlso r.Left > prior.EndBaseLine.X AndAlso r.Right < letter.StartBaseLine.X)
                            End If
                            If cellBoundary OrElse chunk Is Nothing OrElse (CInt(letter.RenderingMode) = 2 AndAlso chunk.Anchor <> id) OrElse chunk.Epoch <> epochs(id) OrElse chunk.Link <> linkIndex OrElse Not SameFace(chunk.Lines(0).Letters(0), letter) Then
                                chunk = New TranslationChunk With {.ParagraphNumber = paragraphNumber, .Anchor = id, .Epoch = epochs(id), .Link = linkIndex}
                                chunks.Add(chunk)
                                part = Nothing
                            End If
                            If part Is Nothing Then
                                part = New SourceLine()
                                chunk.Lines.Add(part)
                            ElseIf newWord Then
                                part.Text.Append(" "c)
                            End If
                            part.Letters.Add(letter)
                            part.Text.Append(letter.Value)
                            newWord = False
                        Next
                    Next
                Next
            Next
            If page.Letters.Any(Function(g) Not System.String.IsNullOrWhiteSpace(g.Value) AndAlso Not seen.Contains(g)) Then Throw New System.IO.InvalidDataException("Incomplete paragraph extraction.")
            Dim wordStarts As New System.Collections.Generic.HashSet(Of Glyph)(words.Where(Function(w) w.Letters.Count > 0).Select(Function(w) w.Letters(0)))
            For Each chunk In chunks
                NormalizeRows(chunk, wordStarts, page.Letters, verticalRules)
                encodings.TryGetValue(chunk.Anchor, chunk.OriginalEncoding)
                chunk.Source = System.String.Join(" ", chunk.Lines.Select(Function(l) l.Text.ToString()))
                chunk.Left = chunk.Lines.Min(Function(l) l.Left)
                chunk.Right = chunk.Lines.Max(Function(l) l.Right)
                Dim minRight = chunk.Lines.Min(Function(l) l.Right)
                Dim maxLeft = chunk.Lines.Max(Function(l) l.Left)
                Dim centers = chunk.Lines.Select(Function(l) (l.Left + l.Right) / 2).ToArray()
                Dim tolerance = System.Math.Max(1, chunk.Lines(0).Letters(0).PointSize * 0.25)
                chunk.Alignment = "left"
                If maxLeft - chunk.Left > tolerance AndAlso centers.Max() - centers.Min() < tolerance Then
                    chunk.Alignment = "center"
                ElseIf (maxLeft - chunk.Left > tolerance AndAlso chunk.Right - minRight < tolerance) OrElse
                    System.Text.RegularExpressions.Regex.IsMatch(chunk.Source, "^\s*(?:[A-Z]{3}\s+)?[-+()]?[\d\s.,'’%€$£()-]+\s*$") Then
                    chunk.Alignment = "right"
                ElseIf chunk.Lines.Count = 1 AndAlso System.Math.Abs(centers(0) - page.Width / 2) <= tolerance Then
                    chunk.Alignment = "center"
                End If
                Dim own As New System.Collections.Generic.HashSet(Of Glyph)(chunk.Lines.SelectMany(Function(l) l.Letters))
                chunk.CanRetainOriginal = own.All(Function(g) ids(g) = chunk.Anchor) AndAlso
                    Not page.Letters.Any(Function(g) ids(g) = chunk.Anchor AndAlso Not own.Contains(g) AndAlso Not System.String.IsNullOrWhiteSpace(g.Value))
                Dim bottom = own.Min(Function(g) g.GlyphRectangle.Bottom)
                Dim top = own.Max(Function(g) g.GlyphRectangle.Top)
                Dim originalLeft = chunk.Left
                Dim originalRight = chunk.Right
                Dim leftRule = verticalRules.Where(Function(r) r.Top >= top AndAlso r.Bottom <= bottom AndAlso r.Right <= originalLeft).OrderByDescending(Function(r) r.Right).ToList()
                Dim rightRule = verticalRules.Where(Function(r) r.Top >= top AndAlso r.Bottom <= bottom AndAlso r.Left >= originalRight).OrderBy(Function(r) r.Left).ToList()
                If leftRule.Count > 0 AndAlso rightRule.Count > 0 Then
                    Dim padding = System.Math.Min(originalLeft - leftRule(0).Right, rightRule(0).Left - originalRight)
                    If chunk.Alignment = "right" Then
                        chunk.Left = leftRule(0).Right + padding
                    ElseIf chunk.Alignment = "left" Then
                        chunk.Right = rightRule(0).Left - padding
                    Else
                        Dim side = System.Math.Min(originalLeft - leftRule(0).Right, rightRule(0).Left - originalRight)
                        chunk.Left -= System.Math.Max(0, side - 2)
                        chunk.Right += System.Math.Max(0, side - 2)
                    End If
                End If
                chunk.Reflow = Not page.Letters.Any(Function(g) Not own.Contains(g) AndAlso Not System.String.IsNullOrWhiteSpace(g.Value) AndAlso
                    g.GlyphRectangle.Right > chunk.Left + 0.5 AndAlso g.GlyphRectangle.Left < chunk.Right - 0.5 AndAlso g.GlyphRectangle.Top > bottom + 0.5 AndAlso g.GlyphRectangle.Bottom < top - 0.5)
                If imageBounds.Any(Function(r) r.Top > bottom AndAlso r.Bottom < top AndAlso
                    ((r.Right > chunk.Left AndAlso r.Left < originalLeft) OrElse (r.Right > originalRight AndAlso r.Left < chunk.Right))) Then chunk.Reflow = False
                If Not chunk.Reflow Then
                    chunk.Left = originalLeft
                    chunk.Right = originalRight
                End If
                If allowGrowth AndAlso chunk.Reflow AndAlso chunk.Lines.Count = 1 AndAlso chunk.Alignment = "left" AndAlso
                    chunk.Link < 0 AndAlso chunk.Right = originalRight AndAlso chunk.Left = originalLeft Then
                    Dim limit = System.Math.Min(textMargin, page.CropBox.Bounds.Right - 2)
                    For Each panel In panels
                        If panel.Bottom > top + 2 OrElse panel.Top < bottom - 2 OrElse panel.Right < originalRight - 0.1 Then Continue For
                        Dim edge = If(panel.Left <= originalLeft + 0.1, panel.Right, panel.Left)
                        limit = System.Math.Min(limit, edge - 2)
                    Next
                    Dim blockers = page.Letters.Where(Function(g) Not own.Contains(g) AndAlso Not System.String.IsNullOrWhiteSpace(g.Value)).
                        Select(Function(g) g.GlyphRectangle).Concat(obstacles)
                    For Each box In blockers
                        If box.Bottom <= top + 2 AndAlso box.Top >= bottom - 2 AndAlso box.Right >= originalRight - 0.1 Then
                            limit = System.Math.Min(limit, box.Left - 2)
                        End If
                    Next
                    chunk.GrowthAllowance = System.Math.Max(0, limit - originalRight)
                End If
                If chunk.Link >= 0 Then
                    Dim area = links(chunk.Link).Bounds
                    chunk.Left = System.Math.Max(chunk.Left, area.Left)
                    chunk.Right = System.Math.Min(chunk.Right, area.Right)
                End If
            Next
            Return chunks
        End Function

        Private Shared Async Function TranslateNumberedPassages(group As TranslationChunk(), instruction As System.String,
                before As System.String, after As System.String, complete As System.Func(Of System.String, System.String, System.Threading.Tasks.Task(Of System.String)),
                cancel As System.Action) As System.Threading.Tasks.Task(Of System.String())
            Dim texts = group.Select(Function(c, i) "[" & (i + 1).ToString(Invariant) & "] " & c.Source)
            Dim input = "[CONTEXT BEFORE]" & before & "[/CONTEXT BEFORE]" & System.Environment.NewLine &
                "[TEXTTOPROCESS]" & System.Environment.NewLine & System.String.Join(System.Environment.NewLine, texts) & System.Environment.NewLine &
                "[/TEXTTOPROCESS]" & System.Environment.NewLine & "[CONTEXT AFTER]" & after & "[/CONTEXT AFTER]"
            cancel()
            Dim answer = Await complete(instruction, input)
            cancel()
            If System.String.IsNullOrWhiteSpace(answer) Then Return Nothing
            Dim matches = System.Text.RegularExpressions.Regex.Matches(answer.Trim(), "^\[(\d+)\][ \t]+(.+?)\s*$", System.Text.RegularExpressions.RegexOptions.Multiline)
            If matches.Count <> group.Length OrElse System.Text.RegularExpressions.Regex.Split(answer.Trim(), "\r?\n").Length <> group.Length Then Return Nothing
            Dim result(group.Length - 1) As System.String
            For i As System.Int32 = 0 To group.Length - 1
                If matches(i).Groups(1).Value <> (i + 1).ToString(Invariant) Then Return Nothing
                result(i) = matches(i).Groups(2).Value.Trim()
                If result(i).Length = 0 Then Return Nothing
            Next
            Return result
        End Function

        Private Shared Async Function DrawPassage(chunk As TranslationChunk, pageHeight As System.Double, graphics As PdfSharp.Drawing.XGraphics,
                location As System.String, instruction As System.String, retryPrompt As System.String, before As System.String, after As System.String,
                complete As System.Func(Of System.String, System.String, System.Threading.Tasks.Task(Of System.String)), cancel As System.Action,
                preferWording As System.Boolean, cached As System.String, trace As System.Action(Of System.String), substitutions As System.Collections.Generic.SortedSet(Of System.String), groupSeparator As System.String, noteNumber As System.Int32) As System.Threading.Tasks.Task
            Dim first = chunk.Lines(0).Letters(0)
            Dim name = first.FontName.ToLowerInvariant()
            Dim style = PdfSharp.Drawing.XFontStyleEx.Regular
            If name.Contains("bold") Then style = style Or PdfSharp.Drawing.XFontStyleEx.Bold
            If name.Contains("italic") OrElse name.Contains("oblique") Then style = style Or PdfSharp.Drawing.XFontStyleEx.Italic
            Dim widths = chunk.Lines.Select(Function(l) If(chunk.Reflow, chunk.Right - chunk.Left, l.Right - l.Left)).ToArray()
            Dim request = "[CONTEXT BEFORE]" & before & "[/CONTEXT BEFORE]" & System.Environment.NewLine &
                "[TEXTTOPROCESS]" & chunk.Source & "[/TEXTTOPROCESS]" & System.Environment.NewLine & "[CONTEXT AFTER]" & after & "[/CONTEXT AFTER]"
            Dim prompt = instruction
            Dim selected As System.String = Nothing
            Dim font As PdfSharp.Drawing.XFont = Nothing
            Dim useOriginal As System.Boolean = False
            Dim size As System.Double = first.PointSize
            Dim measure As System.Func(Of System.String, System.Double) = Nothing
            Dim horizontalScale As System.Double = 1.0
            Dim fitted As System.Collections.Generic.List(Of System.String) = Nothing
            Dim verifiedAnswer As System.String = Nothing
            For attempt As System.Int32 = 0 To 1
                cancel()
                Dim answer As System.String = cached
                If System.String.IsNullOrWhiteSpace(answer) Then answer = Await complete(prompt, request)
                cached = Nothing
                cancel()
                If attempt = 1 AndAlso System.String.IsNullOrWhiteSpace(answer) Then answer = verifiedAnswer
                If System.String.IsNullOrWhiteSpace(answer) Then Throw New System.IO.InvalidDataException(location & ": empty translation.")
                answer = answer.Trim()
                Dim valid = TokensPreserved(chunk.Source, answer, groupSeparator)
                If valid AndAlso chunk.CanRetainOriginal AndAlso answer = chunk.Source Then
                    chunk.RetainOriginal = True
                    fitted = chunk.Lines.Select(Function(l) l.Text.ToString()).ToList()
                ElseIf valid Then
                    If System.Text.RegularExpressions.Regex.IsMatch(answer, "[\u0590-\u0fff]") Then Throw New System.IO.InvalidDataException(location & ": this script requires a shaping or bidirectional renderer.")
                    If verifiedAnswer Is Nothing OrElse (Not preferWording AndAlso answer.Length < verifiedAnswer.Length) Then verifiedAnswer = answer
                    Dim scales As System.Double() = If(preferWording OrElse attempt = 1, New System.Double() {1, 0.97, 0.94}, New System.Double() {1})
                    If attempt = 1 AndAlso widths.Length = 1 AndAlso first.PointSize >= 7.5 Then
                        scales = scales.Concat(New System.Double() {0.9, 0.86, 0.84}).Where(Function(s) s * first.PointSize >= 7.5 OrElse s >= 0.94).ToArray()
                    End If
                    For Each scale In scales
                        size = first.PointSize * scale
                        Dim enc = chunk.OriginalEncoding
                        useOriginal = enc IsNot Nothing AndAlso enc.FontName = first.FontName AndAlso
                            System.Text.RegularExpressions.Regex.Replace(answer, "\s+", " ").All(Function(ch) enc.Codes.ContainsKey(ch) AndAlso Not enc.Ambiguous.Contains(ch))
                        If useOriginal Then
                            measure = Function(value) value.Sum(Function(ch) enc.Advances(ch)) * size
                            selected = enc.FontName & " (PDF resource " & enc.Resource & ")"
                        Else
                            font = ResolveFont(first.FontName, size, style, selected)
                            measure = Function(value) graphics.MeasureString(value, font).Width
                        End If
                        ' Try local layout before requesting another translation. Use the
                        ' widest permitted setting that fits all words in the original slots.
                        If chunk.GrowthAllowance > 0 Then
                            Dim required = measure(System.Text.RegularExpressions.Regex.Replace(answer, "\s+", " "))
                            Dim extra = System.Math.Min(chunk.GrowthAllowance, System.Math.Max(0, required - widths(0)))
                            If extra > 0 Then
                                widths(0) += extra
                                chunk.Right += extra
                                chunk.GrowthAllowance -= extra
                                trace($"{location} extended line by {extra:F2} pt into clear space")
                            End If
                        End If
                        Dim minimumWidthScale As System.Int32 = If(preferWording OrElse attempt = 1, 85, 94)
                        If scale < 0.94 Then minimumWidthScale = 80
                        For percent As System.Int32 = 100 To minimumWidthScale Step -1
                            horizontalScale = percent / 100.0
                            fitted = Wrap(answer, widths, Function(value) measure(value) * horizontalScale)
                            If fitted IsNot Nothing Then Exit For
                        Next
                        If fitted IsNot Nothing Then
                            If horizontalScale < 1 OrElse scale < 1 Then trace($"{location} recovered fit attempt={attempt + 1} size={size:F2} widthScale={horizontalScale:F2}")
                            Exit For
                        End If
                        Dim textWidth = measure(System.Text.RegularExpressions.Regex.Replace(answer, "\s+", " "))
                        trace($"{location} fit attempt={attempt + 1} size={size:F2} nativeFont={useOriginal} reflow={chunk.Reflow} slots={widths.Length} availablePt={widths.Sum():F1} textPt={textWidth:F1} minSlotPt={widths.Min():F1} maxSlotPt={widths.Max():F1}")
                    Next
                End If
                If fitted IsNot Nothing Then
                    chunk.Output = System.String.Join(" ", fitted)
                    Exit For
                End If
                If attempt = 1 Then
                    If verifiedAnswer IsNot Nothing Then
                        chunk.NoteText = System.Text.RegularExpressions.Regex.Replace(verifiedAnswer.Trim(), "\s+", " ")
                        Dim marker = "[" & noteNumber.ToString(Invariant) & "]"
                        size = System.Math.Min(first.PointSize, System.Math.Max(6.0, first.PointSize * 0.72))
                        font = ResolveFont("Arial", size, PdfSharp.Drawing.XFontStyleEx.Regular, selected)
                        measure = Function(value) graphics.MeasureString(value, font).Width
                        useOriginal = False
                        horizontalScale = System.Math.Min(1.0, System.Math.Max(0.1, widths(0)) / measure(marker))
                        fitted = New System.Collections.Generic.List(Of System.String) From {marker}
                        chunk.Output = marker
                        trace($"{location} translation note {noteNumber}: inline width {widths(0):F1} pt.")
                        Exit For
                    End If
                    Dim reason = "protected number, identifier, URL or placeholder differs"
                    trace(location & ": " & reason & ". Source: " & chunk.Source & " / Translation: " & answer)
                    Throw New System.IO.InvalidDataException(location & ": " & reason & ".")
                End If
                prompt = instruction & System.Environment.NewLine & retryPrompt
                If valid Then request &= System.Environment.NewLine & "[LAYOUT] lines=" & widths.Length.ToString(Invariant) & " availablePoints=" & widths.Sum().ToString("0.#", Invariant) & " [/LAYOUT]"
                If Not valid Then request &= System.Environment.NewLine & "[PROTECTED SOURCE TOKENS]" & System.String.Join(" | ", ProtectedTokens(chunk.Source, originalSpelling:=True)) & "[/PROTECTED SOURCE TOKENS]"
                request &= System.Environment.NewLine & "[PREVIOUS TRANSLATION]" & answer & "[/PREVIOUS TRANSLATION]"
            Next
            Dim rgb = first.FillColor.ToRGBValues()
            Dim brush As New PdfSharp.Drawing.XSolidBrush(PdfSharp.Drawing.XColor.FromArgb(CInt(rgb.Item1 * 255), CInt(rgb.Item2 * 255), CInt(rgb.Item3 * 255)))
            Dim native As New System.Text.StringBuilder()
            For i As System.Int32 = 0 To fitted.Count - 1
                If fitted(i).Length = 0 Then Continue For
                Dim line = chunk.Lines(i)
                Dim x = line.Left
                If Not chunk.RetainOriginal Then
                    Dim measured = measure(fitted(i)) * horizontalScale
                    Dim left = If(chunk.Reflow, chunk.Left, line.Left)
                    Dim right = If(chunk.Reflow, chunk.Right, line.Right)
                    Select Case chunk.Alignment
                        Case "right"
                            x = right - measured
                        Case "center"
                            x = (left + right - measured) / 2
                        Case Else
                            x = left
                    End Select
                    If useOriginal Then
                        native.Append(OriginalTextCommands(chunk.OriginalEncoding, fitted(i), x, line.Letters(0).StartBaseLine.Y, size, horizontalScale, rgb))
                    ElseIf horizontalScale = 1 Then
                        graphics.DrawString(fitted(i), font, brush, x, pageHeight - line.Letters(0).StartBaseLine.Y)
                    Else
                        Dim state = graphics.Save()
                        Try
                            graphics.ScaleTransform(horizontalScale, 1.0)
                            graphics.DrawString(fitted(i), font, brush, x / horizontalScale, pageHeight - line.Letters(0).StartBaseLine.Y)
                        Finally
                            graphics.Restore(state)
                        End Try
                    End If
                End If
                Dim tolerance = System.Math.Max(1.0, first.PointSize * 0.15)
                Dim original = System.Drawing.RectangleF.FromLTRB(CSng(line.Letters.Min(Function(g) g.GlyphRectangle.Left)), CSng(line.Letters.Min(Function(g) g.GlyphRectangle.Bottom)),
                    CSng(line.Letters.Max(Function(g) g.GlyphRectangle.Right)), CSng(line.Letters.Max(Function(g) g.GlyphRectangle.Top)))
                Dim topLimit = line.Letters.Max(Function(g) g.GlyphRectangle.Top)
                Dim bottomLimit = line.Letters.Min(Function(g) g.GlyphRectangle.Bottom)
                If chunk.NoteText IsNot Nothing Then
                    topLimit = line.Letters(0).StartBaseLine.Y + size
                    bottomLimit = line.Letters(0).StartBaseLine.Y - size * 0.3
                End If
                chunk.Painted.Add(New PaintedLine With {.OriginalRectangle = original, .Baseline = line.Letters(0).StartBaseLine.Y, .Passage = location, .LinkIndex = chunk.Link, .Count = CanonicalText(fitted(i)).Length,
                    .Left = x - tolerance, .Right = If(chunk.Reflow, chunk.Right, line.Right) + tolerance,
                    .Top = topLimit + tolerance,
                    .Bottom = bottomLimit - tolerance})
            Next
            If useOriginal Then chunk.Commands = System.Text.Encoding.ASCII.GetBytes(native.ToString())
            If Not chunk.RetainOriginal Then substitutions.Add(first.FontName & " → " & selected & " (" & size.ToString("0.##", Invariant) & " pt)")
        End Function

        Private Shared Function TargetGroupSeparator(language As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(language) Then Return Nothing
            Dim key = language.Trim()
            For Each culture In System.Globalization.CultureInfo.GetCultures(System.Globalization.CultureTypes.AllCultures)
                If New System.String() {culture.Name, culture.EnglishName, culture.NativeName, culture.DisplayName}.Contains(key, System.StringComparer.OrdinalIgnoreCase) Then
                    Dim format = culture.NumberFormat
                    If format.NumberGroupSizes.SequenceEqual(New System.Int32() {3}) Then Return format.NumberGroupSeparator
                    Return Nothing
                End If
            Next
            Return Nothing
        End Function

        Private Shared Function TokensPreserved(source As System.String, answer As System.String, separator As System.String) As System.Boolean
            Dim originals = ProtectedTokens(source).ToList()
            Dim remaining As New System.Collections.Generic.List(Of System.String)()
            For Each token In ProtectedTokens(answer)
                If Not originals.Remove(token) Then remaining.Add(token)
            Next
            If remaining.Count <> originals.Count Then Return False
            If remaining.Count = 0 Then Return True
            If System.String.IsNullOrEmpty(separator) Then Return False
            Dim grouped = "\A[-+]?[1-9][0-9]{0,2}(?:" & System.Text.RegularExpressions.Regex.Escape(separator) & "[0-9]{3})+\z"
            For Each token In remaining
                If Not System.Text.RegularExpressions.Regex.IsMatch(token, grouped) OrElse Not originals.Remove(token.Replace(separator, "")) Then Return False
            Next
            Return True
        End Function

        Private Shared Function ProtectedTokens(value As System.String, Optional originalSpelling As System.Boolean = False) As System.String()
            Const pattern As System.String = "\{\{[^{}]+\}\}|\{[\w.:-]+\}|https?://[^\s<>]+|[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}|(?<!\w)(?=[A-Za-z\d._/-]*\d)[A-Za-z]+[-_][A-Za-z\d._/-]+|[-−+]?\d+(?:[.,’']\d+|[\u00A0\u202F]\d{3}(?!\d))*|[%€£$]"
            Return System.Text.RegularExpressions.Regex.Matches(value, pattern).Cast(Of System.Text.RegularExpressions.Match)().
                Select(Function(m) If(originalSpelling, m.Value, System.Text.RegularExpressions.Regex.Replace(m.Value.Replace("−", "-"), "^([-+]?\d+)[,.](\d{1,2}|\d{4,})$", "$1.$2"))).
                OrderBy(Function(s) s, System.StringComparer.Ordinal).ToArray()
        End Function
    End Class
End Namespace
