Option Explicit On
Option Strict On
Option Infer On

Imports System.Linq

Namespace Agents
    Partial Public NotInheritable Class PdfFileTranslation
        Private Sub New()
        End Sub

        Private Shared Sub WriteMarginNotes(target As PdfSharp.Pdf.PdfPage, page As UglyToad.PdfPig.Content.Page,
                chunks As System.Collections.Generic.List(Of TranslationChunk), lines As System.Collections.Generic.List(Of PaintedLine),
                translated As System.Text.StringBuilder, endNotes As System.Collections.Generic.List(Of System.String),
                cancel As System.Action, trace As System.Action(Of System.String))
            Dim occupied = GraphicBounds(page)
            Dim knownAnnotations = AddAnnotationObstacles(target, occupied, trace)
            occupied.AddRange(page.Letters.Select(Function(g) g.GlyphRectangle))
            For Each line In lines
                occupied.Add(New UglyToad.PdfPig.Core.PdfRectangle(line.Left, line.Bottom, line.Right, line.Top))
            Next
            Dim selected As System.String = Nothing
            Dim font = ResolveFont("Arial", 9, PdfSharp.Drawing.XFontStyleEx.Regular, selected)
            Dim depth = CInt(System.Math.Min(60, page.Height * 0.1))
            Dim baselines As New System.Collections.Generic.List(Of System.Int32)()
            For marginOffset As System.Int32 = depth To 9 Step -2
                baselines.Add(marginOffset)
            Next
            For marginOffset As System.Int32 = 14 To depth Step 2
                baselines.Add(CInt(System.Math.Floor(page.Height)) - marginOffset)
            Next
            Using gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(target)
                For Each chunk In chunks.Where(Function(c) c.NoteText IsNot Nothing)
                    cancel()
                    Dim label = chunk.Output & " " & chunk.NoteText
                    Dim inset = System.Math.Min(36.0, page.Width * 0.06)
                    Dim labelWidth = gfx.MeasureString(label, font).Width
                    Dim placed As System.Boolean = False
                    For Each x As System.Double In New System.Double() {inset, 9.0, page.Width - inset - labelWidth}
                        If Not knownAnnotations Then Exit For
                        Dim right = x + labelWidth
                        If x < 6 OrElse right + 6 > page.Width Then Continue For
                        For Each baseline As System.Int32 In baselines
                            Dim band As New UglyToad.PdfPig.Core.PdfRectangle(x - 3, baseline - 4, right + 3, baseline + 11)
                            If occupied.Any(Function(r) r.Right > band.Left AndAlso r.Left < band.Right AndAlso r.Top > band.Bottom AndAlso r.Bottom < band.Top) Then Continue For
                            gfx.DrawString(label, font, PdfSharp.Drawing.XBrushes.Black, x, page.Height - baseline)
                            occupied.Add(band)
                            translated.Append(label)
                            lines.Add(New PaintedLine With {.Baseline = baseline, .Count = CanonicalText(label).Length, .LinkIndex = -1,
                                .Passage = "Translation note " & chunk.Output, .Left = band.Left, .Right = band.Right, .Bottom = band.Bottom, .Top = band.Top})
                            placed = True
                            Exit For
                        Next
                        If placed Then Exit For
                    Next
                    If Not placed Then endNotes.Add(chunk.Output & " (p. " & page.Number.ToString(Invariant) & ") " & chunk.NoteText)
                    trace($"Page {page.Number} note {chunk.Output} placement={If(placed, "margin", "end")}")
                Next
            End Using
        End Sub

        ' Use ordinary searchable PDF text for overflow so that printed copies carry the translation.
        Private Shared Sub WriteEndNotes(document As PdfSharp.Pdf.PdfDocument, notes As System.Collections.Generic.List(Of System.String),
                expected As System.Collections.Generic.List(Of System.String), cancel As System.Action)
            Dim selected As System.String = Nothing
            Dim font = ResolveFont("Arial", 11, PdfSharp.Drawing.XFontStyleEx.Regular, selected)
            Dim rows As New System.Collections.Generic.List(Of System.String)()
            Using measuringDocument As New PdfSharp.Pdf.PdfDocument(), measure = PdfSharp.Drawing.XGraphics.FromPdfPage(measuringDocument.AddPage())
                For Each note In notes
                    cancel()
                    Dim row As System.String = ""
                    For Each word In System.Text.RegularExpressions.Regex.Split(note.Trim(), "\s+")
                        Dim candidate = If(row.Length = 0, word, row & " " & word)
                        If measure.MeasureString(candidate, font).Width <= 515 Then
                            row = candidate
                            Continue For
                        End If
                        If row.Length > 0 Then rows.Add(row)
                        row = ""
                        Dim parts = System.Globalization.StringInfo.GetTextElementEnumerator(word)
                        While parts.MoveNext()
                            Dim part = parts.GetTextElement()
                            If row.Length > 0 AndAlso measure.MeasureString(row & part, font).Width > 515 Then
                                rows.Add(row)
                                row = ""
                            End If
                            row &= part
                        End While
                    Next
                    If row.Length > 0 Then rows.Add(row)
                Next
            End Using
            If CanonicalText(System.String.Concat(rows)) <> CanonicalText(System.String.Concat(notes)) Then Throw New System.IO.InvalidDataException("Translation note text was lost during pagination.")
            For offset As System.Int32 = 0 To rows.Count - 1 Step 49
                cancel()
                Dim sheet = document.AddPage()
                sheet.Width = PdfSharp.Drawing.XUnit.FromPoint(595)
                sheet.Height = PdfSharp.Drawing.XUnit.FromPoint(842)
                Const title As System.String = "Translation notes"
                Dim content As New System.Text.StringBuilder(title)
                Using gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(sheet)
                    gfx.DrawString(title, font, PdfSharp.Drawing.XBrushes.Black, 40, 40)
                    For i As System.Int32 = 0 To System.Math.Min(49, rows.Count - offset) - 1
                        gfx.DrawString(rows(offset + i), font, PdfSharp.Drawing.XBrushes.Black, 40, 64 + i * 15)
                        content.Append(rows(offset + i))
                    Next
                End Using
                expected.Add(CanonicalText(content.ToString()))
            Next
        End Sub

        Private Shared Function CanonicalText(value As System.String) As System.String
            Return System.Text.RegularExpressions.Regex.Replace(value.Normalize(), "\s+", "")
        End Function

        ' Coordinates are PDF points. A writer may serialize the same edge with fewer decimals.
        Private Shared Sub VerifyRectangle(expected As UglyToad.PdfPig.Core.PdfRectangle, actual As UglyToad.PdfPig.Core.PdfRectangle,
                location As System.String, trace As System.Action(Of System.String))
            Const tolerance As System.Double = 0.001
            If System.Math.Abs(expected.Left - actual.Left) <= tolerance AndAlso
                System.Math.Abs(expected.Bottom - actual.Bottom) <= tolerance AndAlso
                System.Math.Abs(expected.Right - actual.Right) <= tolerance AndAlso
                System.Math.Abs(expected.Top - actual.Top) <= tolerance Then Return
            Dim detail = System.FormattableString.Invariant($"{location}: expected [{expected.Left:R}, {expected.Bottom:R}, {expected.Right:R}, {expected.Top:R}], saved [{actual.Left:R}, {actual.Bottom:R}, {actual.Right:R}, {actual.Top:R}] (pt).")
            trace(detail)
            Throw New System.IO.InvalidDataException(detail)
        End Sub

        Private Shared Sub VerifyPageBoxes(expected As UglyToad.PdfPig.Content.Page, actual As UglyToad.PdfPig.Content.Page,
                location As System.String, trace As System.Action(Of System.String))
            VerifyRectangle(expected.MediaBox.Bounds, actual.MediaBox.Bounds, location & " MediaBox", trace)
            VerifyRectangle(expected.CropBox.Bounds, actual.CropBox.Bounds, location & " CropBox", trace)
        End Sub

        Private Shared Function Wrap(text As System.String, widths As System.Double(), measure As System.Func(Of System.String, System.Double)) As System.Collections.Generic.List(Of System.String)
            Dim lines As New System.Collections.Generic.List(Of System.String)()
            Dim current As System.String = ""
            For Each word In System.Text.RegularExpressions.Regex.Split(text.Trim(), "\s+")
                Dim candidate = If(current.Length = 0, word, current & " " & word)
                If measure(candidate) <= widths(lines.Count) + 0.5 Then
                    current = candidate
                    Continue For
                End If
                If current.Length > 0 Then lines.Add(current)
                Dim wordWidth = measure(word)
                While lines.Count < widths.Length AndAlso wordWidth > widths(lines.Count) + 0.5
                    lines.Add(System.String.Empty)
                End While
                If lines.Count = widths.Length Then Return Nothing
                current = word
            Next
            lines.Add(current)
            Return lines
        End Function

        Private Class PaintedLine
            Public Baseline As System.Double
            Public Passage As System.String
            Public Count As System.Int32
            Public Left As System.Double
            Public Right As System.Double
            Public Top As System.Double
            Public Bottom As System.Double
            Public OriginalRectangle As System.Drawing.RectangleF
            Public TranslatedRectangle As System.Drawing.RectangleF
            Public LinkIndex As System.Int32
        End Class

        Private Shared Function GraphicBounds(page As UglyToad.PdfPig.Content.Page) As System.Collections.Generic.List(Of UglyToad.PdfPig.Core.PdfRectangle)
            Dim areas = page.GetImages().Select(Function(image) image.Bounds).ToList()
            For Each path In page.Paths
                If Not path.IsFilled AndAlso Not path.IsStroked Then Continue For
                For Each part In path
                    Dim rect = part.GetBoundingRectangle()
                    If rect.HasValue Then areas.Add(rect.Value)
                Next
            Next
            Return areas
        End Function

        Private Const InkMargin As System.Double = 72.0 / 25.4

        Private Shared Function WithinInkMargin(line As PaintedLine, glyph As UglyToad.PdfPig.Content.Letter,
                crop As UglyToad.PdfPig.Core.PdfRectangle) As System.Boolean
            Dim ink = glyph.GlyphRectangle
            If System.Math.Abs(glyph.StartBaseLine.Y - line.Baseline) > 0.5 OrElse
                ink.Left < line.Left - InkMargin OrElse ink.Right > line.Right + InkMargin OrElse
                ink.Bottom < line.Bottom - InkMargin OrElse ink.Top > line.Top + InkMargin Then Return False
            If ink.Left < crop.Left OrElse ink.Right > crop.Right OrElse ink.Bottom < crop.Bottom OrElse ink.Top > crop.Top Then Return False
            Return True
        End Function

        Private Shared Function TouchesGraphicInMargin(line As PaintedLine, ink As UglyToad.PdfPig.Core.PdfRectangle,
                obstacles As System.Collections.Generic.List(Of UglyToad.PdfPig.Core.PdfRectangle)) As System.Boolean
            For Each obstacle In obstacles
                ' Pad the path bounds so a vertical or horizontal rule also occupies space.
                Dim left = System.Math.Max(ink.Left, obstacle.Left - 0.5)
                Dim right = System.Math.Min(ink.Right, obstacle.Right + 0.5)
                Dim bottom = System.Math.Max(ink.Bottom, obstacle.Bottom - 0.5)
                Dim top = System.Math.Min(ink.Top, obstacle.Top + 0.5)
                If left < right AndAlso bottom < top AndAlso
                    (left < line.Left OrElse right > line.Right OrElse bottom < line.Bottom OrElse top > line.Top) Then Return True
            Next
            Return False
        End Function

        Private Shared Sub CheckOverlaps(lines As System.Collections.Generic.List(Of PaintedLine), pageNumber As System.Int32, review As System.Action(Of System.Int32, System.String))
            Dim sorted = lines.OrderBy(Function(l) l.TranslatedRectangle.Top).ToList()
            For i As System.Int32 = 0 To sorted.Count - 1
                Dim a = sorted(i)
                For j As System.Int32 = i + 1 To sorted.Count - 1
                    Dim b = sorted(j)
                    If b.TranslatedRectangle.Top >= a.TranslatedRectangle.Bottom Then Exit For
                    Dim overlap = System.Drawing.RectangleF.Intersect(a.TranslatedRectangle, b.TranslatedRectangle)
                    If overlap.Width <= 0.75 OrElse overlap.Height <= 0.75 Then Continue For
                    Dim prior = System.Drawing.RectangleF.Intersect(a.OriginalRectangle, b.OriginalRectangle)
                    If prior.Width > 0.75 AndAlso prior.Height > 0.75 Then Continue For
                    Dim detail = $"Page {pageNumber}: new text bounds overlap {overlap.Width:F2} x {overlap.Height:F2} pt near {overlap.X:F1}, {overlap.Y:F1}."
                    If System.Math.Min(overlap.Width, overlap.Height) > InkMargin Then Throw New System.IO.InvalidDataException(detail)
                    review(pageNumber, detail)
                Next
            Next
        End Sub

        Private Shared Function ResolveFont(sourceName As System.String, size As System.Double, style As PdfSharp.Drawing.XFontStyleEx, ByRef selected As System.String) As PdfSharp.Drawing.XFont
            WindowsPdfFontResolver.EnsureAvailable()
            Dim name = System.Text.RegularExpressions.Regex.Replace(sourceName, "^[A-Z]{6}\+", "")
            Dim family = System.Text.RegularExpressions.Regex.Replace(name, "[-,](BoldItalic|BoldOblique|Bold|Italic|Oblique|Regular|Roman|Medium)(MT)?$", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            Dim choices As New System.Collections.Generic.List(Of System.String) From {family}
            Dim plain = family.ToLowerInvariant()
            If plain.Contains("cambria") Then choices.AddRange({"Cambria", "Caladea", "Times New Roman"})
            If plain.Contains("calibri") Then choices.AddRange({"Calibri", "Carlito", "Arial"})
            If plain.Contains("times") OrElse plain.Contains("serif") Then choices.AddRange({"Times New Roman", "Liberation Serif"})
            If plain.Contains("courier") OrElse plain.Contains("mono") Then choices.AddRange({"Courier New", "Liberation Mono"})
            choices.AddRange({"Arial", "Liberation Sans"})
            For Each candidate In choices.Distinct(System.StringComparer.OrdinalIgnoreCase)
                Try
                    Dim font As New PdfSharp.Drawing.XFont(candidate, size, style)
                    selected = font.FontFamily.Name
                    Return font
                Catch ex As System.InvalidOperationException
                    Continue For
                End Try
            Next
            Throw New System.IO.InvalidDataException("No usable font for " & sourceName)
        End Function

        Private Shared Sub RemovePageText(document As PdfSharp.Pdf.PdfDocument, target As PdfSharp.Pdf.PdfPage, source As UglyToad.PdfPig.Content.Page)
            Using stream As New System.IO.MemoryStream()
                For Each operation In source.Operations
                    Select Case operation.[Operator]
                        Case "Tj", "TJ", "'", """"
                            ' Only painting operators are removed. Non-text graphics retain their original order.
                        Case Else
                            operation.Write(stream)
                            stream.WriteByte(10)
                    End Select
                Next
                Dim replacement As New PdfSharp.Pdf.PdfDictionary(document)
                replacement.CreateStream(stream.ToArray())
                document.Internals.AddObject(replacement)
                target.Elements("/Contents") = replacement.Reference
            End Using
        End Sub

        Public Shared Function TranslateWithModelAsync(source As System.String, destination As System.String,
                instruction As System.String, context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext,
                selectedSecondApi As System.Boolean, cancel As System.Action, report As System.Action(Of System.String),
                token As System.Threading.CancellationToken, Optional replaceExisting As System.Boolean = False, Optional preferWording As System.Boolean = False, Optional targetLanguage As System.String = Nothing) As System.Threading.Tasks.Task(Of System.String)
            Dim callContext As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext
            Dim secondApi = selectedSecondApi
            If selectedSecondApi Then
                callContext = Global.SharedLibrary.SharedLibrary.SharedMethods.CreateIsolatedModelCallContext(context)
            Else
                Dim model = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveIsolatedSpecialTaskModel(context, "OfflineDocs")
                callContext = model.Context
                secondApi = model.UsesSecondApi
            End If
            Dim complete As System.Func(Of System.String, System.String, System.Threading.Tasks.Task(Of System.String)) =
                Function(systemText, userText) Global.SharedLibrary.SharedLibrary.SharedMethods.LLM(callContext, systemText, userText, UseSecondAPI:=secondApi, HideSplash:=True, cancellationToken:=token)
            Return TranslateAsync(source, destination, instruction & System.Environment.NewLine & context.SP_Translate_Pdf_Line,
                                  context.SP_Translate_Pdf_Retry, complete, cancel, report, replaceExisting, preferWording, context.SP_Translate_Pdf_Group, targetLanguage)
        End Function

        Public Shared Async Function TranslateAsync(source As System.String, destination As System.String,
                translationPrompt As System.String, retryPrompt As System.String,
                complete As System.Func(Of System.String, System.String, System.Threading.Tasks.Task(Of System.String)),
                cancel As System.Action, report As System.Action(Of System.String), Optional replaceExisting As System.Boolean = False, Optional preferWording As System.Boolean = False, Optional groupPrompt As System.String = Nothing, Optional targetLanguage As System.String = Nothing) As System.Threading.Tasks.Task(Of System.String)
            If System.String.Equals(System.IO.Path.GetFullPath(source), System.IO.Path.GetFullPath(destination), System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.ArgumentException("Choose a separate translated PDF.")
            End If
            Dim groupSeparator = TargetGroupSeparator(targetLanguage)
            Dim expected As New System.Collections.Generic.List(Of System.String)()
            Dim positions As New System.Collections.Generic.List(Of System.Collections.Generic.List(Of PaintedLine))()
            Dim substitutions As New System.Collections.Generic.SortedSet(Of System.String)()
            Dim temporary = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(destination)), System.Guid.NewGuid().ToString("N") & ".pdf")
            Dim cache As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
            Dim nextNote As System.Int32 = 1
            Dim notePages As New System.Collections.Generic.SortedSet(Of System.Int32)()
            Dim endNotes As New System.Collections.Generic.List(Of System.String)()
            Dim imagePages As New System.Collections.Generic.List(Of System.Int32)()
            Dim logPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "RedInk-PdfTranslation-" & System.Guid.NewGuid().ToString("N") & ".log")
            Dim writer As System.IO.StreamWriter = Nothing
            Try
                writer = New System.IO.StreamWriter(logPath, False, System.Text.Encoding.UTF8) With {.AutoFlush = True}
            Catch ex As System.IO.IOException
                ' Diagnostics must not prevent a translation.
            Catch ex As System.UnauthorizedAccessException
            End Try
            Dim trace As System.Action(Of System.String) =
                Sub(message)
                    If writer Is Nothing Then Return
                    Try
                        writer.WriteLine(System.DateTime.UtcNow.ToString("O", Invariant) & " " & message)
                    Catch ex As System.IO.IOException
                        System.Diagnostics.Debug.WriteLine(ex.Message)
                    End Try
                End Sub
            Dim modelCalls As System.Int32 = 0
            Dim modelMilliseconds As System.Int64 = 0
            Dim timedComplete As System.Func(Of System.String, System.String, System.Threading.Tasks.Task(Of System.String)) =
                Async Function(systemText As System.String, userText As System.String) As System.Threading.Tasks.Task(Of System.String)
                    modelCalls += 1
                    Dim timer = System.Diagnostics.Stopwatch.StartNew()
                    Try
                        Return Await complete(systemText, userText)
                    Finally
                        modelMilliseconds += timer.ElapsedMilliseconds
                        trace($"model call={modelCalls} ms={timer.ElapsedMilliseconds} inputChars={userText.Length}")
                    End Try
                End Function
            Dim completeTextVerified As System.Boolean = False
            Dim retainLog As System.Boolean = False
            Dim reviewPages As New System.Collections.Generic.SortedSet(Of System.Int32)()
            Dim review As System.Action(Of System.Int32, System.String) =
                Sub(number, detail)
                    reviewPages.Add(number)
                    trace(detail)
                End Sub
            Dim loadedLibrary = GetType(PdfFileTranslation).Assembly
            trace($"visualVerification=advisory-v1 dll={loadedLibrary.Location} mvid={loadedLibrary.ManifestModule.ModuleVersionId}")
            trace("begin source=" & System.IO.Path.GetFileName(source) & " pdfpig=" & GetType(UglyToad.PdfPig.PdfDocument).Assembly.GetName().Version.ToString() &
                " pdfsharp=" & GetType(PdfSharp.Pdf.PdfDocument).Assembly.GetName().Version.ToString())
            Try
                Using original = UglyToad.PdfPig.PdfDocument.Open(source),
                      document = PdfSharp.Pdf.IO.PdfReader.Open(source, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)
                    For number As System.Int32 = 1 To original.NumberOfPages
                        cancel()
                        Dim page = original.GetPage(number)
                        Dim target = document.Pages(number - 1)
                        Dim hasImages = page.GetImages().Any()
                        If hasImages Then imagePages.Add(number)
                        If page.Letters.Count = 0 Then
                            If hasImages Then Throw New System.IO.InvalidDataException($"Page {number} contains images without selectable text.")
                            Continue For
                        End If
                        If target.Rotate Mod 360 <> 0 OrElse target.MediaBox.X1 <> 0 OrElse target.MediaBox.Y1 <> 0 OrElse
                            page.CropBox.Bounds.Left <> 0 OrElse page.CropBox.Bounds.Bottom <> 0 OrElse System.Math.Abs(page.Height - target.Height.Point) > 0.1 Then
                            Throw New System.IO.InvalidDataException($"Page {number} uses an unsupported page rotation or crop.")
                        End If
                        RemovePageText(document, target, page)
                    Next
                    ' Text inside nested forms is outside the rewritten page stream. Reject the file before translating.
                    Using stripped As New System.IO.MemoryStream()
                        document.Save(stripped, False)
                        Using residual = UglyToad.PdfPig.PdfDocument.Open(stripped.ToArray())
                            For number As System.Int32 = 1 To residual.NumberOfPages
                                Dim savedPage = residual.GetPage(number)
                                VerifyPageBoxes(original.GetPage(number), savedPage, $"Page {number} after first save", trace)
                                If savedPage.Letters.Count > 0 Then Throw New System.IO.InvalidDataException($"Page {number} contains nested PDF text that cannot be replaced safely.")
                            Next
                        End Using
                        ' Reopen the stripped copy because PDFsharp finalizes documents on Save.
                        stripped.Position = 0
                        Using output = PdfSharp.Pdf.IO.PdfReader.Open(stripped, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)
                            Dim sources As New System.Collections.Generic.List(Of OperationSource)()
                            Using marked = UglyToad.PdfPig.PdfDocument.Open(MarkTextObjects(source, original, cancel, sources, trace))
                                Dim preceding As System.String = ""
                                Dim plans As New System.Collections.Generic.List(Of System.Collections.Generic.List(Of TranslationChunk))()
                                For pageNumber As System.Int32 = 1 To original.NumberOfPages
                                    cancel()
                                    trace($"preflight page={pageNumber}")
                                    report($"Checking PDF layout: page {pageNumber}/{original.NumberOfPages}")
                                    plans.Add(Paragraphs(original.GetPage(pageNumber), marked.GetPage(pageNumber), sources(pageNumber - 1), output.Pages(pageNumber - 1), trace))
                                Next
                                For number As System.Int32 = 1 To original.NumberOfPages
                                    Dim chunks = plans(number - 1)
                                    Dim upcoming = If(number < plans.Count, plans(number), Nothing)
                                    cancel()
                                    Dim page = original.GetPage(number)
                                    Dim pageLines As New System.Collections.Generic.List(Of PaintedLine)()
                                    positions.Add(pageLines)
                                    Dim translated As New System.Text.StringBuilder()
                                    If page.Letters.Count > 0 Then
                                        Dim prepared As New System.Collections.Generic.Dictionary(Of TranslationChunk, System.String)()
                                        Dim nextBatch As System.Int32 = 0
                                        Dim target = output.Pages(number - 1)
                                        For index As System.Int32 = 0 To chunks.Count - 1
                                            cancel()
                                            Dim chunk = chunks(index)
                                            Dim location = $"Page {number}, passage {index + 1}"
                                            trace(location)
                                            report($"Page {number}/{original.NumberOfPages}, passage {index + 1}/{chunks.Count}")
                                            Dim before = If(index > 0, chunks(index - 1).Source, preceding)
                                            Dim after = If(index + 1 < chunks.Count, chunks(index + 1).Source, If(upcoming IsNot Nothing AndAlso upcoming.Count > 0, upcoming(0).Source, ""))
                                            Dim first = chunk.Lines(0).Letters(0)
                                            Dim marginal = first.StartBaseLine.Y < page.Height * 0.1 OrElse first.StartBaseLine.Y > page.Height * 0.9
                                            Dim keyBefore = If(marginal AndAlso index > 0 AndAlso System.Math.Abs(chunks(index - 1).Lines(0).Letters(0).StartBaseLine.Y - first.StartBaseLine.Y) < first.PointSize * 2, before, "")
                                            Dim keyAfter = If(marginal AndAlso index + 1 < chunks.Count AndAlso System.Math.Abs(chunks(index + 1).Lines(0).Letters(0).StartBaseLine.Y - first.StartBaseLine.Y) < first.PointSize * 2, after, "")
                                            Dim cacheKey = System.String.Join(ChrW(30), {chunk.Source, first.FontName, chunk.Alignment, keyBefore, keyAfter,
                                                System.Math.Round(chunk.Left, 1).ToString(Invariant), System.Math.Round(chunk.Right, 1).ToString(Invariant)})
                                            Dim cached As System.String = Nothing
                                            If marginal Then cache.TryGetValue(cacheKey, cached)
                                            If cached Is Nothing AndAlso Not System.String.IsNullOrWhiteSpace(groupPrompt) Then
                                                If index >= nextBatch Then
                                                    Dim chars As System.Int32 = 0
                                                    nextBatch = index
                                                    While nextBatch < chunks.Count AndAlso nextBatch - index < 16
                                                        Dim length = chunks(nextBatch).Source.Length
                                                        If nextBatch > index AndAlso chars + length > 5000 Then Exit While
                                                        chars += length
                                                        nextBatch += 1
                                                    End While
                                                    Dim group = chunks.GetRange(index, nextBatch - index).ToArray()
                                                    If group.Length > 1 Then
                                                        Dim following = If(nextBatch < chunks.Count, chunks(nextBatch).Source, If(upcoming IsNot Nothing AndAlso upcoming.Count > 0, upcoming(0).Source, ""))
                                                        trace($"batch page={number} first={index + 1} count={group.Length} chars={chars}")
                                                        Dim wording = Await TranslateNumberedPassages(group, translationPrompt & System.Environment.NewLine & groupPrompt, before, following, timedComplete, cancel)
                                                        If wording Is Nothing Then
                                                            trace($"batch rejected page={number} first={index + 1}")
                                                            report(location & ": invalid numbered response; retrying individual passages with context.")
                                                        Else
                                                            For k As System.Int32 = 0 To group.Length - 1
                                                                prepared(group(k)) = wording(k)
                                                            Next
                                                        End If
                                                    End If
                                                End If
                                                prepared.TryGetValue(chunk, cached)
                                            End If
                                            Using gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(target)
                                                Await DrawPassage(chunk, page.Height, gfx, location, translationPrompt, retryPrompt, before, after,
                                                    timedComplete, cancel, preferWording, cached, trace, substitutions, groupSeparator, nextNote)
                                            End Using
                                            Dim contents = target.Contents
                                            If Not chunk.RetainOriginal AndAlso chunk.Commands Is Nothing Then chunk.Commands = contents.Elements.GetDictionary(contents.Elements.Count - 1).Stream.UnfilteredValue
                                            If chunk.NoteText IsNot Nothing Then
                                                nextNote += 1
                                                notePages.Add(number)
                                            End If
                                            If marginal Then cache(cacheKey) = If(chunk.NoteText, chunk.Output)
                                        Next
                                        If chunks.Count > 0 Then preceding = chunks.Last().Source
                                        InsertAtOriginalDepth(output, target, page, sources(number - 1), chunks)
                                        For Each chunk In chunks.OrderBy(Function(item) item.Anchor)
                                            pageLines.AddRange(chunk.Painted)
                                            translated.Append(chunk.Output)
                                        Next
                                        If chunks.Any(Function(c) c.NoteText IsNot Nothing) Then
                                            WriteMarginNotes(output.Pages(number - 1), page, chunks, pageLines, translated, endNotes, cancel, trace)
                                        End If
                                    End If
                                    expected.Add(CanonicalText(translated.ToString()))
                                Next
                                If expected.All(Function(t) t.Length = 0) Then Throw New System.IO.InvalidDataException("No translatable PDF text was found.")
                                cancel()
                                If endNotes.Count > 0 Then WriteEndNotes(output, endNotes, expected, cancel)
                                output.Save(temporary)
                            End Using
                        End Using
                    End Using
                End Using
                trace("verify output")
                report("Checking translated text and page geometry")
                Using check = UglyToad.PdfPig.PdfDocument.Open(temporary), inputCheck = UglyToad.PdfPig.PdfDocument.Open(source)
                    If check.NumberOfPages <> expected.Count Then Throw New System.IO.InvalidDataException("Output page count differs.")
                    For number As System.Int32 = 1 To check.NumberOfPages
                        cancel()
                        If CanonicalText(check.GetPage(number).Text) <> expected(number - 1) Then Throw New System.IO.InvalidDataException($"Page {number}: output text failed readback.")
                    Next
                    completeTextVerified = True
                    For number As System.Int32 = 1 To inputCheck.NumberOfPages
                        cancel()
                        Dim verified = check.GetPage(number)
                        Dim inputPage = inputCheck.GetPage(number)
                        VerifyPageBoxes(inputPage, verified, $"Page {number} in translated PDF", trace)
                        Dim oldLinks = inputPage.GetHyperlinks()
                        Dim newLinks = verified.GetHyperlinks()
                        If oldLinks.Count <> newLinks.Count Then Throw New System.IO.InvalidDataException($"Page {number}: link count changed.")
                        For k As System.Int32 = 0 To oldLinks.Count - 1
                            If oldLinks(k).Uri <> newLinks(k).Uri Then Throw New System.IO.InvalidDataException($"Page {number}: link target changed.")
                            VerifyRectangle(oldLinks(k).Bounds, newLinks(k).Bounds, $"Page {number} link {k + 1}", trace)
                        Next
                        Dim lineIndex As System.Int32 = 0
                        Dim chars As System.Int32 = 0
                        Dim obstacles As System.Collections.Generic.List(Of UglyToad.PdfPig.Core.PdfRectangle) = Nothing
                        Dim acceptedMargins As System.Int32 = 0
                        Dim lastGraphicLine As System.Int32 = -1
                        For Each glyph In verified.Letters
                            Dim visible = CanonicalText(glyph.Value).Length
                            If visible = 0 Then Continue For
                            Dim bounds = positions(number - 1)(lineIndex)
                            Dim box = glyph.GlyphRectangle
                            If bounds.LinkIndex >= 0 Then
                                Dim linkBox = newLinks(bounds.LinkIndex).Bounds
                                Dim cx = (box.Left + box.Right) / 2
                                Dim cy = (box.Top + box.Bottom) / 2
                                If cx < linkBox.Left OrElse cx > linkBox.Right OrElse cy < linkBox.Bottom OrElse cy > linkBox.Top Then Throw New System.IO.InvalidDataException($"Page {number}: translated label falls outside its link.")
                            End If
                            Dim actual = System.Drawing.RectangleF.FromLTRB(CSng(box.Left), CSng(box.Bottom), CSng(box.Right), CSng(box.Top))
                            bounds.TranslatedRectangle = If(chars = 0, actual, System.Drawing.RectangleF.Union(bounds.TranslatedRectangle, actual))
                            If box.Left < bounds.Left OrElse box.Right > bounds.Right OrElse box.Bottom < bounds.Bottom OrElse box.Top > bounds.Top Then
                                If Not WithinInkMargin(bounds, glyph, verified.CropBox.Bounds) Then
                                    Dim detail = System.FormattableString.Invariant($"{bounds.Passage}, rendered line {lineIndex + 1}: glyph U+{System.Char.ConvertToUtf32(glyph.Value, 0):X4} ({glyph.FontName}) outside the permitted area. Overrun left/right/bottom/top (pt): {System.Math.Max(0, bounds.Left - box.Left):F2}/{System.Math.Max(0, box.Right - bounds.Right):F2}/{System.Math.Max(0, bounds.Bottom - box.Bottom):F2}/{System.Math.Max(0, box.Top - bounds.Top):F2}. Planned/actual baseline: {bounds.Baseline:F2}/{glyph.StartBaseLine.Y:F2}. Glyph bounds: [{box.Left:F2}, {box.Bottom:F2}, {box.Right:F2}, {box.Top:F2}]. Line bounds: [{bounds.Left:F2}, {bounds.Bottom:F2}, {bounds.Right:F2}, {bounds.Top:F2}].")
                                    trace(detail)
                                    Throw New System.IO.InvalidDataException(detail)
                                End If
                                If obstacles Is Nothing Then obstacles = GraphicBounds(inputPage)
                                If lastGraphicLine <> lineIndex AndAlso TouchesGraphicInMargin(bounds, box, obstacles) Then
                                    review(number, $"{bounds.Passage}, rendered line {lineIndex + 1}: minor glyph overhang near a painted graphic.")
                                    lastGraphicLine = lineIndex
                                End If
                                acceptedMargins += 1
                            End If
                            chars += visible
                            If chars >= bounds.Count Then
                                lineIndex += 1
                                chars = 0
                            End If
                        Next
                        If acceptedMargins > 0 Then trace($"Page {number}: {acceptedMargins} glyphs used the additional 1 mm ink margin.")
                        CheckOverlaps(positions(number - 1), number, review)
                    Next
                End Using
                cancel()
                trace("compare rendering")
                report("Checking rendered pages")
                Dim visualStatus = Await VerifyPagePixels(source, temporary, cancel, review, trace)
                cancel()
                If replaceExisting AndAlso System.IO.File.Exists(destination) Then
                    System.IO.File.Replace(temporary, destination, Nothing)
                Else
                    System.IO.File.Move(temporary, destination)
                End If
                Dim visualCheckIncomplete = visualStatus.StartsWith("Visual check incomplete", System.StringComparison.Ordinal)
                retainLog = (notePages.Count > 0 OrElse reviewPages.Count > 0 OrElse visualCheckIncomplete) AndAlso writer IsNot Nothing
                trace($"Saved {expected.Count} pages. Model calls: {modelCalls} ({modelMilliseconds / 1000.0:F1} s). {visualStatus} Font requests: " & System.String.Join(", ", substitutions))
                If retainLog Then System.Diagnostics.Debug.WriteLine("PDF translation diagnostic log: " & logPath)

                Dim notices As New System.Collections.Generic.List(Of System.String)()
                If reviewPages.Count > 0 Then notices.Add("Please check the layout on " & If(reviewPages.Count = 1, "page ", "pages ") & System.String.Join(", ", reviewPages) & ".")
                If visualCheckIncomplete Then notices.Add("Please check the appearance of the PDF. The automatic visual check could not finish.")
                If notePages.Count > 0 Then notices.Add("Numbered references on " & If(notePages.Count = 1, "page ", "pages ") & System.String.Join(", ", notePages) &
                    If(endNotes.Count = 0, " give the full translation in the page margin.", " refer to translation notes in the margin or at the end of the PDF."))
                If imagePages.Count > 0 Then notices.Add("Any text inside images was left in the original language on " & If(imagePages.Count = 1, "page ", "pages ") & System.String.Join(", ", imagePages) & ".")
                Return System.String.Join(" ", notices)
            Catch ex As System.OperationCanceledException
                Throw
            Catch ex As System.Exception
                retainLog = writer IsNot Nothing
                Dim detail = ex.Message
                If completeTextVerified AndAlso TypeOf ex Is System.IO.InvalidDataException AndAlso System.IO.File.Exists(temporary) Then
                    Dim reviewPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(temporary), "RedInk-review-" & System.IO.Path.GetFileName(temporary))
                    Try
                        System.IO.File.Move(temporary, reviewPath)
                        detail &= System.Environment.NewLine & "Translated PDF for layout review: " & reviewPath
                    Catch moveError As System.IO.IOException
                        detail &= System.Environment.NewLine & "Could not save the review copy: " & moveError.Message
                    Catch moveError As System.UnauthorizedAccessException
                        detail &= System.Environment.NewLine & "Could not save the review copy: " & moveError.Message
                    End Try
                End If
                trace("failure type=" & ex.GetType().FullName & " " & detail & System.Environment.NewLine & ex.StackTrace)
                If retainLog Then detail &= System.Environment.NewLine & "Diagnostic log: " & logPath
                Throw New System.IO.InvalidDataException(detail, ex)
            Finally
                Try
                    If writer IsNot Nothing Then writer.Dispose()
                    If Not retainLog AndAlso System.IO.File.Exists(logPath) Then System.IO.File.Delete(logPath)
                Catch ex As System.IO.IOException
                    System.Diagnostics.Debug.WriteLine(ex.Message)
                Catch ex As System.UnauthorizedAccessException
                    System.Diagnostics.Debug.WriteLine(ex.Message)
                End Try
                If System.IO.File.Exists(temporary) Then System.IO.File.Delete(temporary)
            End Try
        End Function
    End Class
End Namespace
