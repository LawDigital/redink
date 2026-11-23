Option Explicit On
Option Strict On
Option Infer On

Imports System.Linq

Namespace Agents
    Partial Public NotInheritable Class PdfFileTranslation
        Private Shared ReadOnly Invariant As System.Globalization.CultureInfo = System.Globalization.CultureInfo.InvariantCulture

        Private Shared Sub Ascii(stream As System.IO.Stream, value As System.String)
            Dim bytes = System.Text.Encoding.ASCII.GetBytes(value & vbLf)
            stream.Write(bytes, 0, bytes.Length)
        End Sub

        Private Shared Function ShowOperator(name As System.String) As System.Boolean
            Return name = "Tj" OrElse name = "TJ" OrElse name = "'" OrElse name = """"
        End Function

        Private Shared Sub SetContents(doc As PdfSharp.Pdf.PdfDocument, page As PdfSharp.Pdf.PdfPage, data As System.Byte())
            Dim content As New PdfSharp.Pdf.PdfDictionary(doc)
            content.CreateStream(data)
            doc.Internals.AddObject(content)
            Dim entries = page.Contents.Elements
            entries.Clear()
            entries.Add(content.Reference)
        End Sub

        ' Color-code the text objects in a private copy. PdfPig then supplies the original
        ' operator location for every glyph, including custom encodings and TJ arrays.
        Private Shared Function MarkTextObjects(path As System.String, original As UglyToad.PdfPig.PdfDocument, cancel As System.Action, sources As System.Collections.Generic.List(Of OperationSource), trace As System.Action(Of System.String)) As System.Byte()
            Using marked = PdfSharp.Pdf.IO.PdfReader.Open(path, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)
                For n As System.Int32 = 1 To original.NumberOfPages
                    cancel()
                    Dim page = original.GetPage(n)
                    Dim source As New OperationSource(marked.Pages(n - 1), page)
                    sources.Add(source)
                    trace($"probe page={n} mode={source.Mode} operations={page.Operations.Count} glyphs={page.Letters.Count}")
                    If page.Letters.Count = 0 Then Continue For
                    Using content As New System.IO.MemoryStream()
                        Dim id As System.Int32 = 0
                        For k As System.Int32 = 0 To page.Operations.Count - 1
                            Dim op = page.Operations(k)
                            If op.[Operator] = "BT" Then id += 1
                            If ShowOperator(op.[Operator]) Then
                                If id > &HFFFFFF Then Throw New System.IO.InvalidDataException("Too many PDF text objects.")
                                Dim rgb = {id >> 16, (id >> 8) And 255, id And 255}
                                Ascii(content, System.String.Join(" ", rgb.Select(Function(v) (v / 255.0).ToString("0.000000000", Invariant))) & " rg")
                            End If
                            source.Write(k, content)
                        Next
                        SetContents(marked, marked.Pages(n - 1), content.ToArray())
                    End Using
                Next
                Using bytes As New System.IO.MemoryStream()
                    marked.Save(bytes, False)
                    Return bytes.ToArray()
                End Using
            End Using
        End Function

        Private Shared Function ObjectIds(source As UglyToad.PdfPig.Content.Page, marked As UglyToad.PdfPig.Content.Page, trace As System.Action(Of System.String)) As System.Collections.Generic.Dictionary(Of UglyToad.PdfPig.Content.Letter, System.Int32)
            Dim ids As New System.Collections.Generic.Dictionary(Of UglyToad.PdfPig.Content.Letter, System.Int32)()
            Dim beginOps = source.Operations.Select(Function(op, index) If(op.[Operator] = "BT", index + 1, 0)).Where(Function(index) index > 0).ToArray()
            Dim count = System.Math.Min(source.Letters.Count, marked.Letters.Count)
            Dim summary = $"page={source.Number} originalGlyphs={source.Letters.Count} probeGlyphs={marked.Letters.Count}"
            trace("map " & summary)
            For index As System.Int32 = 0 To count - 1
                Dim letter = source.Letters(index)
                Dim other = marked.Letters(index)
                Dim color = other.FillColor.ToRGBValues()
                If letter.Value <> other.Value OrElse Not letter.StartBaseLine.Equals(other.StartBaseLine) Then
                    Dim objectId = CInt(color.Item1 * 255) * 65536 + CInt(color.Item2 * 255) * 256 + CInt(color.Item3 * 255)
                    Dim operatorIndex = If(objectId > 0 AndAlso objectId <= beginOps.Length, beginOps(objectId - 1), 0)
                    trace($"mismatch {summary} beginOperator={operatorIndex} glyph={index + 1} probeBT={objectId} sourceSequence={letter.TextSequence} probeSequence={other.TextSequence} originalXY={letter.StartBaseLine} probeXY={other.StartBaseLine} valueEqual={letter.Value = other.Value}")
                    Throw New System.IO.InvalidDataException($"Page {source.Number}: original and probe glyphs diverge at glyph {index + 1}.")
                End If
                Dim id = CInt(color.Item1 * 255) * 65536 + CInt(color.Item2 * 255) * 256 + CInt(color.Item3 * 255)
                If id < 1 OrElse id > beginOps.Length Then
                    trace($"invalid marker {summary} glyph={index + 1} marker={id} objects={beginOps.Length}")
                    Throw New System.IO.InvalidDataException($"Page {source.Number}, glyph {index + 1}: text outside a mapped PDF text object.")
                End If
                ids.Add(letter, id)
            Next
            If source.Letters.Count <> marked.Letters.Count Then
                Dim unpaired = If(source.Letters.Count > count, source.Letters(count), marked.Letters(count))
                trace($"mismatch {summary} firstUnpairedGlyph={count + 1} sequence={unpaired.TextSequence} position={unpaired.StartBaseLine}")
                Throw New System.IO.InvalidDataException($"Page {source.Number}: PDF glyph count changed ({source.Letters.Count} original, {marked.Letters.Count} probe).")
            End If
            Return ids
        End Function

        ' Text objects may share a paragraph until an intervening paint or graphics-state
        ' change makes moving content between them unsafe.
        Private Shared Function PaintGroups(page As UglyToad.PdfPig.Content.Page) As System.Collections.Generic.List(Of System.Int32)
            Dim groups As New System.Collections.Generic.List(Of System.Int32) From {0}
            Dim epoch As System.Int32 = 0
            For Each op In page.Operations
                Select Case op.[Operator]
                    Case "BT"
                        groups.Add(epoch)
                    Case "S", "s", "f", "F", "f*", "B", "B*", "b", "b*", "sh", "Do", "BI", "EI", "W", "W*", "gs", "q", "Q", "cm", "BDC", "BMC", "EMC"
                        epoch += 1
                End Select
            Next
            Return groups
        End Function

        Private Shared Sub InsertAtOriginalDepth(document As PdfSharp.Pdf.PdfDocument, target As PdfSharp.Pdf.PdfPage,
                source As UglyToad.PdfPig.Content.Page, raw As OperationSource, chunks As System.Collections.Generic.List(Of TranslationChunk))
            Dim keep As New System.Collections.Generic.HashSet(Of System.Int32)(chunks.Where(Function(c) c.RetainOriginal).Select(Function(c) c.Anchor))
            Dim byObject = chunks.Where(Function(c) Not c.RetainOriginal).GroupBy(Function(c) c.Anchor).ToDictionary(Function(g) g.Key)
            Dim stack As New System.Collections.Generic.Stack(Of System.Drawing.Drawing2D.Matrix)()
            Dim transform As New System.Drawing.Drawing2D.Matrix()
            Dim id As System.Int32 = 0
            Try
                Using stream As New System.IO.MemoryStream()
                    For k As System.Int32 = 0 To source.Operations.Count - 1
                        Dim op = source.Operations(k)
                        Dim symbol = op.[Operator]
                        Select Case symbol
                            Case "BT"
                                id += 1
                            Case "cm"
                                Using encoded As New System.IO.MemoryStream()
                                    op.Write(encoded)
                                    Dim parts = System.Text.RegularExpressions.Regex.Split(System.Text.Encoding.ASCII.GetString(encoded.ToArray()).Trim(), "\s+")
                                    Dim m = parts.Take(6).Select(Function(s) System.Single.Parse(s, Invariant)).ToArray()
                                    Using delta As New System.Drawing.Drawing2D.Matrix(m(0), m(1), m(2), m(3), m(4), m(5))
                                        transform.Multiply(delta, System.Drawing.Drawing2D.MatrixOrder.Prepend)
                                    End Using
                                End Using
                            Case "q"
                                stack.Push(transform.Clone())
                            Case "Q"
                                If stack.Count = 0 Then Throw New System.IO.InvalidDataException("Unbalanced PDF graphics state.")
                                transform.Dispose()
                                transform = stack.Pop()
                        End Select
                        If Not ShowOperator(symbol) OrElse keep.Contains(id) Then
                            raw.Write(k, stream)
                        ElseIf symbol = """" Then
                            ' The double quote operator also changes persistent word/character spacing.
                            Using encoded As New System.IO.MemoryStream()
                                op.Write(encoded)
                                Dim values = System.Text.RegularExpressions.Regex.Split(System.Text.Encoding.ASCII.GetString(encoded.ToArray()).TrimStart(), "\s+", System.Text.RegularExpressions.RegexOptions.None)
                                Ascii(stream, values(0) & " Tw " & values(1) & " Tc")
                            End Using
                        End If
                        If symbol <> "ET" OrElse Not byObject.ContainsKey(id) Then Continue For
                        Using inverse = transform.Clone()
                            inverse.Invert()
                            Ascii(stream, "q " & System.String.Join(" ", inverse.Elements.Select(Function(v) v.ToString("0.#########", Invariant))) & " cm 0 g")
                            ' Normalize inherited text state for the page-coordinate commands from PDFsharp.
                            Ascii(stream, "BT 0 Tc 0 Tw 100 Tz 0 Ts 0 Tr ET")
                            For Each chunk In byObject(id)
                                Ascii(stream, "q BT " & CInt(chunk.Lines(0).Letters(0).RenderingMode).ToString(Invariant) & " Tr ET")
                                stream.Write(chunk.Commands, 0, chunk.Commands.Length)
                                stream.WriteByte(10)
                                Ascii(stream, "Q")
                            Next
                            Ascii(stream, "Q")
                        End Using
                    Next
                    If stack.Count <> 0 Then Throw New System.IO.InvalidDataException("Unbalanced PDF graphics state.")
                    SetContents(document, target, stream.ToArray())
                End Using
            Finally
                transform.Dispose()
                For Each saved In stack
                    saved.Dispose()
                Next
            End Try
        End Sub
    End Class
End Namespace
