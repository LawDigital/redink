Option Explicit On
Option Strict On
Option Infer On

Namespace Agents
    Partial Public NotInheritable Class PdfFileTranslation
        ' Keep input operands byte-for-byte. Latin-1 is used only for token boundaries:
        ' every character offset is the corresponding byte offset in the decoded stream.
        Private Class OperationSource
            Private ReadOnly page As UglyToad.PdfPig.Content.Page
            Private ReadOnly data As System.Byte()
            Private ReadOnly ends As New System.Collections.Generic.List(Of System.Int32)()
            Public ReadOnly Mode As System.String

            Public Sub New(input As PdfSharp.Pdf.PdfPage, parsed As UglyToad.PdfPig.Content.Page)
                page = parsed
                Using joined As New System.IO.MemoryStream()
                    Dim streams = input.Contents.Elements
                    For i As System.Int32 = 0 To streams.Count - 1
                        Dim content = streams.GetDictionary(i).Stream.UnfilteredValue
                        joined.Write(content, 0, content.Length)
                        joined.WriteByte(10)
                    Next
                    data = joined.ToArray()
                End Using
                Dim text = System.Text.Encoding.GetEncoding(28591).GetString(data)
                Dim offset As System.Int32 = 0
                Dim depth As System.Int32 = 0
                While offset < text.Length
                    Dim start = offset
                    Dim ch = text(offset)
                    offset += 1
                    Select Case ch
                        Case ChrW(0), ChrW(9), ChrW(10), ChrW(12), ChrW(13), " "c
                        Case "%"c
                            While offset < text.Length AndAlso text(offset) <> ChrW(10) AndAlso text(offset) <> ChrW(13)
                                offset += 1
                            End While
                        Case "("c
                            Dim nesting As System.Int32 = 1
                            While nesting > 0 AndAlso offset < text.Length
                                ch = text(offset)
                                offset += 1
                                Select Case ch
                                    Case "\"c
                                        If offset < text.Length Then offset += 1
                                    Case "("c
                                        nesting += 1
                                    Case ")"c
                                        nesting -= 1
                                End Select
                            End While
                            If nesting <> 0 Then Exit While
                        Case "<"c
                            If offset < text.Length AndAlso text(offset) = "<"c Then
                                depth += 1
                                offset += 1
                            Else
                                Dim closing = text.IndexOf(">"c, offset)
                                If closing < 0 Then Exit While
                                offset = closing + 1
                            End If
                        Case ">"c
                            If offset >= text.Length OrElse text(offset) <> ">"c Then Exit While
                            offset += 1
                            depth -= 1
                        Case "["c
                            depth += 1
                        Case "]"c
                            depth -= 1
                        Case "/"c
                            While offset < text.Length AndAlso Not Delimiter(text(offset))
                                offset += 1
                            End While
                        Case Else
                            While offset < text.Length AndAlso Not Delimiter(text(offset))
                                offset += 1
                            End While
                            If depth <> 0 Then Continue While
                            Dim token = text.Substring(start, offset - start)
                            If token = "true" OrElse token = "false" OrElse token = "null" OrElse "+-.0123456789".IndexOf(ch) >= 0 Then Continue While
                            If token = "BI" OrElse ends.Count >= page.Operations.Count OrElse token <> page.Operations(ends.Count).[Operator] Then Exit While
                            ends.Add(offset)
                    End Select
                    If depth < 0 Then Exit While
                End While
                If depth = 0 AndAlso offset >= text.Length AndAlso ends.Count = page.Operations.Count Then
                    Mode = "original-bytes"
                Else
                    ' Inline image data has filter-specific boundaries. Keep using the PDF parser for these pages.
                    Mode = "parser-replay"
                End If
            End Sub

            Private Shared Function Delimiter(ch As System.Char) As System.Boolean
                Return ch = ChrW(0) OrElse ch = ChrW(9) OrElse ch = ChrW(10) OrElse ch = ChrW(12) OrElse ch = ChrW(13) OrElse " ()<>[]{}/%".IndexOf(ch) >= 0
            End Function

            Public Sub Write(index As System.Int32, output As System.IO.Stream)
                If Mode = "original-bytes" Then
                    Dim start = If(index = 0, 0, ends(index - 1))
                    output.Write(data, start, ends(index) - start)
                Else
                    page.Operations(index).Write(output)
                End If
                output.WriteByte(10)
            End Sub
        End Class
    End Class
End Namespace
