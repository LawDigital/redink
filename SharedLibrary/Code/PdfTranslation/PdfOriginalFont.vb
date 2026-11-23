Option Explicit On
Option Strict On
Option Infer On

Imports System.Linq

Namespace Agents
    Partial Public NotInheritable Class PdfFileTranslation
        Private Class SourceEncoding
            Public Resource As System.String
            Public FontName As System.String
            Public Codes As New System.Collections.Generic.Dictionary(Of System.Char, System.String)()
            Public Advances As New System.Collections.Generic.Dictionary(Of System.Char, System.Double)()
            Public Ambiguous As New System.Collections.Generic.HashSet(Of System.Char)()
        End Class

        ' Decode only the literal and hexadecimal string operands of Tj/TJ. Unknown
        ' representations disable original-font reuse for this object; the PDF parser
        ' remains authoritative for character mapping and all other operators.
        Private Shared Function OperandBytes(raw As System.Byte(), result As System.Collections.Generic.List(Of System.Byte)) As System.Boolean
            Dim i As System.Int32 = 0
            While i < raw.Length
                Dim c = raw(i)
                i += 1
                If c = 37 Then
                    While i < raw.Length AndAlso raw(i) <> 10 AndAlso raw(i) <> 13
                        i += 1
                    End While
                ElseIf c = 60 Then
                    Dim hex As New System.Text.StringBuilder()
                    While i < raw.Length AndAlso raw(i) <> 62
                        If Not System.Char.IsWhiteSpace(ChrW(raw(i))) Then hex.Append(ChrW(raw(i)))
                        i += 1
                    End While
                    If i >= raw.Length Then Return False
                    i += 1
                    If hex.Length Mod 2 = 1 Then hex.Append("0"c)
                    For j As System.Int32 = 0 To hex.Length - 1 Step 2
                        Dim value As System.Byte
                        If Not System.Byte.TryParse(hex.ToString(j, 2), System.Globalization.NumberStyles.HexNumber, Invariant, value) Then Return False
                        result.Add(value)
                    Next
                ElseIf c = 40 Then
                    Dim depth As System.Int32 = 1
                    While depth > 0 AndAlso i < raw.Length
                        c = raw(i)
                        i += 1
                        Select Case c
                            Case 40
                                depth += 1
                            Case 41
                                depth -= 1
                                If depth = 0 Then Exit While
                            Case 92
                                If i >= raw.Length Then Return False
                                c = raw(i)
                                i += 1
                                Select Case c
                                    Case 48 To 55
                                        Dim octal As System.Int32 = c - 48
                                        For k As System.Int32 = 1 To 2
                                            If i >= raw.Length OrElse raw(i) < 48 OrElse raw(i) > 55 Then Exit For
                                            octal = octal * 8 + raw(i) - 48
                                            i += 1
                                        Next
                                        c = CByte(octal And 255)
                                    Case 110
                                        c = 10
                                    Case 114
                                        c = 13
                                    Case 116
                                        c = 9
                                    Case 98
                                        c = 8
                                    Case 102
                                        c = 12
                                    Case 10, 13
                                        If c = 13 AndAlso i < raw.Length AndAlso raw(i) = 10 Then i += 1
                                        Continue While
                                End Select
                        End Select
                        result.Add(c)
                    End While
                    If depth <> 0 Then Return False
                End If
            End While
            Return True
        End Function

        Private Shared Function OriginalFonts(page As UglyToad.PdfPig.Content.Page, raw As OperationSource,
                ids As System.Collections.Generic.Dictionary(Of UglyToad.PdfPig.Content.Letter, System.Int32)) As System.Collections.Generic.Dictionary(Of System.Int32, SourceEncoding)
            Dim objectFonts As New System.Collections.Generic.Dictionary(Of System.Int32, SourceEncoding)()
            Dim resources As New System.Collections.Generic.Dictionary(Of System.String, SourceEncoding)(System.StringComparer.Ordinal)
            Dim lettersById = page.Letters.GroupBy(Function(g) ids(g)).ToDictionary(Function(g) g.Key, Function(g) g.ToList())
            Dim id As System.Int32 = 0
            Dim resource As System.String = Nothing
            Dim activeFont As System.String = Nothing
            Dim savedFonts As New System.Collections.Generic.Stack(Of System.String)()
            Dim usable As System.Boolean = False
            Dim bytes As New System.Collections.Generic.List(Of System.Byte)()
            For k As System.Int32 = 0 To page.Operations.Count - 1
                Dim op = page.Operations(k)
                Select Case op.[Operator]
                    Case "q"
                        savedFonts.Push(activeFont)
                    Case "Q"
                        activeFont = If(savedFonts.Count = 0, Nothing, savedFonts.Pop())
                    Case "gs"
                        activeFont = Nothing
                        usable = False
                    Case "BT"
                        id += 1
                        resource = activeFont
                        usable = True
                        bytes.Clear()
                    Case "Tf"
                        Using encoded As New System.IO.MemoryStream()
                            op.Write(encoded)
                            Dim token = System.Text.Encoding.ASCII.GetString(encoded.ToArray()).Trim().Split(" "c)(0)
                            If Not token.StartsWith("/", System.StringComparison.Ordinal) OrElse (bytes.Count > 0 AndAlso token <> resource) Then usable = False
                            activeFont = token
                            resource = token
                        End Using
                    Case "'", """"
                        usable = False
                    Case "TJ", "Tj"
                        If resource Is Nothing Then usable = False
                        If Not usable Then Continue For
                        Using encoded As New System.IO.MemoryStream()
                            raw.Write(k, encoded)
                            usable = OperandBytes(encoded.ToArray(), bytes)
                        End Using
                    Case "ET"
                        If Not usable OrElse resource Is Nothing OrElse Not lettersById.ContainsKey(id) Then Continue For
                        Dim glyphs = lettersById(id)
                        If bytes.Count <> glyphs.Count AndAlso bytes.Count <> 2 * glyphs.Count Then Continue For
                        If glyphs.Any(Function(g) g.Value.Length <> 1 OrElse g.PointSize <= 0 OrElse g.FontName <> glyphs(0).FontName) Then Continue For
                        Dim enc As SourceEncoding = Nothing
                        If Not resources.TryGetValue(resource, enc) Then
                            enc = New SourceEncoding With {.Resource = resource, .FontName = glyphs(0).FontName}
                            resources.Add(resource, enc)
                        End If
                        Dim stride = bytes.Count \ glyphs.Count
                        For j As System.Int32 = 0 To glyphs.Count - 1
                            Dim ch = glyphs(j).Value(0)
                            Dim code = System.BitConverter.ToString(bytes.GetRange(j * stride, stride).ToArray()).Replace("-", "")
                            If enc.Codes.ContainsKey(ch) AndAlso enc.Codes(ch) <> code Then enc.Ambiguous.Add(ch)
                            enc.Codes(ch) = code
                            enc.Advances(ch) = (glyphs(j).EndBaseLine.X - glyphs(j).StartBaseLine.X) / glyphs(j).PointSize
                        Next
                        objectFonts(id) = enc
                End Select
            Next
            Return objectFonts
        End Function

        Private Shared Function OriginalTextCommands(enc As SourceEncoding, text As System.String, x As System.Double,
                y As System.Double, size As System.Double, horizontalScale As System.Double, color As (System.Double, System.Double, System.Double)) As System.String
            Dim numbers = {color.Item1, color.Item2, color.Item3, size, x, y}.Select(Function(v) v.ToString("0.########", Invariant)).ToArray()
            Dim hex = System.String.Concat(text.Select(Function(ch) enc.Codes(ch)))
            Return "q " & System.String.Join(" ", numbers.Take(3)) & " rg BT " & enc.Resource & " " & numbers(3) & " Tf " & horizontalScale.ToString("0.##", Invariant) & " 0 0 1 " &
                numbers(4) & " " & numbers(5) & " Tm <" & hex & "> Tj ET Q" & vbLf
        End Function
    End Class
End Namespace
