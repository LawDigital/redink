' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.WebExtension.FileHelpers.vb
' Purpose: Plaintext extraction helpers for Office documents (Word, Excel, PowerPoint)
'          and text/code-like files. Provides safe COM release utilities, XML
'          escaping, column letter conversion, and encoding-smart file reading.
'
' Architecture:
'   - Entry points: TryExtractOfficeText (Word/Excel/PowerPoint), TryExtractTextLike (text/code files).
'   - Word/Excel: early-bound Microsoft Office Interop with explicit COM release (SafeCloseWord/SafeCloseExcel).
'   - PowerPoint: late binding (CreateObject) to avoid PIA dependency; manual COM release.
'   - Excel extraction: attempts bulk 2D array retrieval (Value2 and Formula) then per-cell fallback.
'   - PowerPoint extraction: iterates slides, shapes, optional notes.
'   - Text-like extraction: extension filter, optional RTF via Word, otherwise direct file read.
'   - ReadAllTextSmart: UTF-8 with BOM detection, fallback Windows-1252, final generic read.
'   - Defensive error handling: returns False on failure, truncates very large outputs (1,500,000 chars).
' =============================================================================

Option Explicit On
Option Strict Off

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Retrieves content from a web URL and stores it as a temporary file.
    ''' Documents are downloaded as original files; websites are retrieved via the shared WebView2 importer and stored as text.
    ''' </summary>
    Public Async Function CreateTempFileFromUrlAsync(url As System.String) As System.Threading.Tasks.Task(Of System.String)
        Return Await SharedLibrary.SharedLibrary.SharedMethods.CreateTempFileFromUrlAsync(
            url,
            SharedLibrary.SharedLibrary.SharedMethods.GetDefaultWebImportSupportedExtensions(INI_AllowLegacyDocFiles, True))
    End Function


    ''' <summary>
    ''' Attempts plaintext extraction for Word (.doc/.docx/.rtf), Excel (.xls/.xlsx), or PowerPoint (.ppt/.pptx),
    ''' plus email files (.msg/.eml). Prefers sandboxed (COM-free) readers for OpenXML formats;
    ''' falls back to COM Interop for legacy binary formats (.doc, .xls, .ppt, .rtf).
    ''' Truncates output at 1,500,000 characters.
    ''' </summary>
    ''' <param name="filePath">Absolute file path.</param>
    ''' <param name="extracted">Plaintext result (output).</param>
    ''' <param name="label">Descriptive label (output).</param>
    ''' <returns>True if extraction succeeded; otherwise False.</returns>
    Private Function TryExtractOfficeText(
    ByVal filePath As System.String,
    ByRef extracted As System.String,
    ByRef label As System.String
) As System.Boolean

        extracted = Nothing
        label = Nothing

        If System.String.IsNullOrWhiteSpace(filePath) Then Return False
        If Not System.IO.File.Exists(filePath) Then Return False

        Dim ext As System.String = System.IO.Path.GetExtension(filePath).ToLowerInvariant()

        Try
            Select Case ext
                ' ── Sandboxed readers (no COM) ──
                Case ".docx"
                    extracted = SharedLibrary.SharedLibrary.SharedMethods.ReadDocxSandboxed(filePath)
                    label = "Word document: " & System.IO.Path.GetFileName(filePath)
                Case ".xlsx"
                    extracted = SharedLibrary.SharedLibrary.SharedMethods.ReadXlsxSandboxed(filePath)
                    label = "Excel workbook: " & System.IO.Path.GetFileName(filePath)
                Case ".pptx"
                    extracted = SharedLibrary.SharedLibrary.SharedMethods.ReadPptxSandboxed(filePath)
                    label = "PowerPoint presentation: " & System.IO.Path.GetFileName(filePath)
                Case ".eml"
                    extracted = SharedLibrary.SharedLibrary.SharedMethods.ReadEmlSandboxed(filePath)
                    label = "Email message: " & System.IO.Path.GetFileName(filePath)
                Case ".msg"
                    extracted = SharedLibrary.SharedLibrary.SharedMethods.ReadMsgSandboxed(filePath)
                    label = "Email message: " & System.IO.Path.GetFileName(filePath)

                ' ── COM fallback for legacy binary formats ──
                Case ".doc"
                    If Not INI_AllowLegacyDocFiles Then
                        extracted = Nothing
                        Return False
                    End If
                    extracted = ExtractWordText(filePath)
                    label = "Word document: " & System.IO.Path.GetFileName(filePath)
                Case ".rtf"
                    extracted = ExtractWordText(filePath)
                    label = "Word document (RTF): " & System.IO.Path.GetFileName(filePath)
                Case ".xls"
                    extracted = ExtractExcelText(filePath)
                    label = "Excel workbook: " & System.IO.Path.GetFileName(filePath)
                Case ".ppt"
                    extracted = ExtractPowerPointText(filePath)
                    label = "PowerPoint presentation: " & System.IO.Path.GetFileName(filePath)
                Case Else
                    Return False
            End Select
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("Office extract failed: " & ex.Message)
            extracted = Nothing
            label = Nothing
            Return False
        End Try

        ' Sandboxed readers return "Error:..." on failure — treat as not extracted
        If extracted IsNot Nothing AndAlso extracted.StartsWith("Error") AndAlso extracted.Length < 200 Then
            extracted = Nothing
            Return False
        End If

        If System.String.IsNullOrWhiteSpace(extracted) Then Return False
        If extracted.Length > 1_500_000 Then
            extracted = extracted.Substring(0, 1_500_000) & System.Environment.NewLine & "[…truncated…]"
        End If

        Return True
    End Function


    ''' <summary>
    ''' Returns full Word document text with normalized line breaks. Uses early-bound Word Interop.
    ''' </summary>
    ''' <param name="path">Absolute file path.</param>
    ''' <returns>Trimmed document text.</returns>
    Private Function ExtractWordText(ByVal path As System.String) As System.String
        Dim app As Microsoft.Office.Interop.Word.Application = Nothing
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim weOwnApp As System.Boolean = False
        Try
            ' Try to reuse an existing Word instance
            Try
                app = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application"),
                            Microsoft.Office.Interop.Word.Application)
            Catch ex As System.Runtime.InteropServices.COMException
                app = New Microsoft.Office.Interop.Word.Application()
                weOwnApp = True
            End Try

            app.Visible = False
            doc = app.Documents.Open(FileName:=path, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)

            ' Full document text - simple & robust
            Dim raw As System.String = doc.Content.Text

            ' Normalize line breaks
            raw = raw.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            raw = System.Text.RegularExpressions.Regex.Replace(raw, "[\f\v]+", vbLf)

            Return raw.Trim()
        Catch ex As System.Exception
            Throw
        Finally
            SafeCloseWord(doc, app, weOwnApp)
        End Try
    End Function

    ''' <summary>
    ''' Safely closes and releases Word Document and Application COM objects without saving changes.
    ''' </summary>
    ''' <param name="doc">Word document instance.</param>
    ''' <param name="app">Word application instance.</param>
    Private Sub SafeCloseWord(
    ByVal doc As Microsoft.Office.Interop.Word.Document,
    ByVal app As Microsoft.Office.Interop.Word.Application,
    ByVal weOwnApp As System.Boolean
)
        Try
            If doc IsNot Nothing Then
                Try : doc.Close(SaveChanges:=False) : Catch : End Try
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(doc) : Catch : End Try
            End If
        Finally
            If app IsNot Nothing Then
                ' Only quit if we created this instance
                If weOwnApp Then
                    Try : app.Quit(SaveChanges:=False) : Catch : End Try
                End If
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app) : Catch : End Try
            End If
        End Try
    End Sub


    ''' <summary>
    ''' Extracts worksheet content, listing each cell with address, formula and value.
    ''' Attempts bulk 2D array retrieval (Value2 and Formula) then per-cell fallback.
    ''' </summary>
    ''' <param name="path">Absolute workbook file path.</param>
    ''' <returns>Trimmed textual representation of workbook data.</returns>
    Private Function ExtractExcelText(ByVal path As System.String) As System.String
        Dim app As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim sb As New System.Text.StringBuilder(4096)
        Dim weOwnApp As System.Boolean = False

        Try
            ' Try to reuse an existing Excel instance
            Try
                app = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"),
                            Microsoft.Office.Interop.Excel.Application)
            Catch ex As System.Runtime.InteropServices.COMException
                app = New Microsoft.Office.Interop.Excel.Application()
                weOwnApp = True
            End Try

            app.Visible = False
            wb = app.Workbooks.Open(Filename:=path, ReadOnly:=True, AddToMru:=False)

            For Each shObj As System.Object In wb.Worksheets
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Try
                    ws = CType(shObj, Microsoft.Office.Interop.Excel.Worksheet)
                    Dim used As Microsoft.Office.Interop.Excel.Range = ws.UsedRange
                    If used Is Nothing Then Continue For

                    sb.AppendLine("=== Sheet: " & ws.Name & " ===")

                    Dim rows As System.Int32 = used.Rows.Count
                    Dim cols As System.Int32 = used.Columns.Count
                    Dim rowOffset As System.Int32 = used.Row      ' 1-based
                    Dim colOffset As System.Int32 = used.Column   ' 1-based

                    ' Fast path: retrieve both arrays at once
                    Dim dataValues As System.Object(,) = Nothing
                    Dim dataFormulas As System.Object(,) = Nothing
                    Try
                        dataValues = TryCast(used.Value2, System.Object(,))
                    Catch
                        dataValues = Nothing
                    End Try
                    Try
                        dataFormulas = TryCast(used.Formula, System.Object(,))
                    Catch
                        dataFormulas = Nothing
                    End Try

                    If dataValues IsNot Nothing AndAlso dataFormulas IsNot Nothing Then
                        Dim rL As System.Int32 = dataValues.GetLength(0)
                        Dim cL As System.Int32 = dataValues.GetLength(1)
                        For r As System.Int32 = 1 To rL
                            For c As System.Int32 = 1 To cL
                                Dim absRow As System.Int32 = rowOffset + r - 1
                                Dim absCol As System.Int32 = colOffset + c - 1
                                Dim addr As System.String = ColToLetters(absCol) & absRow.ToString(System.Globalization.CultureInfo.InvariantCulture)

                                Dim vObj As System.Object = dataValues(r, c)
                                Dim fObj As System.Object = dataFormulas(r, c)

                                Dim vStr As System.String = System.Convert.ToString(vObj, System.Globalization.CultureInfo.InvariantCulture)
                                Dim fStr As System.String = System.Convert.ToString(fObj, System.Globalization.CultureInfo.InvariantCulture)

                                ' Some constant cells have Formula="" or Nothing
                                If fObj IsNot Nothing Then
                                    ' Excel returns the value instead of a formula for constants.
                                    ' If the formula appears identical to the value (often empty), leave blank.
                                End If

                                sb.Append(addr)
                                sb.Append(vbTab)
                                sb.Append("FORMULA:")
                                If Not System.String.IsNullOrEmpty(fStr) Then
                                    sb.Append("="c)
                                    sb.Append(fStr.TrimStart("="c))
                                End If
                                sb.Append(vbTab)
                                sb.Append("VALUE: ")
                                sb.AppendLine(If(vStr, ""))
                            Next
                        Next
                    Else
                        ' Fallback: cell by cell (slower, but robust)
                        For r As System.Int32 = 1 To rows
                            For c As System.Int32 = 1 To cols
                                Dim cell As Microsoft.Office.Interop.Excel.Range = Nothing
                                Try
                                    cell = CType(used.Cells(r, c), Microsoft.Office.Interop.Excel.Range)

                                    Dim absRow As System.Int32 = rowOffset + r - 1
                                    Dim absCol As System.Int32 = colOffset + c - 1
                                    Dim addr As System.String = ColToLetters(absCol) & absRow.ToString(System.Globalization.CultureInfo.InvariantCulture)

                                    Dim vObj As System.Object = Nothing
                                    Dim fObj As System.Object = Nothing
                                    Try : vObj = cell.Value2 : Catch : vObj = Nothing : End Try
                                    Try : fObj = cell.Formula : Catch : fObj = Nothing : End Try

                                    Dim vStr As System.String = System.Convert.ToString(vObj, System.Globalization.CultureInfo.InvariantCulture)
                                    Dim fStr As System.String = System.Convert.ToString(fObj, System.Globalization.CultureInfo.InvariantCulture)

                                    sb.Append(addr)
                                    sb.Append(vbTab)
                                    sb.Append("FORMULA:")
                                    If Not System.String.IsNullOrEmpty(fStr) Then
                                        sb.Append("="c)
                                        sb.Append(fStr.TrimStart("="c))
                                    End If
                                    sb.Append(vbTab)
                                    sb.Append("VALUE: ")
                                    sb.AppendLine(If(vStr, ""))
                                Finally
                                    If cell IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cell)
                                End Try
                            Next
                        Next
                    End If

                    If used IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(used)
                    sb.AppendLine()
                Finally
                    If ws IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ws)
                End Try
            Next

            Return sb.ToString().Trim()
        Catch ex As System.Exception
            Throw
        Finally
            SafeCloseExcel(wb, app, weOwnApp)
        End Try
    End Function

    ' A1 column letters
    ''' <summary>
    ''' Converts 1-based column index to A1-style letters (e.g. 1=A, 27=AA).
    ''' </summary>
    ''' <param name="col">1-based column index.</param>
    ''' <returns>Column letters.</returns>
    Private Function ColToLetters(ByVal col As System.Int32) As System.String
        ' col: 1-based (1=A, 27=AA, …)
        Dim n As System.Int32 = col
        Dim chars As New System.Text.StringBuilder()
        While n > 0
            n -= 1
            Dim ch As System.Char = System.Convert.ToChar((n Mod 26) + System.Convert.ToInt32("A"c))
            chars.Insert(0, ch)
            n \= 26
        End While
        Return chars.ToString()
    End Function

    ''' <summary>
    ''' Safely closes and releases Excel Workbook and Application COM objects without saving changes.
    ''' </summary>
    ''' <param name="wb">Excel workbook instance.</param>
    ''' <param name="app">Excel application instance.</param>
    Private Sub SafeCloseExcel(
    ByVal wb As Microsoft.Office.Interop.Excel.Workbook,
    ByVal app As Microsoft.Office.Interop.Excel.Application,
    ByVal weOwnApp As System.Boolean
)
        Try
            If wb IsNot Nothing Then
                Try : wb.Close(SaveChanges:=False) : Catch : End Try
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wb) : Catch : End Try
            End If
        Finally
            If app IsNot Nothing Then
                ' Only quit if we created this instance
                If weOwnApp Then
                    Try : app.Quit() : Catch : End Try
                End If
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app) : Catch : End Try
            End If
        End Try
    End Sub

    '──────────────────────────────────────────────────────────────────────────────
    ' POWERPOINT
    '──────────────────────────────────────────────────────────────────────────────
    ''' <summary>
    ''' Extracts text content and notes from a PowerPoint presentation using late binding.
    ''' Iterates slides, shapes with text frames, and optional notes pages.
    ''' </summary>
    ''' <param name="path">Absolute presentation file path.</param>
    ''' <returns>Trimmed textual content.</returns>
    Private Function ExtractPowerPointText(ByVal path As System.String) As System.String
        Dim app As System.Object = Nothing
        Dim pres As System.Object = Nothing
        Dim sb As New System.Text.StringBuilder(2048)
        Dim weOwnApp As System.Boolean = False

        Try
            ' Late binding: no PIAs required
            ' Try to get an existing instance first
            Try
                app = System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application")
            Catch ex As System.Runtime.InteropServices.COMException
                ' No running instance – create a new one; we own it
                app = Microsoft.VisualBasic.Interaction.CreateObject("PowerPoint.Application")
                weOwnApp = True
            End Try

            ' Presentations.Open(FileName, ReadOnly, Untitled, WithWindow)
            ' Late bound: True/False as -1/0; here 1=True, 0=False
            Dim presentations As System.Object = app.Presentations
            pres = presentations.Open(path, 1, 0, 0)

            Dim slideCount As System.Int32 = System.Convert.ToInt32(pres.Slides.Count, System.Globalization.CultureInfo.InvariantCulture)
            For i As System.Int32 = 1 To slideCount
                Dim sld As System.Object = pres.Slides(i)
                Try
                    sb.AppendLine("=== Slide " & i.ToString(System.Globalization.CultureInfo.InvariantCulture) & " ===")

                    Dim shapeCount As System.Int32 = System.Convert.ToInt32(sld.Shapes.Count, System.Globalization.CultureInfo.InvariantCulture)
                    For j As System.Int32 = 1 To shapeCount
                        Dim shp As System.Object = sld.Shapes(j)
                        Try
                            Dim hasTf As System.Boolean = False
                            Try
                                ' In Office Interop: True = -1, False = 0
                                hasTf = (System.Convert.ToInt32(shp.HasTextFrame, System.Globalization.CultureInfo.InvariantCulture) <> 0) AndAlso
                                    (Not shp.TextFrame Is Nothing) AndAlso
                                    (System.Convert.ToInt32(shp.TextFrame.HasText, System.Globalization.CultureInfo.InvariantCulture) <> 0)
                            Catch
                                hasTf = False
                            End Try

                            If hasTf Then
                                Dim txt As System.String = System.Convert.ToString(shp.TextFrame.TextRange.Text, System.Globalization.CultureInfo.InvariantCulture)
                                If Not System.String.IsNullOrWhiteSpace(txt) Then
                                    sb.AppendLine(txt.Trim())
                                End If
                            End If
                        Finally
                            Try
                                If shp IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shp)
                            Catch
                            End Try
                        End Try
                    Next

                    ' Notes (optional)
                    Try
                        Dim notesShapes As System.Object = sld.NotesPage.Shapes
                        Dim nCount As System.Int32 = System.Convert.ToInt32(notesShapes.Count, System.Globalization.CultureInfo.InvariantCulture)
                        For k As System.Int32 = 1 To nCount
                            Dim shp2 As System.Object = notesShapes(k)
                            Try
                                Dim hasTf2 As System.Boolean = False
                                Try
                                    hasTf2 = (System.Convert.ToInt32(shp2.HasTextFrame, System.Globalization.CultureInfo.InvariantCulture) <> 0) AndAlso
                                         (Not shp2.TextFrame Is Nothing) AndAlso
                                         (System.Convert.ToInt32(shp2.TextFrame.HasText, System.Globalization.CultureInfo.InvariantCulture) <> 0)
                                Catch
                                    hasTf2 = False
                                End Try
                                If hasTf2 Then
                                    Dim note As System.String = System.Convert.ToString(shp2.TextFrame.TextRange.Text, System.Globalization.CultureInfo.InvariantCulture)
                                    If Not System.String.IsNullOrWhiteSpace(note) Then
                                        sb.AppendLine("--- Notes ---")
                                        sb.AppendLine(note.Trim())
                                    End If
                                End If
                            Finally
                                Try
                                    If shp2 IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shp2)
                                Catch
                                End Try
                            End Try
                        Next
                    Catch
                    End Try

                    sb.AppendLine()
                Finally
                    Try
                        If sld IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sld)
                    Catch
                    End Try
                End Try
            Next

            Return sb.ToString().Trim()
        Catch ex As System.Exception
            Throw
        Finally
            Try
                If pres IsNot Nothing Then
                    Try : pres.Close() : Catch : End Try
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pres) : Catch : End Try
                End If
            Catch
            End Try
            Try
                If app IsNot Nothing Then
                    ' Only quit if we created the instance ourselves
                    If weOwnApp Then
                        Try : app.Quit() : Catch : End Try
                    End If
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app) : Catch : End Try
                End If
            Catch
            End Try
        End Try
    End Function


    ''' <summary>
    ''' XML-escapes a string (returns empty string if input is Nothing).
    ''' </summary>
    ''' <param name="s">Input string.</param>
    ''' <returns>Escaped string.</returns>
    Private Function EscapeForXml(ByVal s As System.String) As System.String
        If s Is Nothing Then Return ""
        Return System.Security.SecurityElement.Escape(s)
    End Function

    ''' <summary>
    ''' Attempts extraction for text/code-like files based on extension list.
    ''' Optional RTF handling via Word Interop; falls back to direct text read.
    ''' Adds header for CSV/TSV. Truncates output at 1,500,000 characters.
    ''' </summary>
    ''' <param name="filePath">Absolute file path.</param>
    ''' <param name="extracted">Extracted text (output).</param>
    ''' <param name="label">Label describing the file (output).</param>
    ''' <returns>True on successful extraction; otherwise False.</returns>
    Private Function TryExtractTextLike(
    ByVal filePath As System.String,
    ByRef extracted As System.String,
    ByRef label As System.String
) As System.Boolean

        extracted = Nothing
        label = Nothing

        If System.String.IsNullOrWhiteSpace(filePath) Then Return False
        If Not System.IO.File.Exists(filePath) Then Return False

        Dim ext As System.String = System.IO.Path.GetExtension(filePath).ToLowerInvariant()

        ' List of common text/code extensions (extensible)
        Dim textLike As System.String() = {
        ".txt", ".log", ".csv", ".tsv", ".md",
        ".json", ".xml", ".yaml", ".yml", ".ini", ".cfg", ".conf", ".toml",
        ".sql",
        ".cs", ".vb", ".vbs", ".js", ".ts", ".jsx", ".tsx",
        ".py", ".rb", ".php", ".java", ".kt", ".kts",
        ".c", ".h", ".hpp", ".hh", ".cpp", ".cc",
        ".ps1", ".psm1", ".bat", ".cmd", ".sh", ".zsh",
        ".rtf", ".html", ".htm"
    }

        If Not textLike.Contains(ext) Then
            Return False
        End If

        Try
            ' Optionally treat RTF as Office document (for plaintext RTF use Word Interop)
            If ext = ".rtf" Then
                Try
                    Dim tmp As System.String = ExtractWordText(filePath) ' uses Word Interop when available
                    If Not System.String.IsNullOrWhiteSpace(tmp) Then
                        extracted = tmp
                        label = "Word-readable (RTF): " & System.IO.Path.GetFileName(filePath)
                        If extracted.Length > 1_500_000 Then
                            extracted = extracted.Substring(0, 1_500_000) & System.Environment.NewLine & "[…truncated…]"
                        End If
                        Return True
                    End If
                Catch
                    ' Fallback: read as text
                End Try
            End If

            Dim content As System.String = ReadAllTextSmart(filePath)
            If System.String.IsNullOrWhiteSpace(content) Then Return False

            ' For CSV/TSV add small header
            If ext = ".csv" OrElse ext = ".tsv" Then
                Dim sepDisplay As System.String = If(ext = ".csv", ",", "\t")
                Dim header As System.String = "=== CSV/TSV Detected (" & ext.Trim("."c).ToUpperInvariant() & ", sep=""" & sepDisplay & """) ==="
                extracted = header & System.Environment.NewLine & content
                label = "Spreadsheet text: " & System.IO.Path.GetFileName(filePath)
            Else
                extracted = content
                label = "Text/code file: " & System.IO.Path.GetFileName(filePath)
            End If

            If extracted.Length > 1_500_000 Then
                extracted = extracted.Substring(0, 1_500_000) & System.Environment.NewLine & "[…truncated…]"
            End If

            Return True
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("Text-like extract failed: " & ex.Message)
            extracted = Nothing
            label = Nothing
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Reads file content with UTF-8 BOM detection; falls back to Windows-1252 then generic default.
    ''' Returns Nothing if all attempts fail.
    ''' </summary>
    ''' <param name="path">Absolute file path.</param>
    ''' <returns>File content or Nothing.</returns>
    Private Function ReadAllTextSmart(ByVal path As System.String) As System.String
        ' UTF-8 (with BOM detection), fallback: Windows-1252 -> UTF-8
        Try
            Using sr As New System.IO.StreamReader(path, System.Text.Encoding.UTF8, detectEncodingFromByteOrderMarks:=True)
                Dim s As System.String = sr.ReadToEnd()
                If Not System.String.IsNullOrEmpty(s) Then Return s
            End Using
        Catch
        End Try
        Try
            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(1252) ' Western Europe Win-1252
            Return System.IO.File.ReadAllText(path, enc)
        Catch
            ' Final fallback
            Try
                Return System.IO.File.ReadAllText(path)
            Catch
                Return Nothing
            End Try
        End Try
    End Function

End Class
