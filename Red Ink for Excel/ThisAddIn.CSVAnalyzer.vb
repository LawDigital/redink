' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.CSVAnalyzer.vb
' Purpose: Provides CSV/TXT analysis via an LLM and writes structured results
'          into the active Excel worksheet. Supports dynamic column selection,
'          chunked processing, progress reporting, and optional secondary model use.
'
' Architecture:
'   1. File Selection: Prompts user to choose a CSV/TXT file (drag/drop dialog).
'   2. Separator Input: Retrieves or prompts for the CSV separator (persisted in settings).
'   3. Header/Line Scan: Reads first line (header) and counts remaining lines efficiently.
'   4. Parameter Collection: Gathers analysis parameters (prompt, chunk size, column subset, line range, output placement, retries, secondary model).
'   5. Column Resolution: Maps requested column names to indices; defaults to all when unspecified.
'   6. Secondary Model Prep: Optionally loads an alternate model configuration.
'   7. Excel Output Prep: Lazily inserts a report header block (title, metadata, column headers) when first needed.
'   8. Chunk Processing Loop: Streams file lines in selected range, builds chunk bodies with a synthetic header line, and flushes chunks.
'   9. LLM Interaction: For each chunk, invokes the model, parses structured response (line-key/result pairs) or handles errors/NORESULT.
'  10. Result Insertion: Writes findings (or diagnostic line) to Excel, aligning and wrapping cells.
'  11. Progress Reporting: Updates a global progress bar; supports cancellation.
'  12. Cleanup: Restores original model config (if switched), closes progress UI, and reports completion/cancellation.
'
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Main entry to analyze a user-selected CSV/TXT file using an LLM.
    ''' Collects parameters, resolves columns, processes the file in chunks,
    ''' and writes structured results to the active Excel worksheet with progress reporting.
    ''' </summary>
    Public Async Sub AnalyzeCsvWithLLM()

        Dim UseSecondAPI As Boolean

        Try
            ' 1) Ask for file (prefer CSV/TXT)
            DragDropFormLabel = "CSV or Text files (*.csv; *.txt)"
            DragDropFormFilter =
            "Supported Files|*.csv;*.txt|" &
            "CSV Files (*.csv)|*.csv|" &
            "Text Files (*.txt)|*.txt"
            Dim filePath As String = GetFileName()
            DragDropFormLabel = ""
            DragDropFormFilter = ""
            If String.IsNullOrWhiteSpace(filePath) Then
                ShowCustomMessageBox("No file has been selected - will abort.")
                Return
            End If

            Dim ext As String = IO.Path.GetExtension(filePath).ToLowerInvariant()
            If ext <> ".csv" AndAlso ext <> ".txt" Then
                Dim answer = ShowCustomYesNoBox("The selected file is not .csv or .txt. Continue anyway?", "Yes", "No", $"{AN} CSV Analyzer")
                If answer <> 1 Then Return
            End If

            ' 2) Ask for separator first (needed to parse header correctly)
            Dim sepDefault As String = GetSetting("CSV_Separator", ";")
            Dim sepInput As String = SLib.ShowCustomInputBox("Enter the CSV separator (single character is recommended):", $"{AN} CSV Analyzer", True, sepDefault)
            If String.IsNullOrWhiteSpace(sepInput) Then sepInput = sepDefault
            Separator = sepInput
            SaveSetting("CSV_Separator", Separator)

            ' 3) Open minimally, read header + count lines efficiently
            Dim headerColumns As List(Of String) = Nothing
            Dim dataLinesCount As Long = 0
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1 << 20, FileOptions.SequentialScan)
                Using sr As New StreamReader(fs, detectEncodingFromByteOrderMarks:=True)
                    Dim header As String = sr.ReadLine()
                    If header Is Nothing Then
                        ShowCustomMessageBox("The file appears to be empty.")
                        Return
                    End If
                    headerColumns = ParseCsvLine(header, Separator)

                    ' Count remaining lines efficiently
                    Dim line As String = Nothing
                    While True
                        line = sr.ReadLine()
                        If line Is Nothing Then Exit While
                        dataLinesCount += 1
                    End While
                End Using
            End Using

            ' Show header + line count and confirm
            Dim headerMsg As String = "Header columns (" & headerColumns.Count & "): " & String.Join(" | ", headerColumns)
            Dim linesMsg As String = "Number of lines (excluding header): " & dataLinesCount
            ShowCustomMessageBox(linesMsg & vbCrLf & headerMsg, $"{AN} CSV Analyzer")

            Dim proceed = ShowCustomYesNoBox("Proceed with parsing and analysis?", "Yes", "No", $"{AN} CSV Analyzer")
            If proceed <> 1 Then Return

            ' 4) Collect parameters (single form)
            Dim promptDefault As String = GetSetting("CSV_Prompt", "")
            Dim colsDefault As String = GetSetting("CSV_Columns", "")
            Dim chunkDefault As Integer = GetSetting(Of Integer)("CSV_ChunkSize", 50)
            Dim startSelDefault As Integer = GetSetting(Of Integer)("CSV_StartSelection", 0)
            Dim endSelDefault As Integer = GetSetting(Of Integer)("CSV_EndSelection", 0)
            Dim resultStartLineDefault As Integer = GetSetting(Of Integer)("CSV_ResultStartLine", 1)
            Dim resultStartColDefault As Integer = GetSetting(Of Integer)("CSV_ResultStartColumn", 1)
            Dim attemptsDefault As Integer = GetSetting(Of Integer)("CSV_Attempts", 2)

            Dim do2ndModel As Boolean? = GetSetting(Of Boolean)("CSV_UseSecondModel", False)
            If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                ' keep checkbox available
            ElseIf INI_SecondAPI Then
                ' keep checkbox available
            Else
                ' no secondary available -> disable by setting to Nothing
                do2ndModel = CType(Nothing, Boolean?)
            End If

            Dim p0 As New SLib.InputParameter("Prompt for analysis", promptDefault) With {
                .Multiline = True,
                .MultilineHeight = 150
            }
            Dim p1 As New SLib.InputParameter("CSV separator", Separator)
            Dim p2 As New SLib.InputParameter("Columns to process (empty = all; separate names by same separator)", colsDefault)

            Dim p3 As New SLib.InputParameter("Chunk size in lines", chunkDefault)
            Dim p4 As New SLib.InputParameter("Starting line in file (0 = entire file)", startSelDefault)
            Dim p5 As New SLib.InputParameter("Ending line in file (0 = entire file)", endSelDefault)

            Dim p6 As New SLib.InputParameter("Result start line (row in Excel)", resultStartLineDefault)
            Dim p7 As New SLib.InputParameter("Result start column (1 = A)", resultStartColDefault)
            Dim p8 As New SLib.InputParameter("Number of attempts (in case of errors)", attemptsDefault)

            Dim p9 As SLib.InputParameter
            If do2ndModel.HasValue Then
                p9 = New SLib.InputParameter("Use the secondary model", do2ndModel.Value)
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    p9 = New SLib.InputParameter("Use a secondary model", do2ndModel.Value)
                End If
            Else
                p9 = New SLib.InputParameter("Use a secondary model", CType(Nothing, Boolean?))
            End If

            Dim prms() As SLib.InputParameter = {p0, p1, p2, p3, p4, p5, p6, p7, p8, p9}
            If ShowCustomVariableInputForm("Please set the CSV analysis parameters:", $"{AN} CSV Analyzer", prms) = False Then
                Return
            End If

            ' Read back + persist
            OtherPrompt = CStr(prms(0).Value)
            Separator = CStr(prms(1).Value)
            Dim columnsToProcessRaw As String = CStr(prms(2).Value)

            Dim chunkSize As Integer = CInt(SafeToInt(prms(3).Value, 50))
            Dim startSelection As Integer = CInt(SafeToInt(prms(4).Value, 0))
            Dim endSelection As Integer = CInt(SafeToInt(prms(5).Value, 0))

            Dim resultStartLine As Integer = CInt(SafeToInt(prms(6).Value, 3))
            Dim resultStartCol As Integer = CInt(SafeToInt(prms(7).Value, 1))
            Dim llmAttempts As Integer = Math.Max(1, CInt(SafeToInt(prms(8).Value, 2)))

            UseSecondAPI = False
            If TypeOf prms(9).Value Is Boolean Then
                UseSecondAPI = CBool(prms(9).Value)
            End If

            SaveSetting("CSV_Separator", Separator)
            SaveSetting("CSV_Columns", columnsToProcessRaw)
            SaveSetting("CSV_ChunkSize", chunkSize)
            SaveSetting("CSV_StartSelection", startSelection)
            SaveSetting("CSV_EndSelection", endSelection)
            SaveSetting("CSV_ResultStartLine", resultStartLine)
            SaveSetting("CSV_ResultStartColumn", resultStartCol)
            SaveSetting("CSV_Attempts", llmAttempts)
            SaveSetting("CSV_UseSecondModel", UseSecondAPI)
            SaveSetting("CSV_Prompt", OtherPrompt)
            Try : My.Settings.Save() : Catch : End Try

            If OtherPrompt.Trim().Length < 5 Then
                ShowCustomMessageBox("Please provide a more detailed prompt (at least 5 characters).")
                Return
            End If

            SaveSetting("CSV_Prompt", OtherPrompt)
            Try : My.Settings.Save() : Catch : End Try

            ' 5) Resolve columns to extract (by header names)
            Dim selectedHeaders As New List(Of String)()
            Dim selectedIdx As New List(Of Integer)()
            Dim err As String = ResolveColumns(headerColumns, columnsToProcessRaw, Separator, selectedHeaders, selectedIdx)
            If Not String.IsNullOrEmpty(err) Then
                ShowCustomMessageBox(err)
                Return
            End If

            ' 6) If a second/alternate model is desired, prepare
            If UseSecondAPI Then
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                        originalConfigLoaded = False
                        ShowCustomMessageBox("The secondary model could not be loaded - aborting.")
                        Return
                    End If
                End If
            End If

            ' 7) Prepare Excel output header
            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Dim outRow As Integer = Math.Max(1, resultStartLine)
            Dim outCol As Integer = Math.Max(1, resultStartCol)

            If Not PromptForFreshWorksheetIfNeeded(ws, outRow, outCol, $"{AN} CSV Analyzer") Then
                Return
            End If

            Dim headerInserted As Boolean = False
            Dim insertedRowsTotal As Integer = 0 ' track how many result rows were inserted
            ' Local action to insert Excel report header block (title, metadata, column headers, alignment).

            Dim InsertHeader As System.Action =
                    Sub()
                        If headerInserted Then Return

                        ' Remember where the header starts
                        Dim headerStartRow As Integer = outRow

                        ' Title
                        Dim defaultSize As Double
                        Try
                            ' Cast Cells to Range to avoid late binding
                            Dim cellsRange As Excel.Range = CType(ws.Cells, Excel.Range)
                            defaultSize = CDbl(cellsRange.Font.Size)
                        Catch
                            defaultSize = 11
                        End Try

                        ' Cast each Cells access to Excel.Range
                        Dim titleCell As Excel.Range = CType(ws.Cells(outRow, outCol), Excel.Range)
                        titleCell.Value = "Analysis Report"
                        With titleCell.Font
                            .Bold = True
                            .Size = defaultSize + 2
                        End With
                        outRow += 1

                        ' Empty line
                        outRow += 1

                        ' Filename row
                        CType(ws.Cells(outRow, outCol), Excel.Range).Value = "Filename:"
                        CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = System.IO.Path.GetFileName(filePath)
                        outRow += 1

                        ' Date row
                        CType(ws.Cells(outRow, outCol), Excel.Range).Value = "Date:"
                        CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = DateTime.Now.ToString("G")
                        outRow += 1

                        ' Instruction row
                        CType(ws.Cells(outRow, outCol), Excel.Range).Value = "Prompt:"
                        CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = OtherPrompt

                        ' Ensure the column where the prompt is inserted has width 100 Characters and wraps
                        Dim promptColumn As Excel.Range = CType(ws.Columns(outCol + 1), Excel.Range)
                        promptColumn.ColumnWidth = 100
                        Dim promptCell As Excel.Range = CType(ws.Cells(outRow, outCol + 1), Excel.Range)
                        promptCell.WrapText = True

                        outRow += 1

                        ' Model row
                        CType(ws.Cells(outRow, outCol), Excel.Range).Value = "Model:"
                        CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = If(UseSecondAPI, INI_Model_2, INI_Model)
                        outRow += 1

                        ' Chunksize row
                        CType(ws.Cells(outRow, outCol), Excel.Range).Value = "Chunksize:"
                        CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = chunkSize.ToString()
                        outRow += 1

                        ' Columns row
                        CType(ws.Cells(outRow, outCol), Excel.Range).Value = "Columns:"
                        CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = columnsToProcessRaw.ToString()
                        outRow += 1

                        ' Empty line
                        outRow += 1

                        ' Data header
                        CType(ws.Cells(outRow, outCol), Excel.Range).Value = "Line(s)"
                        CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = "Result"
                        Dim headerCellsRange As Excel.Range = CType(ws.Range(ws.Cells(outRow, outCol), ws.Cells(outRow, outCol + 1)), Excel.Range)
                        headerCellsRange.Font.Bold = True

                        ' Top- and left-align the entire header block (from title through the data header row)
                        Dim headerEndRow As Integer = outRow
                        Dim headerRange As Excel.Range = CType(ws.Range(ws.Cells(headerStartRow, outCol), ws.Cells(headerEndRow, outCol + 1)), Excel.Range)
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                        headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop

                        ' Left-align "Line(s)" header cell
                        CType(ws.Cells(outRow, outCol), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

                        outRow += 1
                        headerInserted = True
                    End Sub

            ' 8) Stream file again and process in chunks
            Dim firstDataAbsLine As Long = If(startSelection > 0, Math.Max(2, startSelection), 2) ' absolute line in file (1 = header)
            Dim lastDataAbsLine As Long = If(endSelection > 0, endSelection, dataLinesCount + 1) ' inclusive, note: header=1, so data runs to totalLines

            If lastDataAbsLine < firstDataAbsLine Then
                ShowCustomMessageBox("The ending line must be greater than or equal to the starting line.")
                Return
            End If

            Dim sysPrompt As String = InterpolateAtRuntime(SP_ParseFile)
            Dim chunkBuffer As New System.Text.StringBuilder(64 * 1024)
            Dim chunkFirstLine As Long = 0
            Dim chunkLastLine As Long = 0
            Dim chunkCounter As Integer = 0

            ' Compute the chunk header once, make it visible to FlushChunk
            Dim headerOut As String = BuildChunkHeader("LineInFile", selectedHeaders, Separator)

            ' Progress bar initialization
            Dim totalLinesToProcess As Integer = CInt(Math.Max(0, lastDataAbsLine - firstDataAbsLine + 1))
            Try
                ShowProgressBarInSeparateThread(AN & " CSV Analyzer", "Analyzing text...")
                ProgressBarModule.CancelOperation = False
                GlobalProgressValue = 0
                GlobalProgressMax = totalLinesToProcess
                GlobalProgressLabel = $"Processing 0 of {totalLinesToProcess} lines..."
            Catch
            End Try

            Dim processedLines As Integer = 0
            Dim cancelled As Boolean = False

            ' Local flush to handle both full and final partial chunks
            Dim FlushChunk As Func(Of Task) =
                            Async Function() As Task
                                If chunkCounter <= 0 OrElse chunkFirstLine <= 0 Then Return

                                ' Ensure header line is present even for the last, partial chunk
                                Dim body As String = chunkBuffer.ToString()
                                If body.Length > 0 AndAlso Not body.StartsWith(headerOut, StringComparison.Ordinal) Then
                                    body = headerOut & Environment.NewLine & body
                                End If

                                Await ProcessOneChunk(
                                    sysPrompt,
                                    body,
                                    chunkFirstLine,
                                    chunkLastLine,
                                    UseSecondAPI,
                                    llmAttempts,
                                    InsertHeader,
                                    Function() outRow,
                                    Sub(n As Integer)
                                        outRow += n
                                        insertedRowsTotal += n
                                    End Sub,
                                    ws,
                                    outCol,
                                    Separator
                                )
                                Dim linesProcessedThisChunk As Integer = CInt(chunkLastLine - chunkFirstLine + 1)
                                processedLines += linesProcessedThisChunk
                                Try
                                    GlobalProgressValue = processedLines
                                    GlobalProgressLabel = $"Processing {processedLines} of {totalLinesToProcess} lines..."
                                Catch
                                End Try

                                ' reset for next chunk
                                chunkBuffer.Clear()
                                chunkCounter = 0
                                chunkFirstLine = 0
                                chunkLastLine = 0
                            End Function

            Using fs2 As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1 << 20, FileOptions.SequentialScan)
                Using sr2 As New StreamReader(fs2, detectEncodingFromByteOrderMarks:=True)
                    ' Skip header
                    Dim headerLine As String = sr2.ReadLine()
                    Dim absLine As Long = 1

                    ' Skip to firstDataAbsLine
                    While absLine < firstDataAbsLine - 1
                        If sr2.ReadLine() Is Nothing Then Exit While
                        absLine += 1
                    End While

                    chunkBuffer.Clear()
                    Dim line As String = Nothing

                    While True
                        If ProgressBarModule.CancelOperation Then
                            cancelled = True
                            Exit While
                        End If

                        line = sr2.ReadLine()
                        If line Is Nothing Then Exit While
                        absLine += 1
                        If absLine > lastDataAbsLine Then Exit While

                        Dim fields As List(Of String) = ParseCsvLine(line, Separator)
                        Dim one As String = BuildChunkLine(absLine, fields, selectedIdx, Separator)
                        If chunkCounter = 0 Then
                            chunkBuffer.AppendLine(headerOut)
                            chunkFirstLine = absLine
                        End If
                        chunkBuffer.AppendLine(one)
                        chunkCounter += 1
                        chunkLastLine = absLine

                        If chunkCounter >= Math.Max(1, chunkSize) Then
                            If ProgressBarModule.CancelOperation Then
                                cancelled = True
                                Exit While
                            End If
                            Await FlushChunk() ' flush full chunk
                        End If
                    End While

                    ' Always flush the remainder if any (final partial chunk)
                    If Not cancelled Then
                        Await FlushChunk()
                    End If

                    ' Ensure header + "No findings." if nothing was inserted
                    If Not cancelled Then
                        If insertedRowsTotal = 0 Then
                            If Not headerInserted Then InsertHeader()
                            ' First line of the table with "No findings." in the Result column
                            CType(ws.Cells(outRow, outCol), Excel.Range).Value = ""
                            CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = "No findings."
                            Dim nfRange As Excel.Range = ws.Range(ws.Cells(outRow, outCol), ws.Cells(outRow, outCol + 1))
                            nfRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                            nfRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop
                            CType(ws.Cells(outRow, outCol + 1), Excel.Range).WrapText = True
                            outRow += 1
                        End If
                    End If

                    If Not cancelled AndAlso headerInserted Then
                        outRow += 1 ' empty line

                        CType(ws.Cells(outRow, outCol), Excel.Range).Value = "Created by " & AN & " (processed " & processedLines & " of " & totalLinesToProcess & " lines)"
                        CType(ws.Cells(outRow, outCol), Excel.Range).Font.Italic = True
                    End If

                End Using
            End Using

            ' Close progress bar and show cancellation/completion message
            Try
                If ProgressBarModule.CancelOperation Then
                    ' user cancelled
                    ShowCustomMessageBox("CSV analysis cancelled by user (processed " & processedLines & " of " & totalLinesToProcess & " lines).", $"{AN} CSV Analyzer")
                Else
                    ' mark finished so progress thread can close
                    ProgressBarModule.CancelOperation = True
                    ShowCustomMessageBox("CSV analysis completed (processed " & processedLines & " of " & totalLinesToProcess & " lines).", $"{AN} CSV Analyzer")
                End If
            Catch
            End Try

            If UseSecondAPI AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If

        Catch ex As Exception
            ShowCustomMessageBox("Error in AnalyzeCsvWithLLM: " & ex.Message)
        Finally

            If ProgressBarModule.CancelOperation = False Then
                ' ensure progress bar is closed
                Try
                    ProgressBarModule.CancelOperation = True
                Catch
                End Try
            End If

            If UseSecondAPI AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If
        End Try
    End Sub


    ''' <summary>
    ''' Processes one chunk: calls LLM with system and user prompts, parses structured response into (line, result) tuples,
    ''' retries on failure, writes results or a diagnostic line into Excel. Handles special token [NORESULT].
    ''' </summary>
    ''' <param name="sysPrompt">System prompt string used for LLM call.</param>
    ''' <param name="chunkBody">Serialized chunk body including header line.</param>
    ''' <param name="chunkFirstLine">First absolute line number in this chunk.</param>
    ''' <param name="chunkLastLine">Last absolute line number in this chunk.</param>
    ''' <param name="useSecond">True to use secondary model configuration.</param>
    ''' <param name="attempts">Retry attempts on failure.</param>
    ''' <param name="ensureHeader">Action to ensure the report header is inserted.</param>
    ''' <param name="getOutRow">Function returning current output row index.</param>
    ''' <param name="advanceOutRow">Action to advance output row by number of inserted lines.</param>
    ''' <param name="ws">Target Excel worksheet.</param>
    ''' <param name="outCol">Starting output column (Line(s) column).</param>
    ''' <param name="separator">CSV separator used.</param>
    Private Async Function ProcessOneChunk(
                        ByVal sysPrompt As String,
                        ByVal chunkBody As String,
                        ByVal chunkFirstLine As Long,
                        ByVal chunkLastLine As Long,
                        ByVal useSecond As Boolean,
                        ByVal attempts As Integer,
                        ByVal ensureHeader As System.Action,
                        ByVal getOutRow As System.Func(Of Integer),
                        ByVal advanceOutRow As System.Action(Of Integer),
                        ByVal ws As Microsoft.Office.Interop.Excel.Worksheet,
                        ByVal outCol As Integer,
                        ByVal separator As String
                    ) As Task

        Dim userPrompt As String = "<LINESTOPROCESS>" & chunkBody & "</LINESTOPROCESS>"

        Dim lastError As String = ""
        Dim parsed As List(Of Tuple(Of String, String)) = Nothing
        Dim noResult As Boolean = False

        For retry = 1 To Math.Max(1, attempts)
            Dim resp As String = ""
            Try
                resp = Await LLM(sysPrompt, userPrompt, "", "", 0, useSecond, True, OtherPrompt)
            Catch ex As Exception
                resp = ""
                lastError = "LLM call failed: " & ex.Message
            End Try

            resp = If(resp, "").Trim()

            If String.Equals(resp, "[NORESULT]", StringComparison.OrdinalIgnoreCase) Then
                ' Do not retry and do not insert anything into Excel for NORESULT
                noResult = True
                Exit For
            End If

            If Not String.IsNullOrWhiteSpace(resp) Then
                parsed = TryParseLLMResponse(resp)
                If parsed IsNot Nothing AndAlso parsed.Count > 0 Then
                    ' success
                    ensureHeader?.Invoke()

                    Dim startRow As Integer = If(getOutRow Is Nothing, 1, getOutRow())
                    Dim row As Integer = startRow

                    For Each t In parsed
                        Dim lineKey As String = t.Item1.Trim()
                        Dim result As String = t.Item2
                        CType(ws.Cells(row, outCol), Excel.Range).Value = lineKey
                        CType(ws.Cells(row, outCol + 1), Excel.Range).Value = result
                        row += 1
                    Next

                    If row > startRow Then
                        ' Align both columns left/top for all inserted rows
                        Dim listRange As Excel.Range = ws.Range(ws.Cells(startRow, outCol), ws.Cells(row - 1, outCol + 1))
                        listRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                        listRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop

                        ' Wrap text in the Result column
                        Dim resultRange As Excel.Range = ws.Range(ws.Cells(startRow, outCol + 1), ws.Cells(row - 1, outCol + 1))
                        resultRange.WrapText = True

                        ' Keep line keys as text (existing behavior)
                        Dim lineKeyRange As Excel.Range = ws.Range(ws.Cells(startRow, outCol), ws.Cells(row - 1, outCol))
                        lineKeyRange.NumberFormat = "@"
                    End If

                    If advanceOutRow IsNot Nothing Then
                        advanceOutRow(row - startRow)
                    End If

                    Return
                Else
                    lastError = "Could not parse the LLM response."
                End If
            Else
                lastError = "Empty response from LLM."
            End If

            ' Backoff before retry
            Await Task.Delay(750 * retry)
        Next

        ' If NORESULT: insert nothing and return silently
        If noResult Then
            Return
        End If

        ' For other failures, write one diagnostic line
        ensureHeader?.Invoke()
        Dim startErrRow As Integer = If(getOutRow Is Nothing, 1, getOutRow())
        Dim rangeKey As String = If(chunkFirstLine = chunkLastLine, chunkFirstLine.ToString(), $"{chunkFirstLine}-{chunkLastLine}")
        CType(ws.Cells(startErrRow, outCol), Excel.Range).Value = rangeKey
        CType(ws.Cells(startErrRow, outCol + 1), Excel.Range).Value = If(String.IsNullOrWhiteSpace(lastError), "Empty or unparsable response.", lastError)
        If advanceOutRow IsNot Nothing Then
            advanceOutRow(1)
        End If
    End Function

    ''' <summary>
    ''' Parses LLM response in the format "line@@result§§§line@@result..." into a list of (line, result) tuples.
    ''' </summary>
    Private Function TryParseLLMResponse(ByVal response As String) As List(Of Tuple(Of String, String))
        Dim result As New List(Of Tuple(Of String, String))()

        Dim records = response.Split(New String() {"§§§"}, StringSplitOptions.RemoveEmptyEntries)
        For Each rec In records
            Dim trimmed = rec.Trim()
            If trimmed.Length = 0 Then Continue For
            Dim parts = trimmed.Split(New String() {"@@"}, 2, StringSplitOptions.None)
            If parts.Length = 2 Then
                Dim key = parts(0).Trim()
                Dim val = parts(1).Trim()
                If key.Length > 0 AndAlso val.Length >= 0 Then
                    result.Add(Tuple.Create(key, val))
                End If
            End If
        Next

        Return result
    End Function

    ''' <summary>
    ''' Builds the synthetic header line for chunk bodies: first header token followed by column headers joined by separator.
    ''' </summary>
    Private Function BuildChunkHeader(ByVal firstHeader As String, ByVal headers As List(Of String), ByVal separator As String) As String
        Return firstHeader & separator & String.Join(separator, headers)
    End Function

    ''' <summary>
    ''' Builds one chunk data line: absolute line number + selected field values joined by separator.
    ''' </summary>
    Private Function BuildChunkLine(ByVal absLine As Long, ByVal fields As List(Of String), ByVal selectedIdx As List(Of Integer), ByVal separator As String) As String
        Dim vals As New List(Of String)(selectedIdx.Count)
        For Each idx In selectedIdx
            Dim v As String = If(idx >= 0 AndAlso idx < fields.Count, fields(idx), "")
            vals.Add(v)
        Next
        Return absLine.ToString() & separator & String.Join(separator, vals)
    End Function

    ''' <summary>
    ''' Resolves requested column names against the header. If none provided or empty, selects all.
    ''' Returns empty string on success or an error message listing missing columns.
    ''' </summary>
    Private Function ResolveColumns(
        ByVal headerColumns As List(Of String),
        ByVal columnsRaw As String,
        ByVal separator As String,
        ByRef selectedHeaders As List(Of String),
        ByRef selectedIdx As List(Of Integer)
    ) As String
        selectedHeaders = New List(Of String)()
        selectedIdx = New List(Of Integer)()

        If headerColumns Is Nothing OrElse headerColumns.Count = 0 Then
            Return "The header row could not be parsed."
        End If

        If String.IsNullOrWhiteSpace(columnsRaw) Then
            ' all columns
            For i = 0 To headerColumns.Count - 1
                selectedHeaders.Add(headerColumns(i))
                selectedIdx.Add(i)
            Next
            Return ""
        End If

        ' Split by the same separator (user-defined)
        Dim requested = columnsRaw.Split(New String() {separator}, StringSplitOptions.RemoveEmptyEntries).
                                 Select(Function(s) s.Trim()).
                                 Where(Function(s) s.Length > 0).
                                 ToList()

        If requested.Count = 0 Then
            ' fallback to all
            For i = 0 To headerColumns.Count - 1
                selectedHeaders.Add(headerColumns(i))
                selectedIdx.Add(i)
            Next
            Return ""
        End If

        ' Build a case-insensitive map from header -> index
        Dim map As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For i = 0 To headerColumns.Count - 1
            If Not map.ContainsKey(headerColumns(i)) Then
                map.Add(headerColumns(i), i)
            End If
        Next

        Dim missing As New List(Of String)()
        For Each name In requested
            Dim idx As Integer
            If map.TryGetValue(name, idx) Then
                selectedHeaders.Add(name)
                selectedIdx.Add(idx)
            Else
                missing.Add(name)
            End If
        Next

        If missing.Count > 0 Then
            Return "These columns were not found in the header: " & String.Join(", ", missing)
        End If

        Return ""
    End Function

    ''' <summary>
    ''' Parses a CSV line with support for quoted fields and escaped quotes using the chosen separator.
    ''' </summary>
    Private Function ParseCsvLine(ByVal line As String, ByVal separator As String) As List(Of String)
        Dim sep As Char = If(String.IsNullOrEmpty(separator), ";"c, separator(0))
        Dim res As New List(Of String)()
        If line Is Nothing Then Return res

        Dim sb As New System.Text.StringBuilder()
        Dim inQuotes As Boolean = False
        Dim i As Integer = 0

        While i < line.Length
            Dim c As Char = line(i)
            If c = """"c Then
                If inQuotes AndAlso i + 1 < line.Length AndAlso line(i + 1) = """"c Then
                    ' escaped quote
                    sb.Append(""""c)
                    i += 2
                    Continue While
                Else
                    inQuotes = Not inQuotes
                    i += 1
                    Continue While
                End If
            End If

            If Not inQuotes AndAlso c = sep Then
                res.Add(sb.ToString())
                sb.Clear()
                i += 1
                Continue While
            End If

            sb.Append(c)
            i += 1
        End While

        res.Add(sb.ToString())
        Return res
    End Function

    ''' <summary>
    ''' Retrieves a typed setting value or returns a default if not present.
    ''' </summary>
    Private Function GetSetting(Of T)(ByVal key As String, ByVal defaultValue As T) As T
        Try
            Dim p = My.Settings.Properties(key)
            If p IsNot Nothing Then
                Dim v = My.Settings.Item(key)
                If v IsNot Nothing Then
                    Return CType(v, T)
                End If
            End If
        Catch
        End Try
        Return defaultValue
    End Function

    ''' <summary>
    ''' Saves a setting value defensively (ignores missing keys).
    ''' </summary>
    Private Sub SaveSetting(Of T)(ByVal key As String, ByVal value As T)
        Try
            Dim p = My.Settings.Properties(key)
            If p IsNot Nothing Then
                My.Settings.Item(key) = value
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Converts an object to Integer or returns fallback on failure.
    ''' </summary>
    Private Function SafeToInt(value As Object, fallback As Integer) As Integer
        Try
            If value Is Nothing Then Return fallback
            Dim s = value.ToString().Trim()
            Dim n As Integer
            If Integer.TryParse(s, n) Then Return n
        Catch
        End Try
        Return fallback
    End Function

End Class
