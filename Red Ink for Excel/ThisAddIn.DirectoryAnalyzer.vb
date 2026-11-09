' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.DirectoryAnalyzer.vb
' Purpose: Provides directory-of-TXT-and-Markdown-files analysis via an LLM and writes
'          structured results into the active Excel worksheet.
'          Reads every .txt or .md file in a chosen folder on demand (streamed),
'          constructs a virtual CSV (Filename + Content), then reuses the
'          chunk-processing / LLM / Excel-output pipeline from the CSV Analyzer.
'
' Architecture:
'   1. Folder Selection: Prompts user to pick a directory.
'   2. File Discovery: Enumerates .txt and .md files; shows count and confirms.
'   3. Parameter Collection: Prompt, chunk size, content truncation limit,
'      file range (start / count), output placement, retries, secondary model.
'   4. Secondary Model Prep: Optionally loads an alternate model configuration.
'   5. Excel Output Prep: Lazily inserts a report header block.
'   6. Chunk Processing Loop: Streams files on demand (never loads all at once),
'      builds chunk bodies with a synthetic header line, and flushes via
'      ProcessOneChunk. Each file is assigned a numeric index (LineInFile) so
'      that even lightweight LLMs can reliably echo it back.
'   7. Hyperlink Post-Processing: After each chunk, scans newly inserted rows,
'      resolves numeric keys back to filenames, and inserts clickable hyperlinks.
'   8. Result Insertion: Writes findings or diagnostics to Excel.
'   9. Progress Reporting: Updates a global progress bar; supports cancellation.
'  10. Cleanup: Restores original model config, closes progress UI, reports
'      completion.
'
' Scalability notes (100 000+ files):
'   - File content is read one-at-a-time inside the chunk loop, so memory
'     stays proportional to (chunkSize × maxContentChars), not total file count.
'   - Content is truncated to a configurable character limit per file to avoid
'     exceeding LLM context windows.
'   - Progress bar and cancellation remain responsive because the timer-based
'     ProgressForm polls shared state independently.
'   - The "Starting file" and "Number of files" parameters allow processing
'     arbitrary slices of large directories without enumerating everything.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Main entry to analyze all .txt and .md files in a user-selected directory using an LLM.
    ''' Each file becomes one virtual CSV row (Filename + Content), processed in chunks,
    ''' with results written to the active Excel worksheet.
    ''' Files are read on demand to keep memory usage constant regardless of directory size.
    ''' The LLM receives numeric line indices (reliable for lightweight models); post-processing
    ''' resolves them to clickable filename hyperlinks in the first column.
    ''' </summary>
    Public Async Sub AnalyzeDirectoryWithLLM()

        Dim UseSecondAPI As Boolean

        Try
            ' 1) Ask for folder
            Dim selectedFolder As String = Nothing
            Try
                Using dlg As New System.Windows.Forms.FolderBrowserDialog()
                    dlg.Description = "Select folder containing .txt and .md files to analyze"
                    dlg.ShowNewFolderButton = False
                    If dlg.ShowDialog() <> System.Windows.Forms.DialogResult.OK OrElse String.IsNullOrWhiteSpace(dlg.SelectedPath) Then
                        ShowCustomMessageBox("No folder selected — will abort.")
                        Return
                    End If
                    selectedFolder = dlg.SelectedPath
                End Using
            Catch ex As Exception
                ShowCustomMessageBox("Folder selection failed: " & ex.Message)
                Return
            End Try

            ' 2) Enumerate .txt and .md files (only paths — no content loaded yet)
            Dim txtFiles() As String = {}
            Try
                txtFiles = System.IO.Directory.GetFiles(selectedFolder, "*", System.IO.SearchOption.TopDirectoryOnly).
                    Where(Function(filePath) System.String.Equals(System.IO.Path.GetExtension(filePath), ".txt", System.StringComparison.OrdinalIgnoreCase) OrElse
                                             System.String.Equals(System.IO.Path.GetExtension(filePath), ".md", System.StringComparison.OrdinalIgnoreCase)).ToArray()
            Catch ex As Exception
                ShowCustomMessageBox("Failed to enumerate files: " & ex.Message)
                Return
            End Try

            If txtFiles.Length = 0 Then
                ShowCustomMessageBox("The selected folder contains no .txt or .md files.")
                Return
            End If

            ' Sort alphabetically for consistent ordering
            Array.Sort(txtFiles, StringComparer.OrdinalIgnoreCase)

            ' Build lookup: numeric file index (1-based, as string) -> full path
            ' The LLM receives these integers as LineInFile — trivial to echo back correctly.
            Dim filePathByIndex As New Dictionary(Of String, String)()
            For idx = 0 To txtFiles.Length - 1
                filePathByIndex(CStr(idx + 1)) = txtFiles(idx)
            Next

            ShowCustomMessageBox("Found " & txtFiles.Length & " .txt/.md file(s) in:" & vbCrLf & selectedFolder, $"{AN} Text File Analyzer")

            Dim proceed = ShowCustomYesNoBox("Proceed with analysis?", "Yes", "No", $"{AN} Text File Analyzer")
            If proceed <> 1 Then Return

            ' 3) Use TAB as the virtual-CSV separator (avoids collisions with file content)
            Dim dirSeparator As String = vbTab

            ' 4) Collect parameters
            Dim promptDefault As String = GetSetting("DIR_Prompt", "")
            Dim chunkDefault As Integer = GetSetting(Of Integer)("DIR_ChunkSize", 10)
            Dim maxCharsDefault As Integer = GetSetting(Of Integer)("DIR_MaxContentChars", 4000)
            Dim startFileDefault As Integer = GetSetting(Of Integer)("DIR_StartFile", 0)
            Dim fileCountDefault As Integer = GetSetting(Of Integer)("DIR_FileCount", 0)
            Dim resultStartLineDefault As Integer = GetSetting(Of Integer)("DIR_ResultStartLine", 1)
            Dim resultStartColDefault As Integer = GetSetting(Of Integer)("DIR_ResultStartColumn", 1)
            Dim attemptsDefault As Integer = GetSetting(Of Integer)("DIR_Attempts", 2)

            Dim do2ndModel As Boolean? = GetSetting(Of Boolean)("DIR_UseSecondModel", False)
            If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                ' keep checkbox available
            ElseIf INI_SecondAPI Then
                ' keep checkbox available
            Else
                do2ndModel = CType(Nothing, Boolean?)
            End If

            Dim p0 As New SLib.InputParameter("Prompt for analysis", promptDefault) With {
                .Multiline = True,
                .MultilineHeight = 150
            }
            Dim p1 As New SLib.InputParameter("Chunk size (number of files per LLM call)", chunkDefault)
            Dim p2 As New SLib.InputParameter("Max characters per file (0 = no limit)", maxCharsDefault)
            Dim p3 As New SLib.InputParameter("Starting file number (0 = from beginning)", startFileDefault)
            Dim p4 As New SLib.InputParameter("Number of files to process (0 = all)", fileCountDefault)
            Dim p5 As New SLib.InputParameter("Result start line (row in Excel)", resultStartLineDefault)
            Dim p6 As New SLib.InputParameter("Result start column (1 = A)", resultStartColDefault)
            Dim p7 As New SLib.InputParameter("Number of attempts (in case of errors)", attemptsDefault)

            Dim p8 As SLib.InputParameter
            If do2ndModel.HasValue Then
                p8 = New SLib.InputParameter("Use the secondary model", do2ndModel.Value)
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    p8 = New SLib.InputParameter("Use a secondary model", do2ndModel.Value)
                End If
            Else
                p8 = New SLib.InputParameter("Use a secondary model", CType(Nothing, Boolean?))
            End If

            Dim prms() As SLib.InputParameter = {p0, p1, p2, p3, p4, p5, p6, p7, p8}
            If ShowCustomVariableInputForm("Please set the text analysis parameters:", $"{AN} Text File Analyzer", prms) = False Then
                Return
            End If

            ' Read back + persist
            OtherPrompt = CStr(prms(0).Value)
            Dim chunkSize As Integer = Math.Max(1, CInt(SafeToInt(prms(1).Value, 10)))
            Dim maxContentChars As Integer = CInt(SafeToInt(prms(2).Value, 4000))
            Dim startFile As Integer = Math.Max(0, CInt(SafeToInt(prms(3).Value, 0)))
            Dim fileCount As Integer = Math.Max(0, CInt(SafeToInt(prms(4).Value, 0)))
            Dim resultStartLine As Integer = CInt(SafeToInt(prms(5).Value, 3))
            Dim resultStartCol As Integer = CInt(SafeToInt(prms(6).Value, 1))
            Dim llmAttempts As Integer = Math.Max(1, CInt(SafeToInt(prms(7).Value, 2)))

            UseSecondAPI = False
            If TypeOf prms(8).Value Is Boolean Then
                UseSecondAPI = CBool(prms(8).Value)
            End If

            SaveSetting("DIR_Prompt", OtherPrompt)
            SaveSetting("DIR_ChunkSize", chunkSize)
            SaveSetting("DIR_MaxContentChars", maxContentChars)
            SaveSetting("DIR_StartFile", startFile)
            SaveSetting("DIR_FileCount", fileCount)
            SaveSetting("DIR_ResultStartLine", resultStartLine)
            SaveSetting("DIR_ResultStartColumn", resultStartCol)
            SaveSetting("DIR_Attempts", llmAttempts)
            SaveSetting("DIR_UseSecondModel", UseSecondAPI)
            Try : My.Settings.Save() : Catch : End Try

            If OtherPrompt.Trim().Length < 5 Then
                ShowCustomMessageBox("Please provide a more detailed prompt (at least 5 characters).")
                Return
            End If

            ' Compute the file range to process (1-based indices into txtFiles)
            Dim firstFileIdx As Integer = If(startFile > 0, startFile, 1)
            Dim lastFileIdx As Integer = If(fileCount > 0, Math.Min(firstFileIdx + fileCount - 1, txtFiles.Length), txtFiles.Length)

            If firstFileIdx > txtFiles.Length Then
                ShowCustomMessageBox($"Starting file ({firstFileIdx}) exceeds total file count ({txtFiles.Length}).")
                Return
            End If

            ' 5) If a second/alternate model is desired, prepare
            If UseSecondAPI Then
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                        originalConfigLoaded = False
                        ShowCustomMessageBox("The secondary model could not be loaded — aborting.")
                        Return
                    End If
                End If
            End If

            ' 6) Prepare Excel output header
            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Dim outRow As Integer = Math.Max(1, resultStartLine)
            Dim outCol As Integer = Math.Max(1, resultStartCol)

            If Not PromptForFreshWorksheetIfNeeded(ws, outRow, outCol, $"{AN} Text File Analyzer") Then
                Return
            End If

            Dim headerInserted As Boolean = False
            Dim insertedRowsTotal As Integer = 0

            ' Total files in the selected range (for progress reporting)
            Dim totalFiles As Integer = lastFileIdx - firstFileIdx + 1

            Dim InsertHeader As System.Action =
                    Sub()
                        If headerInserted Then Return

                        Dim headerStartRow As Integer = outRow

                        Dim defaultSize As Double
                        Try
                            Dim cellsRange As Excel.Range = CType(ws.Cells, Excel.Range)
                            defaultSize = CDbl(cellsRange.Font.Size)
                        Catch
                            defaultSize = 11
                        End Try

                        ' Title
                        Dim titleCell As Excel.Range = CType(ws.Cells(outRow, outCol), Excel.Range)
                        titleCell.Value = "Text File Analysis Report"
                        With titleCell.Font
                            .Bold = True
                            .Size = defaultSize + 2
                        End With
                        outRow += 1

                        ' Empty line
                        outRow += 1

                        ' Folder row
                        CType(ws.Cells(outRow, outCol), Excel.Range).Value = "Folder:"
                        CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = selectedFolder
                        outRow += 1

                        ' File count row
                        CType(ws.Cells(outRow, outCol), Excel.Range).Value = "Files:"
                        CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = txtFiles.Length.ToString() & If(totalFiles < txtFiles.Length, $" (processing {totalFiles}: #{firstFileIdx}–#{lastFileIdx})", "")
                        outRow += 1

                        ' Date row
                        CType(ws.Cells(outRow, outCol), Excel.Range).Value = "Date:"
                        CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = DateTime.Now.ToString("G")
                        outRow += 1

                        ' Instruction row
                        CType(ws.Cells(outRow, outCol), Excel.Range).Value = "Prompt:"
                        CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = OtherPrompt
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
                        CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = chunkSize.ToString() & " files"
                        outRow += 1

                        ' Max content chars row
                        CType(ws.Cells(outRow, outCol), Excel.Range).Value = "Max chars/file:"
                        CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = If(maxContentChars > 0, maxContentChars.ToString(), "unlimited")
                        outRow += 1

                        ' Empty line
                        outRow += 1

                        ' Data header
                        CType(ws.Cells(outRow, outCol), Excel.Range).Value = "File"
                        CType(ws.Cells(outRow, outCol + 1), Excel.Range).Value = "Result"
                        Dim headerCellsRange As Excel.Range = CType(ws.Range(ws.Cells(outRow, outCol), ws.Cells(outRow, outCol + 1)), Excel.Range)
                        headerCellsRange.Font.Bold = True

                        Dim headerEndRow As Integer = outRow
                        Dim headerRange As Excel.Range = CType(ws.Range(ws.Cells(headerStartRow, outCol), ws.Cells(headerEndRow, outCol + 1)), Excel.Range)
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                        headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop

                        CType(ws.Cells(outRow, outCol), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

                        ' Set File column width to accommodate filenames
                        Dim fileColumn As Excel.Range = CType(ws.Columns(outCol), Excel.Range)
                        fileColumn.ColumnWidth = 25

                        outRow += 1
                        headerInserted = True
                    End Sub

            ' 7) Build system prompt — temporarily set Separator so {Separator} interpolates to TAB
            Dim originalSeparator As String = Separator
            Separator = dirSeparator
            Dim sysPrompt As String = InterpolateAtRuntime(SP_ParseFile)
            Separator = originalSeparator

            Dim headerColumns As New List(Of String) From {"Filename", "Content"}
            Dim headerOut As String = BuildChunkHeader("LineInFile", headerColumns, dirSeparator)

            ' Progress bar initialization
            Try
                ShowProgressBarInSeparateThread(AN & " Text File Analyzer", "Analyzing files...")
                ProgressBarModule.CancelOperation = False
                GlobalProgressValue = 0
                GlobalProgressMax = totalFiles
                GlobalProgressLabel = $"Processing 0 of {totalFiles} files..."
            Catch
            End Try

            Dim processedFiles As Integer = 0
            Dim cancelled As Boolean = False

            ' 8) Stream files and process in chunks — content is read per-file, never all at once
            Dim chunkBuffer As New System.Text.StringBuilder(64 * 1024)
            Dim chunkFirstLine As Long = 0
            Dim chunkLastLine As Long = 0
            Dim chunkCounter As Integer = 0

            Dim FlushChunk As Func(Of Task) =
                            Async Function() As Task
                                If chunkCounter <= 0 OrElse chunkFirstLine <= 0 Then Return

                                Dim body As String = chunkBuffer.ToString()
                                If body.Length > 0 AndAlso Not body.StartsWith(headerOut, StringComparison.Ordinal) Then
                                    body = headerOut & Environment.NewLine & body
                                End If

                                ' Remember where new rows will start so we can post-process them
                                Dim rowBeforeChunk As Integer = outRow

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
                                    dirSeparator
                                )

                                ' Post-process: resolve numeric line indices back to filename hyperlinks.
                                ' The LLM returns simple integers (e.g. "3") which are trivial for even
                                ' lightweight models. We look up the full path and replace the cell value
                                ' with a clickable hyperlink showing the filename.
                                If outRow > rowBeforeChunk Then
                                    For r = rowBeforeChunk To outRow - 1
                                        Try
                                            Dim cell As Excel.Range = CType(ws.Cells(r, outCol), Excel.Range)
                                            Dim cellVal As String = If(TryCast(cell.Value, String), If(cell.Value IsNot Nothing, cell.Value.ToString(), ""))
                                            If String.IsNullOrWhiteSpace(cellVal) Then Continue For

                                            Dim lineKey As String = cellVal.Trim()
                                            Dim fullPath As String = Nothing

                                            If filePathByIndex.TryGetValue(lineKey, fullPath) Then
                                                ' Single file index — replace with clickable hyperlink showing filename
                                                Dim displayName As String = Path.GetFileName(fullPath)
                                                Try
                                                    ' Excel Hyperlinks.Add fails (0x800A03EC) when the address
                                                    ' exceeds ~255 chars or contains characters Excel cannot handle.
                                                    Dim safeAddress As String = fullPath
                                                    If safeAddress.Length > 255 Then
                                                        ' Too long for Excel hyperlink — just show the filename
                                                        cell.Value = displayName
                                                    Else
                                                        ws.Hyperlinks.Add(
                                                Anchor:=cell,
                                                Address:="file:///" & safeAddress.Replace("\", "/"),
                                                TextToDisplay:=displayName)
                                                        cell.Value = displayName
                                                    End If
                                                Catch exLink As Exception
                                                    ' Hyperlink creation failed — fall back to plain filename
                                                    cell.Value = displayName
                                                    Debug.WriteLine($"Hyperlink failed for '{fullPath}': {exLink.Message}")
                                                End Try
                                            Else
                                                ' Try as a range key (e.g. "3-5") from diagnostic/error rows,
                                                ' or as a bare integer that wasn't in the dictionary
                                                If lineKey.Contains("-") Then
                                                    Dim parts = lineKey.Split("-"c)
                                                    If parts.Length = 2 Then
                                                        Dim startIdx As Integer = 0
                                                        Dim endIdx As Integer = 0
                                                        If Integer.TryParse(parts(0).Trim(), startIdx) AndAlso Integer.TryParse(parts(1).Trim(), endIdx) Then
                                                            Dim names As New List(Of String)()
                                                            For fi = startIdx To endIdx
                                                                Dim fp As String = Nothing
                                                                If filePathByIndex.TryGetValue(fi.ToString(), fp) Then
                                                                    names.Add(Path.GetFileName(fp))
                                                                End If
                                                            Next
                                                            If names.Count > 0 Then
                                                                cell.Value = String.Join(", ", names)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Catch
                                            ' Non-critical: skip hyperlink for this cell
                                        End Try
                                    Next
                                End If

                                Dim filesProcessedThisChunk As Integer = CInt(chunkLastLine - chunkFirstLine + 1)
                                processedFiles += filesProcessedThisChunk
                                Try
                                    GlobalProgressValue = processedFiles
                                    GlobalProgressLabel = $"Processing {processedFiles} of {totalFiles} files..."
                                Catch
                                End Try

                                chunkBuffer.Clear()
                                chunkCounter = 0
                                chunkFirstLine = 0
                                chunkLastLine = 0
                            End Function

            Dim fileIndex As Long = 0
            For Each fp In txtFiles
                If ProgressBarModule.CancelOperation Then
                    cancelled = True
                    Exit For
                End If

                fileIndex += 1

                ' Skip files outside the requested range
                If fileIndex < firstFileIdx Then Continue For
                If fileIndex > lastFileIdx Then Exit For

                Dim fileName As String = Path.GetFileName(fp)

                ' Read file content on demand (not cached) to keep memory flat
                Dim content As String = ""
                Try
                    content = File.ReadAllText(fp)
                    ' Flatten to single line and strip TABs (our separator)
                    content = content.Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbCr, " ").Replace(vbTab, " ")
                    ' Truncate to avoid exceeding LLM context window
                    If maxContentChars > 0 AndAlso content.Length > maxContentChars Then
                        content = content.Substring(0, maxContentChars) & " [TRUNCATED]"
                    End If
                Catch ex As Exception
                    content = "[ERROR reading file: " & ex.Message & "]"
                End Try

                ' Build virtual CSV line: numeric index <TAB> Filename <TAB> Content
                ' The LLM sees the integer as LineInFile and echoes it back reliably.
                If chunkCounter = 0 Then
                    chunkBuffer.AppendLine(headerOut)
                    chunkFirstLine = fileIndex
                End If
                chunkBuffer.Append(fileIndex.ToString())
                chunkBuffer.Append(dirSeparator)
                chunkBuffer.Append(fileName)
                chunkBuffer.Append(dirSeparator)
                chunkBuffer.AppendLine(content)
                chunkCounter += 1
                chunkLastLine = fileIndex

                ' Let content be GC'd before next iteration
                content = Nothing

                If chunkCounter >= chunkSize Then
                    If ProgressBarModule.CancelOperation Then
                        cancelled = True
                        Exit For
                    End If
                    Await FlushChunk()
                End If
            Next

            ' Flush remaining partial chunk
            If Not cancelled Then
                Await FlushChunk()
            End If

            ' Ensure header + "No findings." if nothing was inserted
            If Not cancelled Then
                If insertedRowsTotal = 0 Then
                    If Not headerInserted Then InsertHeader()
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
                outRow += 1
                CType(ws.Cells(outRow, outCol), Excel.Range).Value = "Created by " & AN & " (processed " & processedFiles & " of " & totalFiles & " files)"
                CType(ws.Cells(outRow, outCol), Excel.Range).Font.Italic = True
            End If

            ' Close progress bar and show completion/cancellation message
            Try
                If ProgressBarModule.CancelOperation Then
                    ShowCustomMessageBox("Text file analysis cancelled by user (processed " & processedFiles & " of " & totalFiles & " files).", $"{AN} Text File Analyzer")
                Else
                    ProgressBarModule.CancelOperation = True
                    ShowCustomMessageBox("Text file analysis completed (processed " & processedFiles & " of " & totalFiles & " files).", $"{AN} Text File Analyzer")
                End If
            Catch
            End Try

            If UseSecondAPI AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If

        Catch ex As Exception
            ShowCustomMessageBox("Error in AnalyzeDirectoryWithLLM: " & ex.Message)
        Finally
            If ProgressBarModule.CancelOperation = False Then
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

End Class