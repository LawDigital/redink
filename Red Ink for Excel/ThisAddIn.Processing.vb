' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.vb
' Purpose: Cell and range processing with LLM integration. Handles individual cell
'          processing (formulas and values), range-based operations, batch file
'          processing, and worksheet content extraction for LLM prompts.
'
' Architecture:
'   - ProcessSelectedRange: Main entry point for cell/range processing. Supports
'     two modes: (1) individual cell iteration with LLM calls per cell, and
'     (2) range-level processing where the entire selection is serialized and
'     sent to the LLM. Batch mode processes files from a directory.
'   - GatherSelectedWorksheets: Collects content from user-selected worksheets
'     and wraps them in <RANGEOFCELLS> tags for multi-sheet LLM prompts.
'   - ConvertRangeToString: Serializes Excel ranges into a structured text format,
'     including cell values, formulas, comments (legacy and threaded), dropdown
'     validation options, and optional color information.
'   - Supports undo state tracking via undoStates collection for reverting changes.
'   - Handles formula insertion with fallback to FormulaLocal and locale conversion.
' =============================================================================

Option Strict Off  ' late binding required for legacy comments  
Option Explicit On

Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Processes the selected range of cells in Excel or operates on entire ranges/batches.
    ''' </summary>
    ''' <param name="SysCommand">System command to be executed by the LLM.</param>
    ''' <param name="CheckMaxToken">Whether to check if estimated tokens exceed the LLM's maximum output.</param>
    ''' <param name="DoRange">If True, processes the selection as a range (single LLM call); if False, iterates individual cells.</param>
    ''' <param name="DoFormulas">Whether to include formulas in the processing.</param>
    ''' <param name="DoBubbles">Whether to insert comments (bubbles) into cells.</param>
    ''' <param name="SelectionMandatory">Whether a cell selection is required before processing.</param>
    ''' <param name="UseSecondAPI">Whether to use the secondary API configuration.</param>
    ''' <param name="ShortenPercentValue">Percentage by which to shorten text (0 = no shortening).</param>
    ''' <param name="Freestyle">Whether to enable freestyle mode.</param>
    ''' <param name="DoColor">Whether to include color information in range serialization.</param>
    ''' <param name="DoPane">Whether output should be displayed in a pane instead of a dialog.</param>
    ''' <param name="FileObject">Path to a file or clipboard object to include in the LLM call.</param>
    ''' <param name="InsertWS">Text representation of additional worksheets to include in the prompt.</param>
    ''' <param name="BatchFilePath">Path to a directory containing files for batch processing.</param>
    Private Async Function ProcessSelectedRange(ByVal SysCommand As String, CheckMaxToken As Boolean, DoRange As Boolean, DoFormulas As Boolean, DoBubbles As Boolean, SelectionMandatory As Boolean, ByVal UseSecondAPI As Boolean, Optional ShortenPercentValue As Integer = 0, Optional Freestyle As Boolean = False, Optional DoColor As Boolean = False, Optional DoPane As Boolean = False, Optional FileObject As String = "", Optional InsertWS As String = "", Optional BatchFilePath As String = "") As Task(Of Boolean)

        Dim excelApp As Excel.Application = CType(Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Excel.Application)

        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        Dim NoSelectedCells As Boolean = False
        Dim DoShorten As Boolean = False

        If DoBubbles Then SelectionMandatory = True

        ' Get the used range of the active sheet        
        Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim usedRange As Excel.Range = activeSheet.UsedRange

        ' Check if a selection has been made
        If selectedRange Is Nothing Then
            NoSelectedCells = True
        Else
            ' If the entire row, column, or sheet is selected, limit to used range
            selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)

            ' If the intersection results in no cells, set NoSelectedCells to True
            If selectedRange Is Nothing Then
                NoSelectedCells = True
                If Freestyle Or Not SelectionMandatory Then
                    DoRange = True
                    Freestyle = True
                    SysCommand = SP_RangeOfCells
                End If
            End If
        End If

        ' Check if cells are selected and show message if mandatory selection is required
        If NoSelectedCells AndAlso SelectionMandatory Then
            ShowCustomMessageBox("Please select cells (with content) to be processed.")
            Return False
        End If

        ' Check if all selected cells are blocked
        If AreAllCellsBlocked(selectedRange) And Not DoRange Then
            ShowCustomMessageBox($"{AN} cannot do anything because the cells are blocked.")
            Return False
        End If

        If ShortenPercentValue > 0 Then
            DoShorten = True
        End If

        Dim MaxToken As Integer = If(UseSecondAPI, INI_MaxOutputToken_2, INI_MaxOutputToken)
        If Not NoSelectedCells And CheckMaxToken And MaxToken > 0 Then

            SelectedText = GetSelectedText(selectedRange, DoFormulas)

            Dim EstimatedTokens As Integer = EstimateTokenCount(SelectedText)

            If EstimatedTokens > MaxToken Then
                ShowCustomMessageBox("The content of the selected cells is larger than the maximum output your LLM can supposedly generate. Therefore, the output may be shorter than expected based on maximum tokens supported, which is " & MaxToken & " tokens. Your input (with formatting information, as the case may be) is estimated to be " & EstimatedTokens & " tokens. Therefore, check whether the output is complete.", AN, 15)
            End If

        End If

        If Not DoShorten And BatchFilePath = "" Then SysCommand = InterpolateAtRuntime(SysCommand)

        If DoBubbles Then SysCommand = InterpolateAtRuntime(SP_BubblesExcel)

        undoStates.Clear()

        If Not DoRange Then

            Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")
            splash.Show()
            splash.Refresh()

            'Application.ScreenUpdating = False ' Prevent UI updates during processing
            Try
                For Each cell As Excel.Range In selectedRange.Cells

                    System.Windows.Forms.Application.DoEvents()

                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For

                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                        Exit For
                    End If
                    Try
                        If Not IsNothing(cell.Value) AndAlso Not CellProtected(cell) Then
                            If CBool(cell.HasFormula) AndAlso DoFormulas Then
                                ' Handle formulas
                                SelectedText = CStr(cell.Formula)

                                If DoShorten Then
                                    Dim Textlength As Integer = getnumberofwords(SelectedText)
                                    ShortenLength = Textlength * (100 - ShortenPercentValue) / 100
                                    SysCommand = InterpolateAtRuntime(SysCommand)
                                End If

                                Await System.Threading.Tasks.Task.Delay(500)

                                Dim LLMResult As String = Await LLM(SysCommand & " " & SP_Add_KeepFormulasIntact, If(NoSelectedCells, "", "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>"), "", "", 0, UseSecondAPI, True, "", FileObject)

                                LLMResult = Trim(LLMResult)

                                If Not String.IsNullOrEmpty(LLMResult) Then
                                    LLMResult = Await PostCorrection(LLMResult, UseSecondAPI)
                                End If
                                If Not String.IsNullOrWhiteSpace(LLMResult) Then
                                    Dim state As New CellState With {
                                                                    .WorksheetName = cell.Worksheet.Name,
                                                                    .CellAddress = cell.Address,
                                                                    .OldValue = cell.Value,
                                                                    .HadFormula = CBool(cell.HasFormula),
                                                                    .OldFormula = If(CBool(cell.HasFormula), CStr(cell.Formula), "")
                                                                }
                                    Try
                                        cell.Formula = LLMResult
                                        undoStates.Add(state)
                                    Catch ex As Exception
                                        If ex.Message.Contains("HRESULT: 0x800A03EC") Then
                                            Try
                                                cell.FormulaLocal = LLMResult
                                                undoStates.Add(state)
                                            Catch ex2 As Exception
                                                If ex2.Message.Contains("HRESULT: 0x800A03EC") Then
                                                    Try
                                                        cell.FormulaLocal = Trim(ConvertFormulaToLocale(LLMResult, excelApp))
                                                        undoStates.Add(state)
                                                    Catch ex3 As Exception
                                                        If ex.Message.Contains("HRESULT: 0x800A03EC") Then
                                                            ShowCustomMessageBox($"Error: Excel rejected the formula '{LLMResult}' that {AN} tried to assign to the cell {cell.Address(False, False)}.")
                                                        Else
                                                            ShowCustomMessageBox($"An error occurred when trying to insert the formula '{LLMResult}' in cell {cell.Address(False, False)}: {ex.Message}")
                                                        End If
                                                    End Try
                                                Else
                                                    ShowCustomMessageBox($"An error occurred when trying to insert the formula '{LLMResult}' in cell {cell.Address(False, False)}: {ex.Message}")
                                                End If
                                            End Try
                                        Else
                                            ShowCustomMessageBox($"An error occurred when trying to insert the formula '{LLMResult}' in cell {cell.Address(False, False)}: {ex.Message}")
                                        End If
                                    End Try
                                End If
                            ElseIf Not CBool(cell.HasFormula) Then
                                ' Handle plain text cells
                                SelectedText = CStr(cell.Value)

                                If DoShorten Then
                                    Dim Textlength As Integer = SelectedText.Length
                                    ShortenLength = (Textlength - (Textlength * (100 - ShortenPercentValue) / 100))
                                    SysCommand = InterpolateAtRuntime(SysCommand)
                                End If

                                Await System.Threading.Tasks.Task.Delay(500)

                                Dim LLMResult As String = Await LLM(SysCommand, If(NoSelectedCells, "", "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>"), "", "", 0, UseSecondAPI, True, "", FileObject)

                                If Not String.IsNullOrEmpty(LLMResult) Then
                                    LLMResult = Await PostCorrection(LLMResult, UseSecondAPI)
                                End If

                                ' Remove all trailing CR/LF characters in one pass
                                LLMResult = Trim(LLMResult).TrimEnd(ControlChars.Lf, ControlChars.Cr).TrimEnd(ControlChars.Lf, ControlChars.Cr).TrimEnd(ControlChars.Lf, ControlChars.Cr).TrimEnd(ControlChars.Lf, ControlChars.Cr)

                                If Not String.IsNullOrWhiteSpace(LLMResult) Then
                                    Dim state As New CellState With {
                                                                    .WorksheetName = cell.Worksheet.Name,
                                                                    .CellAddress = cell.Address,
                                                                    .OldValue = cell.Value,
                                                                    .HadFormula = CBool(cell.HasFormula),
                                                                    .OldFormula = If(CBool(cell.HasFormula), CStr(cell.Formula), "")
                                                                }
                                    cell.Value = LLMResult
                                    undoStates.Add(state)
                                End If
                            End If
                        End If

                    Catch ex As Exception
                        Debug.WriteLine($"ProcessSelectedRange Error processing cell {cell.Address}: {ex.Message}")
                    End Try
                Next
            Finally
                'Application.ScreenUpdating = True ' Re-enable UI updates
            End Try

            splash.Close()

        Else
            Try

                If NoSelectedCells Then
                    activeSheet.Application.ActiveCell.Select()
                    selectedRange = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
                    If selectedRange Is Nothing Then
                        SelectedText = ""
                        Try
                            SelectedText = $"Current cell = {activeSheet.Application.ActiveCell.Address(False, False)} Text = '{activeSheet.Application.ActiveCell.Text}' Formula = '{activeSheet.Application.ActiveCell.Formula}' (use this for your output unless instructed otherwise)"
                            Debug.WriteLine("NoSelectedCell - SelectedText = " & SelectedText)
                        Catch
                        End Try
                    Else
                        NoSelectedCells = False
                    End If
                End If

                If Not NoSelectedCells Then
                    SelectedText = ConvertRangeToString(selectedRange, DoFormulas, DoColor)
                End If

                Dim RangeToInsert As String = ""

                If InsertWS = "" Then
                    RangeToInsert = "<RANGEOFCELLS>" & SelectedText & "</RANGEOFCELLS>"
                Else
                    RangeToInsert = "Currently active Worksheet: <RANGEOFCELLS>" & SelectedText & "</RANGEOFCELLS>  " & InsertWS
                End If

                If BatchFilePath = "" Then

                    Dim LLMResult As String = Await LLM(SysCommand, If(NoSelectedCells, SelectedText, RangeToInsert), "", "", 0, UseSecondAPI, False, OtherPrompt, FileObject)

                    If InsertWS = "" Then
                        LLMResult = LLMResult.Replace("<RANGEOFCELLS>", "").Replace("</RANGEOFCELLS>", "")
                    Else
                        LLMResult = Regex.Replace(LLMResult, "</?RANGEOFCELLS\d*>", "", RegexOptions.IgnoreCase)
                    End If

                    OtherPrompt = ""

                    If Not String.IsNullOrEmpty(LLMResult) Then
                        LLMResult = Await PostCorrection(LLMResult, UseSecondAPI)
                    End If

                    Dim instructions As New List(Of String)
                    instructions = ParseLLMResponse(LLMResult)

                    If instructions.Count > 0 Then

                        If DoPane Then
                            SP_MergePrompt_Cached = ""
                            ShowPaneAsync("The LLM has provided the following result (you can edit it):", LLMResult, $"You can let {AN} insert the square brackets into your worksheet, where possible", AN, False, True, True)
                        Else
                            Dim FinalText = ShowCustomWindow("The LLM has provided the following result (you can edit it):", LLMResult, $"Shall {AN} insert the square brackets into your worksheet, where possible?", AN, False, False, False, True, Nothing, True)

                            If FinalText = "Pane" Then
                                SP_MergePrompt_Cached = ""
                                ShowPaneAsync("The LLM has provided the following result (you can edit it):", LLMResult, $"You can let {AN} insert the square brackets into your worksheet, where possible", AN, False, True, True)
                            ElseIf Not String.IsNullOrWhiteSpace(FinalText) Then
                                instructions = ParseLLMResponse(FinalText)
                                ApplyLLMInstructions(instructions, DoBubbles)
                                PutInClipboard(FinalText)
                                ShowCustomMessageBox("Implementation of the instructions completed (to the extent possible). They are also in the clipboard.")
                            End If
                        End If
                    Else
                        If DoPane Then
                            SP_MergePrompt_Cached = ""
                            ShowPaneAsync("The LLM has provided the following result (you can edit it):", LLMResult, "Choose to copy your edited or original text to clipboard. You can also copy & paste from the pane.", AN, False, True)
                        Else
                            Dim FinalText = ShowCustomWindow("The LLM has provided the following result (you can edit it):", LLMResult, "If you chose OK, it will be put in the clipboard.", AN)
                            If FinalText = "Pane" Then
                                SP_MergePrompt_Cached = ""
                                ShowPaneAsync("The LLM has provided the following result (you can edit it):", LLMResult, "Choose to copy your edited or original text to clipboard. You can also copy & paste from the pane.", AN, False, True)
                            ElseIf Not String.IsNullOrWhiteSpace(FinalText) Then
                                PutInClipboard(FinalText)
                            End If
                        End If
                    End If
                Else

                    Dim FileContentString As String = ""
                    Dim TempSysCommand As String = ""

                    Dim eligibleFiles As New System.Collections.Generic.List(Of System.String)()
                    Dim DoOCR As Boolean = False
                    Dim HasPDF As Boolean = False

                    Try
                        If Not System.String.IsNullOrWhiteSpace(BatchFilePath) AndAlso System.IO.Directory.Exists(BatchFilePath) Then
                            For Each filePath As System.String In System.IO.Directory.EnumerateFiles(BatchFilePath, "*.*", System.IO.SearchOption.TopDirectoryOnly)
                                Dim ext As System.String = System.IO.Path.GetExtension(filePath)
                                If allowedExtensions.Contains(ext) Then
                                    eligibleFiles.Add(filePath)
                                    If String.Equals(ext, ".pdf", StringComparison.OrdinalIgnoreCase) Then
                                        HasPDF = True
                                    End If
                                End If
                            Next
                        Else
                            eligibleFiles.Clear()
                        End If
                    Catch ex As System.Exception
                        eligibleFiles.Clear()
                    End Try

                    If HasPDF Then
                        If SharedMethods.IsOcrAvailable(_context) Then
                            Dim answer As Integer = ShowCustomYesNoBox(
                                "The selected directory contains PDF files. Do you want to use your model's OCR capabilities to extract text from PDFs that do not appear to contain searchable text?",
                                "Yes, use OCR as needed",
                                "No, do it without OCR")
                            If answer = 1 Then
                                DoOCR = True
                            ElseIf answer <> 2 Then
                                Return Nothing
                            End If
                        Else
                            Dim Answer = ShowCustomYesNoBox("The selected directory contains PDF files. However, you have not configured a model that can do OCR. Therefore, only searchable text can be extracted from the PDFs. Continue anyway?", "Yes, continue", "No, abort")
                            If Answer <> 1 Then
                                Return Nothing
                            End If
                        End If
                    End If

                    Dim MaxEligibleFiles As System.Int32 = eligibleFiles.Count
                    Dim FileCounter As System.Int32 = 0

                    ShowProgressBarInSeparateThread($"{AN} Freestyle Batch", "Starting file processing...")
                    ProgressBarModule.CancelOperation = False
                    GlobalProgressValue = 0
                    GlobalProgressMax = MaxEligibleFiles

                    ' Main processing loop
                    For Each filePath As System.String In eligibleFiles
                        FileCounter += 1

                        If ProgressBarModule.CancelOperation Then
                            ShowCustomMessageBox("Batch processing cancelled by user.")
                            Exit For
                        End If

                        GlobalProgressValue = FileCounter
                        GlobalProgressLabel = $"Processing {FileCounter} of {MaxEligibleFiles} files..."

                        Try
                            If Not System.IO.File.Exists(filePath) Then
                                FileContentString = "Error: File not found: " & filePath
                                Continue For
                            End If

                            FileContentString = Await GetFileContent(filePath, True, DoOCR, False)

                        Catch ex As System.Exception
                            FileContentString = "File Error: " & ex.Message
                        End Try

                        TempSysCommand = InterpolateAtRuntime(SysCommand & " " & SP_Add_Batch)

                        Debug.WriteLine(TempSysCommand)

                        Dim LLMResult As String = Await LLM(TempSysCommand, $"File = '{System.IO.Path.GetFileName(filePath)}' <FILECONTENT>" & FileContentString & "</FILECONTENT>" & vbCrLf & "Other input to consider (if any): " & If(NoSelectedCells, SelectedText, RangeToInsert), "", "", 0, UseSecondAPI, False, OtherPrompt, FileObject)

                        If InsertWS = "" Then
                            LLMResult = LLMResult.Replace("<RANGEOFCELLS>", "").Replace("</RANGEOFCELLS>", "")
                        Else
                            LLMResult = Regex.Replace(LLMResult, "</?RANGEOFCELLS\d*>", "", RegexOptions.IgnoreCase)
                        End If

                        LLMResult = LLMResult.Replace("<FILECONTENT>", "").Replace("</FILECONTENT>", "")

                        If Not String.IsNullOrEmpty(LLMResult) Then
                            LLMResult = Await PostCorrection(LLMResult, UseSecondAPI)
                        End If

                        Dim instructions As New List(Of String)
                        instructions = ParseLLMResponse(LLMResult)

                        If instructions.Count > 0 Then
                            ApplyLLMInstructions(instructions, DoBubbles)
                        End If

                        LineNumber += 1

                    Next

                    If Not ProgressBarModule.CancelOperation Then
                        ProgressBarModule.CancelOperation = True
                        ShowCustomMessageBox($"Batch processing completed (processed {FileCounter} files).")
                    End If

                End If

            Catch ex As Exception
                MessageBox.Show("Error in Range: " & ex.Message)
            End Try

        End If

        Dim result = Globals.Ribbons.Ribbon1.UpdateUndoButton()

        Return True

    End Function

    ''' <summary>
    ''' Prompts the user to select one or more worksheets and gathers their content.
    ''' </summary>
    ''' <param name="includeActiveWorksheet">Whether to include the currently active worksheet in the selection list.</param>
    ''' <returns>String containing the selected worksheets' content wrapped in <RANGEOFCELLS> tags, or "NONE"/"ERROR" on failure.</returns>
    Public Function GatherSelectedWorksheets(
    Optional ByVal includeActiveWorksheet As System.Boolean = False
) As System.String
        Try
            Dim app As Microsoft.Office.Interop.Excel.Application =
            Globals.ThisAddIn.Application
            Dim activeWs As Microsoft.Office.Interop.Excel.Worksheet =
            TryCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

            ' build list of worksheets (optionally including the active one)
            Dim sheetList As New System.Collections.Generic.List(
            Of Microsoft.Office.Interop.Excel.Worksheet)()
            Dim selItems As New System.Collections.Generic.List(Of SelectionItem)()

            For Each wb As Microsoft.Office.Interop.Excel.Workbook In app.Workbooks
                For Each ws As Microsoft.Office.Interop.Excel.Worksheet In wb.Worksheets
                    If includeActiveWorksheet OrElse ws IsNot activeWs Then
                        sheetList.Add(ws)
                        selItems.Add(New SelectionItem(
                        $"{ws.Name} ({wb.FullName})",
                        sheetList.Count))
                    End If
                Next
            Next

            ' if no sheets matched the filter, bail
            If sheetList.Count = 0 Then Return "NONE"

            ' add “All worksheets …” option
            Dim allOptionIndex As System.Int32 = selItems.Count + 1
            Dim allOptionText As System.String = If(
            includeActiveWorksheet,
            "Add all worksheets",
            "Add all other worksheets")
            selItems.Add(New SelectionItem(allOptionText, allOptionIndex))

            Dim itemsArray As SelectionItem() = selItems.ToArray()
            Dim picked As System.Int32 = SelectValue(itemsArray, allOptionIndex, "Choose worksheet to add …")
            If picked < 1 Then Return System.String.Empty

            Dim targets As New System.Collections.Generic.List(
            Of Microsoft.Office.Interop.Excel.Worksheet)()
            If picked = allOptionIndex Then
                targets.AddRange(sheetList)
            Else
                targets.Add(sheetList(picked - 1))
            End If

            Dim InsertedWorksheet As System.String = System.String.Empty
            Dim tagIndex As System.Int32 = 2
            For Each ws As Microsoft.Office.Interop.Excel.Worksheet In targets
                InsertedWorksheet &= $"<RANGEOFCELLS{tagIndex}>" & vbCrLf

                InsertedWorksheet &= ConvertRangeToString(
                CellRange:=CType(ws.UsedRange, Microsoft.Office.Interop.Excel.Range),
                IncludeFormulas:=True,
                DoColor:=False,
                TargetWorksheet:=ws) & vbCrLf

                InsertedWorksheet &= $"</RANGEOFCELLS{tagIndex}>" & vbCrLf
                tagIndex += 1
            Next

            If System.String.IsNullOrEmpty(InsertedWorksheet) Then
                ShowCustomMessageBox("No content could be retrieved from the selected worksheet(s).")
                Return System.String.Empty
            End If

            Return InsertedWorksheet

        Catch ex As System.Exception
            Return "ERROR " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' Converts an Excel range to a structured string representation.
    ''' </summary>
    ''' <param name="CellRange">The Excel range to convert.</param>
    ''' <param name="IncludeFormulas">Whether to include cell formulas in the output.</param>
    ''' <param name="DoColor">Whether to include font and background color information.</param>
    ''' <param name="TargetWorksheet">Optional worksheet to use if CellRange is Nothing.</param>
    ''' <returns>Formatted string containing cell values, formulas, comments, validation options, and optional colors.</returns>
    Public Function ConvertRangeToString(
    ByVal CellRange As Excel.Range,
    ByVal IncludeFormulas As Boolean,
    Optional ByVal DoColor As Boolean = False,
     Optional ByVal TargetWorksheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
) As String

        Dim splash As New SplashScreen("Gathering the content from your worksheet...")
        splash.Show()
        splash.Refresh()

        If CellRange Is Nothing AndAlso TargetWorksheet Is Nothing Then
            splash.Close()
            Return String.Empty
        End If

        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim origWb As Microsoft.Office.Interop.Excel.Workbook = app.ActiveWorkbook
        Dim origWs As Microsoft.Office.Interop.Excel.Worksheet = TryCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

        If TargetWorksheet IsNot Nothing Then
            Dim workbook As Microsoft.Office.Interop.Excel.Workbook = CType(TargetWorksheet.Parent, Microsoft.Office.Interop.Excel.Workbook)
            workbook.Activate()
            TargetWorksheet.Activate()
        End If

        If CellRange Is Nothing AndAlso TargetWorksheet IsNot Nothing Then
            CellRange = TargetWorksheet.UsedRange
        End If

        Dim sb As New System.Text.StringBuilder()
        If TargetWorksheet IsNot Nothing Then
            sb.AppendLine($"From Worksheet: {TargetWorksheet.Name}, File: {CType(TargetWorksheet.Parent, Microsoft.Office.Interop.Excel.Workbook).FullName}")
        Else
            sb.AppendLine($"From Worksheet {CellRange.Worksheet.Name}, File {CType(CellRange.Worksheet.Parent, Microsoft.Office.Interop.Excel.Workbook).FullName}")
        End If

        With app
            .ScreenUpdating = False
            .EnableEvents = False
            .Calculation = Excel.XlCalculation.xlCalculationManual
        End With

        Dim normalStyle As Excel.Style = app.ActiveWorkbook.Styles("Normal")
        Dim defaultFontColor As Long = CLng(normalStyle.Font.Color)
        Dim defaultInteriorColor As Long = CLng(normalStyle.Interior.Color)

        Try
            Dim rawVals As Object = CellRange.Value2
            Dim vals(,) As Object

            If TypeOf rawVals Is Object(,) Then
                vals = CType(rawVals, Object(,))
            Else
                ReDim vals(0, 0)
                vals(0, 0) = rawVals
            End If

            Dim rowLB As Integer = vals.GetLowerBound(0)
            Dim rowUB As Integer = vals.GetUpperBound(0)
            Dim colLB As Integer = vals.GetLowerBound(1)
            Dim colUB As Integer = vals.GetUpperBound(1)

            For r As Integer = rowLB To rowUB
                For c As Integer = colLB To colUB
                    Dim raw = vals(r, c)

                    Dim relativeRow As Integer = r - rowLB + 1
                    Dim relativeCol As Integer = c - colLB + 1
                    Dim cell As Excel.Range = CType(CellRange.Cells(relativeRow, relativeCol), Excel.Range)
                    Dim addr As String = cell.Address(False, False)

                    Dim shouldProcess As Boolean = False

                    If raw IsNot Nothing Then
                        shouldProcess = True
                    End If

                    If Not shouldProcess AndAlso cell.Comment IsNot Nothing Then
                        shouldProcess = True
                    End If

                    If Not shouldProcess Then
                        Try
                            Dim tc = CType(cell, Object).CommentThreaded
                            If tc IsNot Nothing Then shouldProcess = True
                        Catch ex As COMException When ex.ErrorCode = &H800A03EC
                        End Try
                    End If

                    If Not shouldProcess Then
                        Try
                            If cell.Validation.Type = Excel.XlDVType.xlValidateList Then
                                shouldProcess = True
                            End If
                        Catch
                        End Try
                    End If

                    If Not shouldProcess AndAlso DoColor Then
                        Dim hasFill As Boolean = False
                        Try
                            hasFill = (CLng(cell.Interior.ColorIndex) <> Excel.XlColorIndex.xlColorIndexNone)
                        Catch
                        End Try
                        If hasFill Then
                            shouldProcess = True
                        End If
                    End If

                    If shouldProcess Then

                        Try
                            sb.AppendLine($"Cell {addr} has")
                            sb.AppendLine($"- Value {raw}")

                            If IncludeFormulas AndAlso CBool(cell.HasFormula) Then
                                Dim f As String = String.Empty
                                Try
                                    f = cell.Formula2.ToString()
                                Catch comEx As System.Runtime.InteropServices.COMException _
                                When comEx.ErrorCode = &H800A03EC
                                    f = cell.Formula.ToString()
                                End Try
                                sb.AppendLine($"- Formula {If(String.IsNullOrEmpty(f), "none", f)}")
                            End If

                            If cell.Comment IsNot Nothing Then
                                sb.AppendLine($"- Comment {cell.Comment.Text()}")
                            End If

                            Dim cellObj As Object = cell
                            Try
                                Dim topObj As Object = cellObj.CommentThreaded
                                If topObj IsNot Nothing Then
                                    Dim txt = topObj.Text
                                    Dim authName = topObj.Author.Name
                                    sb.AppendLine($"- Threaded comment {txt} (by {authName})")

                                    For Each rep In topObj.Replies
                                        sb.AppendLine($"- Reply comment {rep.Text} (by {rep.Author.Name})")
                                    Next
                                End If
                            Catch ex As System.Runtime.InteropServices.COMException When ex.ErrorCode = &H800A03EC
                            End Try

                            Try
                                Dim hasList As Boolean
                                Try
                                    hasList = (cell.Validation.Type = Excel.XlDVType.xlValidateList)
                                Catch comEx As COMException When comEx.ErrorCode = &H800A03EC
                                    hasList = False
                                End Try

                                If hasList Then
                                    Dim formula1 As String = cell.Validation.Formula1
                                    If Not String.IsNullOrWhiteSpace(formula1) Then
                                        Dim options As New List(Of String)()
                                        Dim wb As Microsoft.Office.Interop.Excel.Workbook = CType(cell.Worksheet.Parent, Microsoft.Office.Interop.Excel.Workbook)
                                        Dim refRange As Excel.Range = Nothing
                                        Dim formulaResolved As Boolean = False

                                        If formula1.StartsWith("="c) Then
                                            ' Handle INDIRECT/INDIREKT formulas first
                                            Dim upperFormula As String = formula1.ToUpperInvariant()
                                            If upperFormula.Contains("INDIRECT") OrElse upperFormula.Contains("INDIREKT") Then
                                                Try
                                                    Dim oldCalc As Excel.XlCalculation = app.Calculation
                                                    app.Calculation = Excel.XlCalculation.xlCalculationAutomatic

                                                    Try
                                                        ' Extract the argument inside INDIREKT(...)
                                                        Dim m As Match = Regex.Match(formula1, "=\s*INDIREKT?\s*\((.+)\)\s*$", RegexOptions.IgnoreCase)
                                                        If m.Success Then
                                                            Dim innerFormula As String = "=" & m.Groups(1).Value

                                                            ' Store original cell state
                                                            Dim originalHasFormula As Boolean = CBool(cell.HasFormula)
                                                            Dim originalFormulaLocal As String = If(originalHasFormula, CStr(cell.FormulaLocal), "")
                                                            Dim originalValue As Object = If(Not originalHasFormula, cell.Value2, Nothing)

                                                            ' Store and temporarily remove data validation
                                                            Dim hadValidation As Boolean = False
                                                            Dim validationType As Integer = 0
                                                            Dim validationFormula1 As String = ""
                                                            Dim validationFormula2 As String = ""
                                                            Dim validationOperator As Integer = 0
                                                            Dim validationAlertStyle As Integer = 0
                                                            Dim validationIgnoreBlank As Boolean = True
                                                            Dim validationInCellDropdown As Boolean = True
                                                            Try
                                                                validationType = cell.Validation.Type
                                                                hadValidation = True
                                                                ' Keep the formula exactly as Excel returns it (in local format)
                                                                validationFormula1 = cell.Validation.Formula1
                                                                Try : validationFormula2 = cell.Validation.Formula2 : Catch : End Try
                                                                validationOperator = cell.Validation.Operator
                                                                validationAlertStyle = cell.Validation.AlertStyle
                                                                validationIgnoreBlank = cell.Validation.IgnoreBlank
                                                                validationInCellDropdown = cell.Validation.InCellDropdown

                                                                Debug.WriteLine($"Captured validation: Type={validationType}, Formula1={validationFormula1}")

                                                                cell.Validation.Delete()
                                                            Catch
                                                            End Try

                                                            Try
                                                                ' Try FormulaLocal first (for localized function names)
                                                                ' But ensure cell is in a state that accepts formulas
                                                                cell.NumberFormat = "General"

                                                                ' Clear any existing content first
                                                                cell.ClearContents()

                                                                ' Set formula - try Formula property first, fall back to FormulaLocal
                                                                Try
                                                                    cell.FormulaLocal = innerFormula
                                                                Catch
                                                                    ' If FormulaLocal fails, the formula might need conversion
                                                                End Try

                                                                ' Force recalculation of the entire worksheet
                                                                cell.Worksheet.Calculate()

                                                                ' Check if it was actually parsed as a formula
                                                                If CBool(cell.HasFormula) Then
                                                                    Dim addressResult As Object = cell.Value2

                                                                    'Debug.WriteLine($"Cell: {cell.Address(False, False)}")
                                                                    'Debug.WriteLine($"HasFormula: {cell.HasFormula}")
                                                                    'Debug.WriteLine($"Inner formula: {innerFormula}")
                                                                    'Debug.WriteLine($"addressResult type: {If(addressResult Is Nothing, "Nothing", addressResult.GetType().ToString())}")
                                                                    'Debug.WriteLine($"addressResult value: {If(addressResult Is Nothing, "Nothing", addressResult.ToString())}")

                                                                    If addressResult IsNot Nothing AndAlso Not IsError(addressResult) Then
                                                                        Dim rangeAddrStr As String = addressResult.ToString().Trim()

                                                                        If Not String.IsNullOrWhiteSpace(rangeAddrStr) AndAlso Not rangeAddrStr.StartsWith("="c) Then
                                                                            If rangeAddrStr.Contains("!") Then
                                                                                Dim parts = rangeAddrStr.Split("!"c)
                                                                                Dim sheetName = parts(0).Trim("'"c)
                                                                                Dim rangeRef = parts(1)

                                                                                Dim targetSheet As Microsoft.Office.Interop.Excel.Worksheet =
                                                                                    CType(wb.Sheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)

                                                                                refRange = targetSheet.Range(rangeRef)
                                                                            Else
                                                                                refRange = cell.Worksheet.Range(rangeAddrStr)
                                                                            End If

                                                                            formulaResolved = (refRange IsNot Nothing)
                                                                        End If
                                                                    End If
                                                                Else
                                                                    Debug.WriteLine($"Formula not parsed - cell.HasFormula = False, cell.Text = {cell.Text}")
                                                                End If
                                                            Finally
                                                                ' Restore original cell content
                                                                Try
                                                                    If originalHasFormula Then
                                                                        cell.FormulaLocal = originalFormulaLocal
                                                                    ElseIf originalValue IsNot Nothing Then
                                                                        cell.ClearContents()
                                                                        cell.Value2 = originalValue
                                                                    Else
                                                                        cell.ClearContents()
                                                                    End If
                                                                Catch
                                                                End Try

                                                                ' Restore data validation
                                                                If hadValidation Then
                                                                    Dim valFormulaEnglish As String = validationFormula1
                                                                    Dim valFormulaLocal As String = validationFormula1

                                                                    ' CRITICAL: Disable alerts. 
                                                                    ' INDIRECT/VLOOKUP validations often evaluate to #N/A during setup. 
                                                                    ' Without this, Excel throws a COM exception instead of accepting the "invalid" state.
                                                                    Dim oldAlerts As Boolean = app.DisplayAlerts
                                                                    app.DisplayAlerts = False

                                                                    Try
                                                                        ' 1. Robust normalization using the cell engine
                                                                        ' Ensure we have both confirmed Local and English versions.
                                                                        If Not String.IsNullOrEmpty(validationFormula1) AndAlso validationFormula1.StartsWith("=") Then
                                                                            Try
                                                                                cell.ClearContents()
                                                                                cell.FormulaLocal = validationFormula1
                                                                                valFormulaEnglish = cell.Formula      ' English (commas)
                                                                                valFormulaLocal = cell.FormulaLocal   ' Local (semicolons, localized names)
                                                                            Catch
                                                                                ' If translation fails, assume capture was correct as-is
                                                                            End Try
                                                                        End If

                                                                        ' 2. Restore Cell Content
                                                                        Try
                                                                            If originalHasFormula Then
                                                                                cell.FormulaLocal = originalFormulaLocal
                                                                            ElseIf originalValue IsNot Nothing Then
                                                                                cell.ClearContents()
                                                                                cell.Value2 = originalValue
                                                                            Else
                                                                                cell.ClearContents()
                                                                            End If
                                                                        Catch
                                                                        End Try

                                                                        ' 3. Restore Validation
                                                                        Try
                                                                            cell.Validation.Delete()
                                                                        Catch
                                                                        End Try

                                                                        If validationType = Excel.XlDVType.xlValidateList Then
                                                                            Try
                                                                                ' Use Placeholder pattern for stability
                                                                                cell.Validation.Add(
                                                                                    Type:=Excel.XlDVType.xlValidateList,
                                                                                    AlertStyle:=CType(validationAlertStyle, Excel.XlDVAlertStyle),
                                                                                    Operator:=Excel.XlFormatConditionOperator.xlBetween,
                                                                                    Formula1:="placeholder")

                                                                                ' OPTIMIZED: Try LOCAL formula first.
                                                                                ' Your environment uses Semicolons/German. Validation.Modify works best with Local syntax.
                                                                                ' DisplayAlerts=False allows it to pass even if the reference is currently #N/A.
                                                                                Try
                                                                                    cell.Validation.Modify(
                                                                                        Type:=Excel.XlDVType.xlValidateList,
                                                                                        AlertStyle:=CType(validationAlertStyle, Excel.XlDVAlertStyle),
                                                                                        Operator:=Excel.XlFormatConditionOperator.xlBetween,
                                                                                        Formula1:=valFormulaLocal)
                                                                                Catch exLocal As Exception
                                                                                    ' Fallback: Try English only if Local failed
                                                                                    Debug.WriteLine($"Modify Local failed ({exLocal.Message}), retrying English: {valFormulaEnglish}")
                                                                                    cell.Validation.Modify(
                                                                                        Type:=Excel.XlDVType.xlValidateList,
                                                                                        AlertStyle:=CType(validationAlertStyle, Excel.XlDVAlertStyle),
                                                                                        Operator:=Excel.XlFormatConditionOperator.xlBetween,
                                                                                        Formula1:=valFormulaEnglish)
                                                                                End Try

                                                                                cell.Validation.IgnoreBlank = validationIgnoreBlank
                                                                                cell.Validation.InCellDropdown = validationInCellDropdown
                                                                            Catch ex As Exception
                                                                                Debug.WriteLine($"List Validation failed completely: {ex.Message}")
                                                                            End Try
                                                                        Else
                                                                            ' Non-list validation
                                                                            Try
                                                                                cell.Validation.Add(
                                                                                    Type:=CType(validationType, Excel.XlDVType),
                                                                                    AlertStyle:=CType(validationAlertStyle, Excel.XlDVAlertStyle),
                                                                                    Operator:=CType(validationOperator, Excel.XlFormatConditionOperator),
                                                                                    Formula1:=valFormulaLocal,
                                                                                    Formula2:=If(String.IsNullOrEmpty(validationFormula2), Type.Missing, validationFormula2))

                                                                                cell.Validation.IgnoreBlank = validationIgnoreBlank
                                                                                cell.Validation.InCellDropdown = validationInCellDropdown
                                                                            Catch
                                                                            End Try
                                                                        End If
                                                                    Catch restoreEx As Exception
                                                                        Debug.WriteLine($"Validation restore failed: {restoreEx.Message}")
                                                                    Finally
                                                                        app.DisplayAlerts = oldAlerts
                                                                    End Try
                                                                End If
                                                            End Try
                                                        End If
                                                    Finally
                                                        app.Calculation = oldCalc
                                                    End Try
                                                Catch ex As Exception
                                                    Debug.WriteLine($"INDIRECT eval failed: {ex.Message}")
                                                End Try
                                            End If

                                            ' ORIGINAL: Try named range lookup
                                            If Not formulaResolved AndAlso refRange Is Nothing Then
                                                Dim nameKey As String = formula1.Substring(1)
                                                Try
                                                    Dim nm As Excel.Name = CType(wb.Names(nameKey), Excel.Name)
                                                    refRange = nm.RefersToRange
                                                Catch ex As Exception
                                                End Try
                                            End If
                                        End If

                                        ' ORIGINAL: Try static range reference
                                        If refRange Is Nothing AndAlso formula1.StartsWith("="c) AndAlso Not formulaResolved Then
                                            Dim addrx As String = formula1.Substring(1)
                                            Try
                                                refRange = cell.Worksheet.Range(addrx)
                                            Catch ex1 As COMException
                                                ' ORIGINAL: Try cross-sheet reference
                                                Dim parts = addrx.Split("!"c)
                                                If parts.Length >= 2 Then
                                                    Dim sheetName = parts(0).Trim("'"c)
                                                    If Not sheetName.StartsWith("[") Then
                                                        Dim rangeAddr = parts(1)
                                                        Try
                                                            Dim otherSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(wb.Sheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                                                            refRange = otherSheet.Range(rangeAddr)
                                                        Catch
                                                        End Try
                                                    End If
                                                End If
                                            End Try
                                        End If

                                        ' ORIGINAL: Extract values from resolved range
                                        If refRange IsNot Nothing Then
                                            Dim tmp = refRange.Value2
                                            If TypeOf tmp Is Object(,) Then
                                                Dim arr = CType(tmp, Object(,))
                                                Dim rCount = arr.GetLength(0)
                                                Dim cCount = arr.GetLength(1)
                                                For rx As Integer = 1 To rCount
                                                    For cx As Integer = 1 To cCount
                                                        Dim v = arr(rx, cx)
                                                        options.Add(If(v Is Nothing, String.Empty, v.ToString()))
                                                    Next
                                                Next
                                            ElseIf tmp IsNot Nothing Then
                                                options.Add(tmp.ToString())
                                            End If
                                            Marshal.ReleaseComObject(refRange)
                                        ElseIf Not formulaResolved Then
                                            ' ORIGINAL: Fallback - split by comma
                                            If formula1.StartsWith("="c) Then
                                                options.AddRange(formula1.Substring(1) _
                                                .Split(","c) _
                                                .Select(Function(s) s.Trim()))
                                            Else
                                                options.AddRange(formula1 _
                                                .Split(","c) _
                                                .Select(Function(s) s.Trim()))
                                            End If
                                        End If

                                        sb.AppendLine($"- Dropdown options (separated by §) {String.Join("§", options)}")
                                    End If
                                End If

                            Catch ex As Exception
                                sb.AppendLine($"- Error reading dropdown {ex.Message}")
                            End Try

                            If DoColor Then
                                If CLng(cell.Font.Color) <> defaultFontColor Then
                                    Try
                                        Dim fc As System.Drawing.Color = System.Drawing.ColorTranslator.FromOle(CInt(cell.Font.Color))
                                        sb.AppendLine($"- FontColor #{fc.R:X2}{fc.G:X2}{fc.B:X2} (rgb {fc.R},{fc.G},{fc.B})")
                                    Catch
                                        sb.AppendLine($"- FontColor {cell.Font.Color}")
                                    End Try
                                End If

                                If CLng(cell.Interior.Color) <> defaultInteriorColor Then
                                    Try
                                        Dim bc As System.Drawing.Color = System.Drawing.ColorTranslator.FromOle(CInt(cell.Interior.Color))
                                        sb.AppendLine($"- BackgroundColor #{bc.R:X2}{bc.G:X2}{bc.B:X2} (rgb {bc.R},{bc.G},{bc.B})")
                                    Catch
                                        sb.AppendLine($"- BackgroundColor {cell.Interior.Color}")
                                    End Try
                                End If
                            End If

                            sb.AppendLine(New String("-"c, 5))

                        Catch ex As System.Runtime.InteropServices.COMException _
                        When ex.ErrorCode = &H800A03EC
                            sb.AppendLine($"- COM-Error in Cell {addr} {ex.Message}")
                        Catch ex As System.Exception
                            sb.AppendLine($"- Error in Cell {addr} {ex.Message}")
                        Finally
                            Marshal.ReleaseComObject(cell)
                        End Try
                    End If
                Next
            Next

        Finally
            With app
                .ScreenUpdating = True
                .EnableEvents = True
                .Calculation = Excel.XlCalculation.xlCalculationAutomatic
            End With

            If origWb IsNot Nothing Then origWb.Activate()
            If origWs IsNot Nothing Then origWs.Activate()

            splash.Close()
        End Try

        Return sb.ToString()
    End Function

End Class