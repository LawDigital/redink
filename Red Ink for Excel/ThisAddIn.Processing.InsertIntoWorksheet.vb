' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.InsertIntoWorksheet.vb
' Purpose: Parses LLM response text for cell directives and applies them to the
'          active worksheet. Supports insertion of formulas, values, and threaded
'          comments with undo state capture and locale-aware formula handling.
'
' Architecture:
'   Parsing:
'     - ParseLLMResponse scans a response string for blocks starting with "[Cell:".
'       Each block is retained if it contains a top-level [Formula:], [Value:], or [Comment:].
'     - GetFormulaOrValueFromInstruction walks bracketed segments at top-level only and
'       extracts formula/value/comment content. Comments are prefixed externally via AN5.
'     - GetCellFromInstruction extracts the cell reference inside "[Cell: <addr>]".
'     - Bracket parsing uses depth tracking (FindMatchingBracket, ExtractBracketContent).
'
'   Application:
'     - ApplyLLMInstructions iterates parsed instruction blocks, resolves target cell,
'       records prior state (CellState) into undoStates (external collection), and applies
'       a formula, numeric value, text value (with cleaning), or threaded comment.
'     - Temporarily disables Excel autocorrect/list expansion/autocomplete features and
'       restores them afterward. Escape key aborts processing.
'     - Uses external helpers: DecodeTextLiterals, CleanExcelFormulaStrings, ShowCustomMessageBox,
'       GetAsyncKeyState, AN / AN5 constants, SplashScreen class.
'
'   Formula Localization:
'     - SetFormulaSafe attempts Formula2, then Formula2Local, then FormulaLocal with localized
'       list separator. If #NAME? persists it calls ConvertFormulaToLocale, replaces separators,
'       and retries. Reports COM or general errors.
'     - ConvertFormulaToLocale opens a temporary workbook, assigns an English formula to A1,
'       reads FormulaLocal, then disposes workbook and releases COM objects.
'
'   Sanitization:
'     - PreSanitizeNonJson converts RTF escape sequences, replaces \line with line breaks,
'       collapses invalid backslashes, trims wrapping single quotes, removes control characters.
'
'   Utilities:
'     - GetActiveWorksheetSafe returns ActiveSheet or first sheet of first open workbook.
'     - SizeOfWorksheet returns total used cell count of ActiveSheet.
'     - ReleaseObject releases COM objects and forces GC.
'
' External Dependencies (by usage, not defined here):
'     - undoStates (collection of CellState)
'     - CellState class
'     - AN / AN5 constants for comment tagging
'     - DecodeTextLiterals, CleanExcelFormulaStrings, ShowCustomMessageBox, GetAsyncKeyState
'     - SplashScreen
' =============================================================================

Option Strict Off           ' Allow for late binding for handling legacy comments
Option Explicit On

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Excel
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Parses an LLM response string and extracts instruction blocks beginning with "[Cell:".
    ''' Retains blocks containing a top-level formula, value, or comment directive.
    ''' </summary>
    ''' <param name="Response">Full response text potentially containing multiple cell directives.</param>
    ''' <returns>List of raw instruction block strings.</returns>
    Public Function ParseLLMResponse(ByVal Response As String) As List(Of String)
        Dim instructions As New List(Of String)()
        Dim startPos As Integer, instructionEnd As Integer
        Dim tempInstruction As String
        Dim cellPattern As String

        ' Ensure we remove any newlines that might affect parsing
        ' Response = Response.Replace(vbCrLf, " ").Replace(vbLf, " ")
        Response = Response.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)

        ' Pattern for finding Cell
        cellPattern = "[Cell:"

        ' Start parsing the response
        startPos = Response.IndexOf(cellPattern, StringComparison.OrdinalIgnoreCase)

        Do While startPos >= 0
            ' Find next cell occurrence to extract the block between this and next [Cell:]
            instructionEnd = Response.IndexOf(cellPattern, startPos + cellPattern.Length, StringComparison.OrdinalIgnoreCase)

            ' If there's no further [Cell:], capture till the end of the string
            If instructionEnd = -1 Then instructionEnd = Response.Length

            ' Extract the instruction block between the current and next [Cell:]
            tempInstruction = Response.Substring(startPos, instructionEnd - startPos)

            ' Only keep blocks that contain a top-level action (ignore nested brackets inside other brackets)
            Dim extracted As String = GetFormulaOrValueFromInstruction(tempInstruction)
            Dim hasAction As Boolean = Not String.IsNullOrWhiteSpace(extracted) OrElse
                                       tempInstruction.IndexOf(AN5 & ":", StringComparison.OrdinalIgnoreCase) >= 0

            If hasAction Then
                instructions.Add(tempInstruction)
            End If

            ' Move to the next instruction start, exit if at the end
            startPos = Response.IndexOf(cellPattern, instructionEnd, StringComparison.OrdinalIgnoreCase)
        Loop

        Return instructions
    End Function

    ''' <summary>
    ''' Returns the active worksheet or the first worksheet of the first open workbook if ActiveSheet is unavailable.
    ''' </summary>
    ''' <param name="app">Excel application instance.</param>
    ''' <returns>Worksheet instance or Nothing if none found.</returns>
    Private Function GetActiveWorksheetSafe(app As Excel.Application) As Worksheet
        ' Try the currently active sheet
        Dim ws = TryCast(app.ActiveSheet, Worksheet)
        If ws IsNot Nothing Then Return ws

        ' Fallback: first worksheet of first open workbook
        For Each wb As Workbook In app.Workbooks
            If wb IsNot Nothing AndAlso wb.Worksheets.Count > 0 Then
                Return CType(wb.Worksheets(1), Worksheet)
            End If
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Applies a list of instruction blocks to the active worksheet. Handles formulas, values, and comments.
    ''' Records prior cell state for undo logic and temporarily disables certain Excel automation features.
    ''' Escape key aborts processing early.
    ''' If the worksheet contains a Liftlock trigger cell, protection is temporarily lifted and restored after processing.
    ''' </summary>
    ''' <param name="instructions">List of parsed instruction blocks.</param>
    ''' <param name="DoAlsoBubbles">Flag currently unused in the conditional branch (retained for interface consistency).</param>
    Sub ApplyLLMInstructions(ByVal instructions As List(Of String), DoAlsoBubbles As Boolean)

        Dim instruction As String
        Dim cellAddress As String
        Dim formulaOrValue As String
        Dim cleanedValue As String
        Dim ii As Integer

        ' Get the active Excel application and sheet
        Dim excelApp As Excel.Application = Globals.ThisAddIn.Application
        Dim activeSheet As Worksheet = GetActiveWorksheetSafe(excelApp)

        If activeSheet Is Nothing Then
            ShowCustomMessageBox("No worksheet available to apply instructions.")
            Exit Sub
        End If

        ' Lift worksheet protection if a Liftlock trigger cell exists
        Dim lockInfo As LiftlockInfo = TryLiftProtection(activeSheet)

        ii = 0

        undoStates.Clear()

        Dim splash As New SplashScreen("Implementing... press 'Esc' to abort")
        splash.Show()
        splash.Refresh()

        Debug.WriteLine("Instructions: " & String.Join(Environment.NewLine, instructions))

        Dim prevAutoFillFormulasInLists As Boolean = excelApp.AutoCorrect.AutoFillFormulasInLists
        Dim prevAutoExpandListRange As Boolean = excelApp.AutoCorrect.AutoExpandListRange
        Dim prevEnableAutoComplete As Boolean = excelApp.EnableAutoComplete
        Dim prevExtendList As Boolean = excelApp.ExtendList

        excelApp.AutoCorrect.AutoFillFormulasInLists = False
        excelApp.AutoCorrect.AutoExpandListRange = False
        excelApp.EnableAutoComplete = False
        excelApp.ExtendList = False

        Try
            ' Loop through the parsed instructions and ask for confirmation before applying
            For Each instruction In instructions

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For
                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then Exit For

                cellAddress = GetCellFromInstruction(instruction)
                formulaOrValue = GetFormulaOrValueFromInstruction(instruction)

                'If Not String.IsNullOrWhiteSpace(cellAddress) AndAlso Not String.IsNullOrWhiteSpace(formulaOrValue) Then
                If Not String.IsNullOrWhiteSpace(cellAddress) Then
                    ii += 1
                    Debug.WriteLine($"Processing: Cell='{cellAddress}', Value='{formulaOrValue}'")

                    Try
                        If activeSheet IsNot Nothing AndAlso activeSheet.Range(cellAddress) IsNot Nothing Then
                            Dim targetRange As Range
                            Try
                                ' Ensure the address is valid before accessing it
                                If Regex.IsMatch(cellAddress, "^[A-Z]+\d+$") Then
                                    targetRange = activeSheet.Range(cellAddress)

                                    ' Store the state BEFORE any changes
                                    Dim state As New CellState With {
                                        .WorksheetName = targetRange.Worksheet.Name,
                                        .CellAddress = targetRange.Address,
                                        .OldValue = targetRange.Value,
                                        .HadFormula = CBool(targetRange.HasFormula),
                                        .OldFormula = If(CBool(targetRange.HasFormula), CStr(targetRange.Formula), "")
                                    }

                                    ' Handle merged cells properly
                                    If targetRange.MergeArea IsNot Nothing Then
                                        targetRange = CType(targetRange.MergeArea.Cells(1, 1), Range)
                                    End If

                                    ' Add the state to undoStates - do this BEFORE making changes
                                    undoStates.Add(state)

                                    ' Strip wrapping quotes around formulas that LLMs sometimes produce
                                    ' e.g. '=SUM(A1:A10)' or "=SUM(A1:A10)"
                                    If formulaOrValue.Length >= 3 Then
                                        If (formulaOrValue.StartsWith("'") AndAlso formulaOrValue.EndsWith("'")) OrElse
                                           (formulaOrValue.StartsWith("""") AndAlso formulaOrValue.EndsWith("""")) Then
                                            Dim inner As String = formulaOrValue.Substring(1, formulaOrValue.Length - 2)
                                            If inner.StartsWith("=") Then
                                                formulaOrValue = inner
                                            End If
                                        End If
                                    End If

                                    If formulaOrValue.StartsWith($"{AN5}: ") Then
                                        'If DoAlsoBubbles And formulaOrValue.StartsWith($"{AN5}: ") Then

                                        ' Add a comment to the cell
                                        Dim commentText As String = formulaOrValue.Trim()
                                        commentText = DecodeTextLiterals(commentText)
                                        If commentText <> $"{AN5}: " Then
                                            If targetRange.CommentThreaded Is Nothing Then
                                                targetRange.AddCommentThreaded(Text:=$"{commentText}")
                                            Else
                                                targetRange.CommentThreaded.AddReply(Text:=$"{commentText}")
                                            End If
                                        End If

                                    ElseIf formulaOrValue.StartsWith("=") Then

                                        ' Fix cell format issues
                                        targetRange.Value = ""
                                        targetRange.NumberFormat = "General"

                                        formulaOrValue = CleanExcelFormulaStrings(formulaOrValue)

                                        SetFormulaSafe(targetRange, formulaOrValue, excelApp)

                                    Else

                                        ' Assign values properly
                                        If IsNumeric(formulaOrValue) Then
                                            targetRange.Value = formulaOrValue
                                        Else
                                            ' Remove unwanted apostrophes
                                            cleanedValue = CStr(formulaOrValue).Trim("'"c)
                                            ' Unescape doubled single quotes produced by LLMs when inside '...'
                                            cleanedValue = cleanedValue.Replace("''", "'")

                                            cleanedValue = DecodeTextLiterals(cleanedValue)
                                            targetRange.NumberFormat = "@"
                                            Debug.WriteLine($"Set cleaned text value in {cellAddress}: '{cleanedValue}'")
                                            targetRange.Value = cleanedValue
                                        End If

                                    End If
                                Else
                                    Debug.WriteLine($"Invalid cell address: {cellAddress}")
                                End If
                            Catch ex As Exception
                                If ex.Message.Contains("HRESULT: 0x800A03EC") Then
                                    ShowCustomMessageBox($"Error: Excel rejected the formula or value '{formulaOrValue}' that {AN} tried to assign to the cell {cellAddress}.")
                                Else
                                    ShowCustomMessageBox($"An error occurred when trying to insert the formula or value '{formulaOrValue}' in cell {cellAddress}: {ex.Message}")
                                End If
                            End Try
                        Else
                            Debug.WriteLine($"Invalid or missing cell address: {cellAddress}")
                        End If
                    Catch ex As Exception
                        If ex.Message.Contains("HRESULT: 0x800A03EC") Then
                            ShowCustomMessageBox($"Error: Excel rejected the formula '{formulaOrValue}' that {AN} tried to assign to the cell {cellAddress}.")
                        Else
                            ShowCustomMessageBox($"An error occurred when trying to insert the formula '{formulaOrValue}' in cell {cellAddress}: {ex.Message}")
                        End If
                    End Try
                End If
            Next

        Finally
            ' --- Always restore Excel settings, even if the loop exits early or errors ---
            excelApp.AutoCorrect.AutoFillFormulasInLists = prevAutoFillFormulasInLists
            excelApp.AutoCorrect.AutoExpandListRange = prevAutoExpandListRange
            excelApp.EnableAutoComplete = prevEnableAutoComplete
            excelApp.ExtendList = prevExtendList

            ' Re-protect worksheet if it was lifted
            ReprotectWorksheet(activeSheet, lockInfo)
        End Try
        splash.Close()

    End Sub


    ''' <summary>
    ''' Sanitizes a raw non-JSON text payload with possible RTF escape sequences and extraneous characters.
    ''' </summary>
    ''' <param name="raw">Input raw string.</param>
    ''' <returns>Cleaned string after transformations.</returns>
    Private Function PreSanitizeNonJson(raw As String) As String
        If String.IsNullOrWhiteSpace(raw) Then Return raw

        ' 1. Convert RTF unicode escapes: \u####?  (decimal to hex → actual char)
        raw = Regex.Replace(raw, "\\u(-?\d+)\?", Function(m)
                                                     Dim dec = Integer.Parse(m.Groups(1).Value)
                                                     Return ChrW(dec)
                                                 End Function)

        ' 2. Replace RTF line markers
        raw = Regex.Replace(raw, "(?i)\\line", vbCrLf)

        ' 3. Collapse excessive backslashes that are not valid JSON escapes (optional)
        ' Leave valid JSON escapes (\n, \r, \t, \") alone
        raw = Regex.Replace(raw, "\\\\(?![nrt""\\/u])", "\")

        ' 4. Strip leading / trailing single quotes the LLM sometimes wraps around payload
        raw = raw.Trim()
        If raw.StartsWith("'") AndAlso raw.EndsWith("'") AndAlso raw.Length > 2 Then
            raw = raw.Substring(1, raw.Length - 2).Trim()
        End If

        ' 5. Remove zero-width or control chars except CR/LF/TAB
        raw = New String(raw.Where(Function(c) AscW(c) >= 32 OrElse c = vbCr OrElse c = vbLf OrElse c = vbTab).ToArray())

        Return raw
    End Function

    ''' <summary>
    ''' Safely assigns a formula to a cell with multi-step locale fallback: Formula2, Formula2Local,
    ''' FormulaLocal with localized list separator, then localized function name conversion.
    ''' Reports #NAME? rejection and COM/general errors.
    ''' </summary>
    ''' <param name="cell">Target Excel.Range.</param>
    ''' <param name="formulaOrValue">English formula string.</param>
    ''' <param name="excelApp">Excel application for locale info.</param>
    Public Sub SetFormulaSafe(cell As Excel.Range, formulaOrValue As String, excelApp As Excel.Application)
        ' 0. Get the list separator (in DE it is ";")
        Dim localSep As String = CStr(excelApp.International(XlApplicationInternational.xlListSeparator))

        ' 1. English base string
        Dim englishFormula As String = formulaOrValue

        Try
            ' 2. First attempt: dynamic-array formula in English
            Try
                cell.Formula2 = englishFormula
            Catch ex1 As System.Runtime.InteropServices.COMException When ex1.ErrorCode = &H800A03EC
                ' 0x800A03EC = locale error
                ' → attempt Formula2Local
                Try
                    cell.Formula2Local = englishFormula
                Catch ex2 As System.Runtime.InteropServices.COMException
                    ' ignore, handled below
                End Try
            End Try

            ' 2b. If the formula was not accepted at all (no formula in cell),
            '     the LLM may have sent a localized formula (e.g. =SUMME on English Excel).
            '     Try FormulaLocal directly with the raw string before further fallbacks.
            If Not CBool(cell.HasFormula) Then
                Try
                    cell.FormulaLocal = englishFormula
                Catch ex2b As System.Runtime.InteropServices.COMException
                    ' ignore, handled below
                End Try
            End If

            ' 3. If #NAME? appears, retry with FormulaLocal and localized separator
            If CBool(cell.HasFormula) AndAlso Trim(cell.Text.ToString()) = "#NAME?" Then
                Try
                    cell.FormulaLocal = englishFormula.Replace(",", localSep)
                Catch ex3 As System.Runtime.InteropServices.COMException
                    ' ignore
                End Try

                ' 4. If still #NAME? → translate names
                If Trim(cell.Text.ToString()) = "#NAME?" Then
                    Dim converted As String = Trim(ConvertFormulaToLocale(englishFormula, excelApp))
                    If Not String.IsNullOrEmpty(converted) Then
                        converted = converted.Replace(",", localSep)
                        Try
                            cell.FormulaLocal = converted
                        Catch ex4 As System.Runtime.InteropServices.COMException
                            ShowCustomMessageBox($"Failed to set converted formula: {ex4.Message}")
                        End Try
                    End If

                    ' 5. Final check
                    If Trim(cell.Text.ToString()) = "#NAME?" Then
                        ShowCustomMessageBox(
                        $"Excel rejected the formula '{englishFormula}' for cell {cell.Address}. Resulted in #NAME?."
                    )
                    End If
                End If
            End If

            ' 6. General COM error
        Catch comEx As System.Runtime.InteropServices.COMException
            ShowCustomMessageBox($"COM Error setting formula: {comEx.Message}")

            ' 7. Other errors
        Catch ex As System.Exception
            ShowCustomMessageBox($"General error setting formula: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Converts an English formula string to a localized variant using a temporary workbook.
    ''' </summary>
    ''' <param name="englishFormula">English formula text.</param>
    ''' <param name="excelApp">Excel application instance.</param>
    ''' <returns>Localized formula or the original English on failure.</returns>
    Public Function ConvertFormulaToLocale(ByVal englishFormula As String, ByVal excelApp As Excel.Application) As String
        Dim wb As Workbook = Nothing
        Dim ws As Worksheet = Nothing
        Dim localizedFormula As String = englishFormula   ' default fallback

        Dim previousScreenUpdating As Boolean = excelApp.ScreenUpdating
        Dim previousDisplayAlerts As Boolean = excelApp.DisplayAlerts

        Try
            excelApp.ScreenUpdating = False
            excelApp.DisplayAlerts = False

            ' Create temporary workbook (safer than adding a sheet to the active workbook)
            wb = excelApp.Workbooks.Add()
            ws = CType(wb.Sheets(1), Worksheet)

            Dim tempRange As Excel.Range = ws.Range("A1")
            tempRange.Formula = englishFormula

            localizedFormula = CStr(tempRange.FormulaLocal)
        Catch
            ' Ignore: fallback (englishFormula) already set
        Finally
            ' Ensure workbook is closed even if an error occurred
            Try
                If wb IsNot Nothing Then wb.Close(SaveChanges:=False)
            Catch
                ' Suppress close errors
            End Try

            ' Restore UI state
            excelApp.DisplayAlerts = previousDisplayAlerts
            excelApp.ScreenUpdating = previousScreenUpdating

            ' Release COM objects after closing
            If ws IsNot Nothing Then ReleaseObject(ws)
            If wb IsNot Nothing Then ReleaseObject(wb)
        End Try

        Return localizedFormula
    End Function

    ''' <summary>
    ''' Returns total number of used cells (rows * columns) in the active sheet.
    ''' </summary>
    ''' <returns>Total cell count or 0 on error.</returns>
    Public Function SizeOfWorksheet() As Integer
        Try
            Dim excelApp As Excel.Application = CType(Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Excel.Application)
            Dim activeSheet As Worksheet = CType(excelApp.ActiveSheet, Worksheet)
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            Dim rowCount As Integer = usedRange.Rows.Count
            Dim colCount As Integer = usedRange.Columns.Count
            Dim totalCells As Integer = rowCount * colCount

            Return totalCells

        Catch ex As System.Exception
            ShowCustomMessageBox("Error in SizeOfWorksheet: " & ex.Message)
        End Try

    End Function

    ''' <summary>
    ''' Releases a COM object reference if it is not Nothing using Marshal.ReleaseComObject,
    ''' sets the reference to Nothing, then forces a garbage collection. Swallows release exceptions.
    ''' </summary>
    ''' <param name="obj">COM object reference variable to release.</param>

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    ''' <summary>
    ''' Extracts the cell reference token from an instruction block that starts with "[Cell: ".
    ''' </summary>
    ''' <param name="instruction">Instruction block.</param>
    ''' <returns>Cell address string or empty string.</returns>
    Function GetCellFromInstruction(ByVal instruction As String) As String
        Dim startPos As Integer = instruction.IndexOf("[Cell: ") + 7
        Dim endPos As Integer = instruction.IndexOf("]", startPos)
        If startPos > 6 AndAlso endPos > startPos Then
            Return instruction.Substring(startPos, endPos - startPos).Trim()
        End If
        Return String.Empty
    End Function

    ''' <summary>
    ''' Extracts either a formula, value, or comment content from a top-level directive inside an instruction block.
    ''' Walks bracket nesting depth to skip unrelated nested segments.
    ''' </summary>
    ''' <param name="instruction">Instruction block string.</param>
    ''' <returns>Extracted formula/value/comment text or empty string.</returns>
    Public Function GetFormulaOrValueFromInstruction(ByVal instruction As String) As String
        If String.IsNullOrEmpty(instruction) Then Return String.Empty

        Dim i As Integer = 0
        While i < instruction.Length
            If instruction(i) = "["c Then
                ' Only consider directives that start at top-level
                If StartsWithAt(instruction, "[Formula: ", i) Then
                    Dim content As String = ExtractBracketContent(instruction, i, "[Formula: ".Length)
                    Return content
                ElseIf StartsWithAt(instruction, "[Value: ", i) Then
                    Dim content As String = ExtractBracketContent(instruction, i, "[Value: ".Length)
                    Return content
                ElseIf StartsWithAt(instruction, "[Comment: ", i) Then
                    Dim content As String = ExtractBracketContent(instruction, i, "[Comment: ".Length)
                    If content IsNot Nothing Then
                        Return $"{AN5}: " & content
                    End If
                    Return String.Empty
                Else
                    ' Some other bracketed segment at top-level: skip it entirely (including nested), then continue
                    Dim closing As Integer = FindMatchingBracket(instruction, i)
                    If closing = -1 Then Exit While
                    i = closing + 1
                    Continue While
                End If
            End If
            i += 1
        End While

        Return String.Empty
    End Function

    ' Helpers used by GetFormulaOrValueFromInstruction

    ''' <summary>
    ''' Determines if a string starts with a given prefix at a specified index using case-insensitive comparison.
    ''' </summary>
    Private Function StartsWithAt(s As String, prefix As String, index As Integer) As Boolean
        If index + prefix.Length > s.Length Then Return False
        Return s.IndexOf(prefix, index, StringComparison.OrdinalIgnoreCase) = index
    End Function

    ''' <summary>
    ''' Extracts content between a top-level opening '[' and its matching closing ']' after a tag prefix length.
    ''' </summary>
    Private Function ExtractBracketContent(s As String, openBracketIndex As Integer, tagLen As Integer) As String
        Dim closeIndex As Integer = FindMatchingBracket(s, openBracketIndex)
        If closeIndex = -1 Then Return String.Empty
        Dim contentStart As Integer = openBracketIndex + tagLen
        Dim len As Integer = closeIndex - contentStart
        If len <= 0 Then Return String.Empty
        Return s.Substring(contentStart, len).Trim()
    End Function

    ''' <summary>
    ''' Finds the matching closing bracket for a '[' starting at openBracketIndex, accounting for nested brackets.
    ''' </summary>
    Private Function FindMatchingBracket(s As String, openBracketIndex As Integer) As Integer
        Dim depth As Integer = 0
        For j As Integer = openBracketIndex To s.Length - 1
            Dim c As Char = s(j)
            If c = "["c Then
                depth += 1
            ElseIf c = "]"c Then
                depth -= 1
                If depth = 0 Then
                    Return j
                End If
            End If
        Next
        Return -1
    End Function

    ' -------------------------------------------------------------------------
    ' Worksheet Protection Lift/Restore via in-cell trigger
    ' -------------------------------------------------------------------------

    ''' <summary>
    ''' <summary>
    ''' Result of scanning the worksheet for a Liftlock trigger cell.
    ''' </summary>
    Private Structure LiftlockInfo
        ''' <summary>True if a trigger cell was found and the sheet is currently protected.</summary>
        Public Found As Boolean
        ''' <summary>Password extracted from the trigger (empty string if none).</summary>
        Public Password As String
        ''' <summary>True if the sheet was actually unprotected by <see cref="TryLiftProtection"/>.</summary>
        Public WasUnprotected As Boolean

        ' --- Captured protection settings (read BEFORE unprotecting) ---
        ''' <summary>Whether drawing objects were protected.</summary>
        Public DrawingObjects As Boolean
        ''' <summary>Whether contents (cells) were protected.</summary>
        Public Contents As Boolean
        ''' <summary>Whether scenarios were protected.</summary>
        Public Scenarios As Boolean
        ''' <summary>Whether formatting cells was allowed.</summary>
        Public AllowFormattingCells As Boolean
        ''' <summary>Whether formatting columns was allowed.</summary>
        Public AllowFormattingColumns As Boolean
        ''' <summary>Whether formatting rows was allowed.</summary>
        Public AllowFormattingRows As Boolean
        ''' <summary>Whether inserting columns was allowed.</summary>
        Public AllowInsertingColumns As Boolean
        ''' <summary>Whether inserting rows was allowed.</summary>
        Public AllowInsertingRows As Boolean
        ''' <summary>Whether inserting hyperlinks was allowed.</summary>
        Public AllowInsertingHyperlinks As Boolean
        ''' <summary>Whether deleting columns was allowed.</summary>
        Public AllowDeletingColumns As Boolean
        ''' <summary>Whether deleting rows was allowed.</summary>
        Public AllowDeletingRows As Boolean
        ''' <summary>Whether sorting was allowed.</summary>
        Public AllowSorting As Boolean
        ''' <summary>Whether auto-filtering was allowed.</summary>
        Public AllowFiltering As Boolean
        ''' <summary>Whether using pivot tables was allowed.</summary>
        Public AllowUsingPivotTables As Boolean
    End Structure

    ''' <summary>
    ''' Scans the active worksheet's UsedRange for a cell whose text value contains
    ''' <c>redink_Liftlock</c> or <c>ri_Liftlock</c> (case-insensitive).
    ''' Supports the forms:
    '''   <c>redink_Liftlock</c>  /  <c>ri_Liftlock</c>          → no password
    '''   <c>redink_Liftlock = myPwd</c>  /  <c>ri_Liftlock = myPwd</c>  → with password
    ''' If found and the worksheet is protected, unprotects it and returns info for later re-protection.
    ''' Uses Range.Find for a fast single-pass search instead of iterating every cell.
    ''' </summary>
    ''' <param name="ws">The worksheet to scan.</param>
    ''' <returns>A <see cref="LiftlockInfo"/> describing the outcome.</returns>
    Private Function TryLiftProtection(ws As Microsoft.Office.Interop.Excel.Worksheet) As LiftlockInfo
        Dim info As New LiftlockInfo With {.Found = False, .Password = "", .WasUnprotected = False}

        If ws Is Nothing Then Return info
        If Not ws.ProtectContents Then Return info  ' Nothing to lift

        ' Capture current protection settings BEFORE unprotecting
        Try
            Dim prot As Excel.Protection = ws.Protection
            info.DrawingObjects = ws.ProtectDrawingObjects
            info.Contents = ws.ProtectContents
            info.Scenarios = ws.ProtectScenarios
            info.AllowFormattingCells = prot.AllowFormattingCells
            info.AllowFormattingColumns = prot.AllowFormattingColumns
            info.AllowFormattingRows = prot.AllowFormattingRows
            info.AllowInsertingColumns = prot.AllowInsertingColumns
            info.AllowInsertingRows = prot.AllowInsertingRows
            info.AllowInsertingHyperlinks = prot.AllowInsertingHyperlinks
            info.AllowDeletingColumns = prot.AllowDeletingColumns
            info.AllowDeletingRows = prot.AllowDeletingRows
            info.AllowSorting = prot.AllowSorting
            info.AllowFiltering = prot.AllowFiltering
            info.AllowUsingPivotTables = prot.AllowUsingPivotTables
        Catch
            ' If we cannot read the settings, defaults (False) will apply
        End Try

        Dim used As Excel.Range = Nothing
        Try
            used = ws.UsedRange
        Catch
            Return info
        End Try
        If used Is Nothing Then Return info

        ' Try both prefixes: "redink_liftlock" (AN2) and "ri_liftlock" (AN5)
        Dim prefixes() As String = {AN2 & "_liftlock", AN5 & "_liftlock"}

        For Each prefix In prefixes
            Dim found As Excel.Range = Nothing
            Try
                found = used.Find(
                    What:=prefix,
                    LookIn:=Excel.XlFindLookIn.xlValues,
                    LookAt:=Excel.XlLookAt.xlPart,
                    SearchOrder:=Excel.XlSearchOrder.xlByRows,
                    SearchDirection:=Excel.XlSearchDirection.xlNext,
                    MatchCase:=False)
            Catch
                ' COM error – skip this prefix
                Continue For
            End Try

            If found IsNot Nothing Then
                info.Found = True

                ' Extract the cell text and parse an optional password
                Dim cellText As String = CStr(found.Value).Trim()

                ' Find the trigger inside the cell text (case-insensitive)
                Dim idx As Integer = cellText.IndexOf(prefix, StringComparison.OrdinalIgnoreCase)
                If idx >= 0 Then
                    Dim remainder As String = cellText.Substring(idx + prefix.Length).Trim()
                    If remainder.StartsWith("=") Then
                        info.Password = remainder.Substring(1).Trim()
                    End If
                End If

                ' Attempt to unprotect
                Try
                    If String.IsNullOrEmpty(info.Password) Then
                        ws.Unprotect()
                    Else
                        ws.Unprotect(info.Password)
                    End If
                    info.WasUnprotected = True
                Catch
                    ' Wrong password or structure-level protection – cannot lift
                    info.WasUnprotected = False
                End Try

                Exit For  ' First match wins
            End If
        Next

        Return info
    End Function

    ''' <summary>
    ''' Re-applies worksheet protection that was previously lifted by <see cref="TryLiftProtection"/>.
    ''' Only acts when the info indicates the sheet was actually unprotected.
    ''' </summary>
    ''' <param name="ws">The worksheet to re-protect.</param>
    ''' <param name="info">The <see cref="LiftlockInfo"/> returned by <see cref="TryLiftProtection"/>.</param>
    Private Sub ReprotectWorksheet(ws As Microsoft.Office.Interop.Excel.Worksheet, info As LiftlockInfo)
        If ws Is Nothing OrElse Not info.WasUnprotected Then Return
        Try
            ws.Protect(
                Password:=If(String.IsNullOrEmpty(info.Password), Type.Missing, info.Password),
                DrawingObjects:=info.DrawingObjects,
                Contents:=info.Contents,
                Scenarios:=info.Scenarios,
                AllowFormattingCells:=info.AllowFormattingCells,
                AllowFormattingColumns:=info.AllowFormattingColumns,
                AllowFormattingRows:=info.AllowFormattingRows,
                AllowInsertingColumns:=info.AllowInsertingColumns,
                AllowInsertingRows:=info.AllowInsertingRows,
                AllowInsertingHyperlinks:=info.AllowInsertingHyperlinks,
                AllowDeletingColumns:=info.AllowDeletingColumns,
                AllowDeletingRows:=info.AllowDeletingRows,
                AllowSorting:=info.AllowSorting,
                AllowFiltering:=info.AllowFiltering,
                AllowUsingPivotTables:=info.AllowUsingPivotTables)
        Catch
            ' Best-effort; avoid surfacing errors for re-protection
        End Try
    End Sub


End Class
