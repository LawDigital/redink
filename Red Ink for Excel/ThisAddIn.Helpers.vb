' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Helpers.vb
' Purpose: Helper functions for the Excel Add-In including cell protection detection,
'          text extraction from ranges, undo functionality for cell modifications,
'          runtime string interpolation, and VBA module validation.
'
' Architecture:
' - GetSelectedText: Extracts text/formulas from a range, respecting cell protection
' - AreAllCellsBlocked/CellProtected: Determines if cells are protected considering
'   worksheet protection, cell lock status, and AllowEditRanges
' - GetNumberOfWords: Counts Unicode word tokens, ignoring purely numeric or symbol entries
' - UndoAction: Restores cell states from undoStates collection using multiple fallback
'   strategies for formulas (Formula, Formula2, FormulaR1C1) and values
' - InterpolateAtRuntime: Replaces placeholders in templates with ThisAddIn field/property values
' - VBAModuleWorking: Validates VBA helper module availability and version
' =============================================================================


Option Strict Off   ' Late binding required for Formula2
Option Explicit On

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Extracts text or formulas from the selected range, excluding protected cells.
    ''' </summary>
    ''' <param name="selectedRange">The range to extract text from.</param>
    ''' <param name="DoFormulas">If True, includes formulas; otherwise only values.</param>
    ''' <returns>Concatenated text from all non-protected cells.</returns>
    Function GetSelectedText(selectedRange As Excel.Range, DoFormulas As Boolean) As String
        Dim selectedTextBuilder As New StringBuilder()

        For Each cell As Excel.Range In selectedRange.Cells
            If Not IsNothing(cell.Value) AndAlso Not CellProtected(cell) Then
                If CBool(cell.HasFormula) Then
                    If DoFormulas Then
                        selectedTextBuilder.AppendLine(CStr(cell.Formula))
                    End If
                Else
                    selectedTextBuilder.AppendLine(CStr(cell.Value))
                End If
            End If
        Next

        Return selectedTextBuilder.ToString()
    End Function

    ''' <summary>
    ''' Determines if all cells in the range are protected and cannot be modified.
    ''' </summary>
    ''' <param name="rng">The range to check.</param>
    ''' <returns>True if all cells are locked and the worksheet is protected; otherwise False.</returns>
    Private Function AreAllCellsBlocked(ByVal rng As Excel.Range) As Boolean
        Dim allLocked As Boolean = True ' Assume all cells are locked by default

        If rng Is Nothing Then Return False

        If rng.Worksheet.ProtectContents Then
            For Each cell As Excel.Range In rng.Cells
                If Not CellProtected(cell) Then
                    allLocked = False
                    Exit For
                End If
            Next
        Else
            allLocked = False
        End If

        Return allLocked
    End Function


    ''' <summary>
    ''' Returns the count of real words: sequences of letters (A–Z, unicode letters) that may include internal apostrophes or hyphens; numeric or symbol-only tokens are ignored.
    ''' </summary>
    ''' <param name="text">Text to evaluate.</param>
    ''' <returns>Real word count.</returns>
    Public Function GetNumberOfWords(text As String) As Integer
        If String.IsNullOrWhiteSpace(text) Then
            Return 0
        End If

        ' Pattern explanation:
        ' \b                 Word boundary
        ' [\p{L}]+           At least one Unicode letter
        ' (?:['’-][\p{L}]+)* Optional groups of (apostrophe or hyphen) + letters (e.g., don't, mother-in-law)
        ' \b                 Word boundary
        ' This excludes tokens containing digits or standalone punctuation.
        Dim pattern As String = "\b[\p{L}]+(?:['’-][\p{L}]+)*\b"

        Return Regex.Matches(text, pattern).Count
    End Function

    ''' <summary>
    ''' Determines if a specific cell is protected considering worksheet protection,
    ''' cell lock status, and AllowEditRanges.
    ''' </summary>
    ''' <param name="cell">The cell to check.</param>
    ''' <returns>True if the cell is effectively protected; otherwise False.</returns>
    Private Function CellProtected(ByVal cell As Excel.Range) As Boolean
        If Not cell.Worksheet.ProtectContents Then
            Return False
        End If

        If Not CBool(cell.Locked) Then
            Return False
        End If

        For Each aer As Excel.AllowEditRange In cell.Worksheet.Protection.AllowEditRanges
            If cell.Application.Intersect(aer.Range, cell) IsNot Nothing Then
                Return False
            End If
        Next

        Return True
    End Function

    ''' <summary>
    ''' Restores cells to their previous state from the undoStates collection.
    ''' Uses multiple fallback strategies for formula restoration and disables
    ''' calculation/events during the operation.
    ''' </summary>
    Public Sub UndoAction()
        Try
            Dim app As Excel.Application = Globals.ThisAddIn.Application
            Dim totalCount As Integer = undoStates.Count
            Dim restoredCount As Integer = 0

            app.ScreenUpdating = False
            app.EnableEvents = False
            app.Calculation = Excel.XlCalculation.xlCalculationManual

            Debug.WriteLine($"Starting undo of {undoStates.Count} states")

            ' Process each saved state to restore the previous value or formula
            For i As Integer = 0 To undoStates.Count - 1
                Dim state = undoStates(i)
                Try
                    Dim ws As Excel.Worksheet = Nothing
                    ' Get worksheet - use error handling to be safe
                    Try
                        ws = CType(app.ActiveWorkbook.Worksheets(CStr(state.WorksheetName)), Excel.Worksheet)
                    Catch wsEx As Exception
                        Debug.WriteLine($"Could not find worksheet {state.WorksheetName}: {wsEx.Message}")
                        Continue For
                    End Try

                    ' Get the range on the worksheet
                    Dim rng As Excel.Range = Nothing
                    Try
                        rng = ws.Range(state.CellAddress)
                    Catch rngEx As Exception
                        Debug.WriteLine($"Failed to get range {state.CellAddress}: {rngEx.Message}")
                        Continue For
                    End Try

                    If rng IsNot Nothing Then
                        Debug.WriteLine($"Processing {i + 1}/{totalCount}: {state.WorksheetName}!{state.CellAddress}")

                        ' First, check if it's in a table
                        Dim isTableCell As Boolean = False
                        Try
                            For Each tbl As Microsoft.Office.Interop.Excel.ListObject In ws.ListObjects
                                If app.Intersect(tbl.Range, rng) IsNot Nothing Then
                                    isTableCell = True
                                    Debug.WriteLine($"  Cell is in table: {tbl.Name}")
                                    Exit For
                                End If
                            Next
                        Catch tableEx As Exception
                            Debug.WriteLine($"  Table check error: {tableEx.Message}")
                        End Try

                        ' Now restore the value using different strategies
                        If state.HadFormula Then
                            Debug.WriteLine($"  Restoring formula: {state.OldFormula}")

                            ' Try multiple approaches for formula restoration
                            Dim success As Boolean = False

                            ' Approach 1: Direct formula setting
                            If Not success Then
                                Try
                                    rng.Formula = state.OldFormula
                                    success = True
                                    Debug.WriteLine("  Set using Formula")
                                Catch ex As Exception
                                    Debug.WriteLine($"  Formula method failed: {ex.Message}")
                                End Try
                            End If

                            ' Approach 2: Formula2 (newer Excel versions)
                            If Not success Then
                                Try
                                    rng.Formula2 = state.OldFormula
                                    success = True
                                    Debug.WriteLine("  Set using Formula2")
                                Catch ex As Exception
                                    Debug.WriteLine($"  Formula2 method failed: {ex.Message}")
                                End Try
                            End If

                            ' Approach 3: FormulaR1C1 as fallback
                            If Not success Then
                                Try
                                    rng.FormulaR1C1 = state.OldFormula
                                    success = True
                                    Debug.WriteLine("  Set using FormulaR1C1")
                                Catch ex As Exception
                                    Debug.WriteLine($"  FormulaR1C1 method failed: {ex.Message}")
                                End Try
                            End If

                            ' Last resort: Set as value
                            If Not success Then
                                Try
                                    rng.Value = state.OldValue
                                    success = True
                                    Debug.WriteLine("  Set using Value (fallback)")
                                Catch ex As Exception
                                    Debug.WriteLine($"  Value fallback failed: {ex.Message}")
                                End Try
                            End If

                            If success Then restoredCount += 1
                        Else
                            Debug.WriteLine($"  Restoring value: {state.OldValue}")
                            Try
                                ' For non-formula cells, just set the value
                                rng.Value = state.OldValue
                                restoredCount += 1
                            Catch ex As Exception
                                Debug.WriteLine($"  Value restore error: {ex.Message}")
                            End Try
                        End If

                        ' Force immediate update of this cell
                        Try
                            rng.Calculate()
                        Catch ex As Exception
                            ' Ignore calculation errors
                        End Try
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error processing state {i}: {ex.Message}")
                End Try

                ' Force periodic UI refresh during long undos
                If i Mod 5 = 0 Then
                    app.ScreenUpdating = True
                    System.Threading.Thread.Sleep(10)
                    app.ScreenUpdating = False
                End If
            Next

            Debug.WriteLine($"Undo complete: {restoredCount}/{totalCount} cells restored")
            undoStates.Clear()
            Dim result = Globals.Ribbons.Ribbon1.UpdateUndoButton()

        Catch ex As System.Exception
            MessageBox.Show("Error during undo: " & ex.Message)
        Finally
            ' Always restore Excel's calculation settings
            Dim app As Excel.Application = Globals.ThisAddIn.Application
            app.ScreenUpdating = True
            app.EnableEvents = True
            app.Calculation = Excel.XlCalculation.xlCalculationAutomatic

            ' Force a full recalculation to ensure all dependencies update
            Try
                app.CalculateFull()
            Catch ex As Exception
                Debug.WriteLine($"Final calculation error: {ex.Message}")
            End Try
        End Try
    End Sub

    ''' <summary>
    ''' Replaces placeholders in a template string with values from ThisAddIn fields and properties.
    ''' Removes specific API-related placeholders before processing.
    ''' </summary>
    ''' <param name="template">The template string containing {PlaceholderName} patterns.</param>
    ''' <returns>The template with placeholders replaced by corresponding values.</returns>
    Public Function InterpolateAtRuntime(ByVal template As String) As String
        If template Is Nothing Then
            MessageBox.Show("Error InterpolateAtRuntime: Template is Nothing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ""
        End If

        template = Regex.Replace(template, "{Codebasis}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_DecodedAPI}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_DecodedAPI_2}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKey}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKeyBack}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKey_2}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKeyBack_2}", "", RegexOptions.IgnoreCase)

        Dim result As String = template

        If Regex.IsMatch(template, "{Dictionary}", RegexOptions.IgnoreCase) Then
            Dictionary = Global.SharedLibrary.SharedLibrary.SharedMethods.GetTranslationDictionaryText(_context)
        End If

        Dim placeholderPattern As String = "\{([^}]+)\}"
        Dim matches As MatchCollection = Regex.Matches(template, placeholderPattern)

        For Each m As Match In matches
            Dim placeholder As String = m.Value          ' e.g. "{Name}"
            Dim varName As String = m.Groups(1).Value    ' e.g. "Name"

            ' Debug.WriteLine($"placeholder = {placeholder}  Varname = {varName}")
            ' Search for Field
            Dim fieldInfo = Me.GetType().GetField(varName)
            If fieldInfo IsNot Nothing Then
                Dim fieldValue = fieldInfo.GetValue(Me)
                If fieldValue IsNot Nothing Then
                    result = result.Replace(placeholder, fieldValue.ToString())
                End If
                Continue For
            End If

            ' Search for Property
            Dim propInfo = Me.GetType().GetProperty(varName)
            If propInfo IsNot Nothing Then
                Dim propValue = propInfo.GetValue(Me)
                If propValue IsNot Nothing Then
                    result = result.Replace(placeholder, propValue.ToString())
                End If
            End If
        Next

        Return result
    End Function

    ''' <summary>
    ''' Verifies that the VBA helper module is available and meets the minimum version requirement.
    ''' </summary>
    ''' <returns>True if VBA module version is >= MinHelperVersion; otherwise False.</returns>
    Public Function VBAModuleWorking() As Boolean
        Dim xlApp As Microsoft.Office.Interop.Excel.Application = Me.Application

        Try
            ' Call the VBA function
            Dim HelperVersion As Integer = CType(xlApp.Run("CheckAppHelper"), Integer)

            If HelperVersion >= MinHelperVersion Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
        End Try

        Return False
    End Function

    ''' <summary>
    ''' Returns True if the worksheet already contains text or values in the intended output strip.
    ''' </summary>
    Private Function WorksheetHasExistingTextInWriteArea(ByVal ws As Excel.Worksheet, ByVal startRow As Integer, ByVal startCol As Integer) As Boolean
        If ws Is Nothing Then Return False

        startRow = Math.Max(1, startRow)
        startCol = Math.Max(1, startCol)

        Try
            Dim usedRange As Excel.Range = ws.UsedRange
            If usedRange Is Nothing Then Return False

            Dim lastUsedRow As Integer = usedRange.Row + usedRange.Rows.Count - 1
            Dim lastUsedCol As Integer = usedRange.Column + usedRange.Columns.Count - 1

            If lastUsedRow < startRow OrElse lastUsedCol < startCol Then Return False

            Dim checkEndCol As Integer = Math.Min(lastUsedCol, startCol + 1)
            Dim checkRange As Excel.Range = CType(ws.Range(ws.Cells(startRow, startCol), ws.Cells(lastUsedRow, checkEndCol)), Excel.Range)
            Dim values As Object = checkRange.Value2

            If values Is Nothing Then Return False

            If TypeOf values Is Object(,) Then
                For Each value As Object In CType(values, Object(,))
                    If value IsNot Nothing AndAlso value.ToString().Trim().Length > 0 Then
                        Return True
                    End If
                Next
            Else
                Return values.ToString().Trim().Length > 0
            End If
        Catch
        End Try

        Return False
    End Function

    ''' <summary>
    ''' Warns when the output area is already occupied and, if requested, creates a new worksheet.
    ''' </summary>
    Private Function PromptForFreshWorksheetIfNeeded(ByRef ws As Excel.Worksheet, ByVal startRow As Integer, ByVal startCol As Integer, ByVal dialogTitle As String) As Boolean
        If ws Is Nothing Then Return False
        If Not WorksheetHasExistingTextInWriteArea(ws, startRow, startCol) Then Return True

        Dim targetCell As String = CType(ws.Cells(startRow, startCol), Excel.Range).Address(False, False)
        Dim answer As Integer = SharedLibrary.SharedLibrary.SharedMethods.ShowCustomYesNoBox(
            "The target output area starting at " & targetCell & " already contains text." & vbCrLf & vbCrLf &
            "This may cause errors." & vbCrLf & vbCrLf &
            "Run the analysis in a new worksheet instead?",
            "Use new worksheet",
            "Keep current worksheet",
            dialogTitle
        )

        If answer <> 1 Then Return True

        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim wb As Excel.Workbook = TryCast(app.ActiveWorkbook, Excel.Workbook)
        If wb Is Nothing Then
            SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("No active workbook was found.")
            Return False
        End If

        Try
            Dim newWs As Excel.Worksheet = CType(wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)), Excel.Worksheet)
            newWs.Activate()
            ws = newWs
            Return True
        Catch ex As Exception
            SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Could not create a new worksheet: " & ex.Message)
            Return False
        End Try
    End Function

End Class
