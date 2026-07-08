' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.ExcelHelpers.vb
' Purpose: Excel helper routines for row height adjustment (including merged cells),
'          legacy comment (note) shape sizing, and multi-pattern regex search/replace
'          across a selected range or entire worksheet.
'
' Architecture:
' - Operates on the active worksheet (Globals.ThisAddIn.Application.ActiveSheet).
' - Each routine acquires the current selection; if empty, user may opt to use UsedRange.
' - SplashScreen shown; ESC key aborts loops (checked via GetAsyncKeyState).
' - AdjustHeight: AutoFits rows, measures required heights (merged cells handled by temporary unmerge and width aggregation), tracks original/max heights, applies final capped height (<= 409).
' - AdjustLegacyNotes: Resizes legacy Comment shapes; constrains width (70–250) and computes height from text length and font size approximation.
' - RegexSearchReplace: Collects multi-line regex patterns and optional replacements, validates all patterns, applies ordered replacements to string cells, counts modifications.
' - Error handling: Each method catches System.Exception and reports via MessageBox.
' =============================================================================

Option Strict Off
Option Explicit On

Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Hosts helper routines used by the Excel add-in for row sizing, legacy note sizing, and regex-based edits.
''' </summary>
Partial Public Class ThisAddIn


    ''' <summary>
    ''' Removes leading "RI: " prefix from all threaded comments (including replies) in the active workbook.
    ''' Shows a progress bar during processing and reports the number of occurrences removed.
    ''' </summary>
    ''' <remarks>
    ''' Iterates through all worksheets and their threaded comments. ESC key or Cancel button aborts processing.
    ''' Uses late binding to support various Excel versions.
    ''' </remarks>
    Public Sub RemoveRIPrefixFromComments()

        Dim RIPrefix As String = $"{AN5}: "
        Dim activeAuthorName As String = String.Empty

        Try
            activeAuthorName = CStr(Globals.ThisAddIn.Application.UserName)
        Catch
        End Try

        If String.IsNullOrWhiteSpace(activeAuthorName) Then
            ShowCustomMessageBox("Unable to determine the active author name.")
            Exit Sub
        End If

        Try
            Dim activeWorkbook As Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
            If activeWorkbook Is Nothing Then
                ShowCustomMessageBox("No active workbook found.")
                Exit Sub
            End If

            Dim activeSheet As Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Worksheet)
            Dim selectionObj As Object = Globals.ThisAddIn.Application.Selection
            Dim selectedRange As Range = TryCast(selectionObj, Range)

            Dim processSelectionOnly As Boolean = selectedRange IsNot Nothing AndAlso selectedRange.Count > 1
            Dim rangeToProcess As Range = If(processSelectionOnly, selectedRange, Nothing)

            Dim candidates As New List(Of (ws As Worksheet, cellAddr As String, isReply As Boolean, replyIndex As Integer, originalText As String))()

            ' Pass 1: Collect only items that actually have the prefix
            If processSelectionOnly Then
                For Each cell As Range In rangeToProcess.Cells
                    Try
                        Dim cellObj As Object = cell
                        Dim topObj As Object = cellObj.CommentThreaded
                        If topObj Is Nothing Then Continue For

                        Dim cellAddr As String = CStr(cell.Address)

                        ' Main comment
                        Dim commentText As String = CStr(topObj.Text)
                        If Not String.IsNullOrEmpty(commentText) AndAlso
                           commentText.StartsWith(RIPrefix, StringComparison.Ordinal) AndAlso
                           IsCommentByActiveAuthor(topObj, activeAuthorName) Then
                            candidates.Add((activeSheet, cellAddr, False, 0, commentText))
                        End If

                        ' Replies
                        Dim replies As Object = Nothing
                        Try
                            replies = topObj.Replies
                        Catch
                        End Try

                        If replies IsNot Nothing Then
                            Dim replyCount As Integer = CInt(replies.Count)
                            For replyIndex As Integer = 1 To replyCount
                                Try
                                    Dim reply As Object = replies(replyIndex)
                                    If reply Is Nothing Then Continue For

                                    Dim replyText As String = CStr(reply.Text)
                                    If Not String.IsNullOrEmpty(replyText) AndAlso
                                       replyText.StartsWith(RIPrefix, StringComparison.Ordinal) AndAlso
                                       IsCommentByActiveAuthor(reply, activeAuthorName) Then
                                        candidates.Add((activeSheet, cellAddr, True, replyIndex, replyText))
                                    End If
                                Catch
                                End Try
                            Next
                        End If
                    Catch ex As COMException When ex.ErrorCode = &H800A03EC
                    Catch
                    End Try
                Next
            Else
                Dim threadedComments As Object = Nothing
                Try
                    threadedComments = CallByName(activeSheet, "CommentsThreaded", CallType.Get)
                Catch
                End Try

                If threadedComments Is Nothing Then
                    ShowCustomMessageBox("No threaded comments collection found on the active worksheet.")
                    Exit Sub
                End If

                Dim commentCount As Integer = 0
                Try
                    commentCount = CInt(CallByName(threadedComments, "Count", CallType.Get))
                Catch
                End Try

                For i As Integer = 1 To commentCount
                    Try
                        Dim topObj As Object = CallByName(threadedComments, "Item", CallType.Get, i)
                        If topObj Is Nothing Then Continue For
                        If Not IsCommentByActiveAuthor(topObj, activeAuthorName) Then Continue For

                        Dim parentObj As Object = CallByName(topObj, "Parent", CallType.Get)
                        Dim cell As Range = CType(parentObj, Range)
                        Dim cellAddr As String = CStr(cell.Address)

                        ' Main comment
                        Dim commentText As String = CStr(topObj.Text)
                        If Not String.IsNullOrEmpty(commentText) AndAlso
                           commentText.StartsWith(RIPrefix, StringComparison.Ordinal) Then
                            candidates.Add((activeSheet, cellAddr, False, 0, commentText))
                        End If

                        ' Replies
                        Dim replies As Object = Nothing
                        Try
                            replies = topObj.Replies
                        Catch
                        End Try

                        If replies IsNot Nothing Then
                            Dim replyCount As Integer = CInt(replies.Count)
                            For replyIndex As Integer = 1 To replyCount
                                Try
                                    Dim reply As Object = replies(replyIndex)
                                    If reply Is Nothing Then Continue For
                                    If Not IsCommentByActiveAuthor(reply, activeAuthorName) Then Continue For

                                    Dim replyText As String = CStr(reply.Text)
                                    If Not String.IsNullOrEmpty(replyText) AndAlso
                                       replyText.StartsWith(RIPrefix, StringComparison.Ordinal) Then
                                        candidates.Add((activeSheet, cellAddr, True, replyIndex, replyText))
                                    End If
                                Catch
                                End Try
                            Next
                        End If
                    Catch ex As COMException When ex.ErrorCode = &H800A03EC
                    Catch
                    End Try
                Next
            End If

            If candidates.Count = 0 Then
                ShowCustomMessageBox($"No '{AN5}:' prefixes found in any threaded comments from the active author.")
                Exit Sub
            End If

            ShowProgressBarInSeparateThread(AN & $" Remove {AN5} Prefix", "Processing comments...")
            ProgressBarModule.CancelOperation = False
            GlobalProgressMax = candidates.Count
            GlobalProgressValue = 0
            GlobalProgressLabel = "Starting..."

            Dim prefixesRemoved As Integer = 0

            ' Pass 2: Apply changes
            For i As Integer = 0 To candidates.Count - 1

                System.Windows.Forms.Application.DoEvents()

                If ProgressBarModule.CancelOperation Then Exit For
                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                    ProgressBarModule.CancelOperation = True
                    Exit For
                End If

                GlobalProgressValue = i + 1
                GlobalProgressLabel = $"Processing {i + 1} of {candidates.Count}..."

                Dim item = candidates(i)
                Dim newText As String = item.originalText.Substring(RIPrefix.Length)

                Try
                    Dim cell As Range = item.ws.Range(item.cellAddr)
                    Dim cellObj As Object = cell
                    Dim topObj As Object = cellObj.CommentThreaded
                    If topObj Is Nothing Then Continue For

                    If item.isReply Then
                        Dim replies As Object = topObj.Replies
                        If replies Is Nothing OrElse item.replyIndex > CInt(replies.Count) Then Continue For

                        Dim reply As Object = replies(item.replyIndex)
                        If reply Is Nothing Then Continue For
                        If Not IsCommentByActiveAuthor(reply, activeAuthorName) Then Continue For

                        ' Try set via method (1-arg), then fallback to 3-arg overwrite
                        CallByName(reply, "Text", CallType.Method, newText)

                        Dim verifyReplyText As String = CStr(reply.Text)
                        If verifyReplyText.StartsWith(RIPrefix, StringComparison.Ordinal) Then
                            CallByName(reply, "Text", CallType.Method, newText, 1, True)
                            verifyReplyText = CStr(reply.Text)
                        End If

                        If Not verifyReplyText.StartsWith(RIPrefix, StringComparison.Ordinal) Then
                            prefixesRemoved += 1
                        End If
                    Else
                        ' Main comment
                        If Not IsCommentByActiveAuthor(topObj, activeAuthorName) Then Continue For

                        CallByName(topObj, "Text", CallType.Method, newText)

                        Dim verifyCommentText As String = CStr(topObj.Text)
                        If verifyCommentText.StartsWith(RIPrefix, StringComparison.Ordinal) Then
                            CallByName(topObj, "Text", CallType.Method, newText, 1, True)
                            verifyCommentText = CStr(topObj.Text)
                        End If

                        If Not verifyCommentText.StartsWith(RIPrefix, StringComparison.Ordinal) Then
                            prefixesRemoved += 1
                        End If
                    End If

                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"Error modifying comment at {item.cellAddr}: {ex.Message}")
                End Try
            Next

            ProgressBarModule.CancelOperation = True

            If prefixesRemoved > 0 Then
                ShowCustomMessageBox($"Removed '{AN5}:' prefix from {prefixesRemoved} comments out of {candidates.Count} matched item(s). Note that this feature only works on comments with the same author name as you have.")
            Else
                ShowCustomMessageBox($"Found {candidates.Count} comment(s) with prefix but could not modify them. Note that this feature only works on comments with the same author name as you have.")
            End If

        Catch ex As System.Exception
            ProgressBarModule.CancelOperation = True
            ShowCustomMessageBox($"Error in RemoveRIPrefixFromComments: {ex.Message}")
        End Try

    End Sub

    Private Function IsCommentByActiveAuthor(ByVal commentObj As Object, ByVal activeAuthorName As String) As Boolean
        If commentObj Is Nothing Then Return False

        Try
            Dim authorObj As Object = CallByName(commentObj, "Author", CallType.Get)
            If authorObj Is Nothing Then Return False

            Dim authorName As String = CStr(CallByName(authorObj, "Name", CallType.Get))
            If String.IsNullOrEmpty(authorName) Then Return False

            Return String.Equals(authorName, activeAuthorName, StringComparison.OrdinalIgnoreCase)
        Catch
        End Try

        Return False
    End Function


    ''' <summary>
    ''' Adjusts row heights for the selected cell range (or entire UsedRange if approved), handling merged cells,
    ''' preserving original heights when larger, and capping at 409 points. ESC aborts processing.
    ''' </summary>
    ''' <param name="Silent">True to suppress prompt when selection is empty; False to ask user.</param>
    ''' <remarks>
    ''' Uses AutoFit, forces WrapText for height calculation, temporarily unmerges horizontally merged cells to aggregate column widths,
    ''' restores widths, then re-merges. Tracks original and maximum computed heights per row.
    ''' </remarks>
    Public Sub AdjustHeight(Optional Silent As Boolean = False)

        Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")

        Try
            ' Get the active Excel worksheet
            Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            ' Get the current selection
            Dim selectedRange As Excel.Range = CType(Globals.ThisAddIn.Application.Selection, Excel.Range)
            selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)

            ' Check if the selection is empty or null
            If selectedRange Is Nothing OrElse selectedRange.Count = 0 Then
                Dim result As Integer = 0
                If Not Silent Then
                    result = ShowCustomYesNoBox("No cells are selected. Would you like to perform the operation on the entire worksheet?", "Yes", "No", "Adjust Height")
                End If
                If result = 1 Then
                    selectedRange = activeSheet.UsedRange
                Else
                    If Not Silent Then ShowCustomMessageBox("Operation cancelled.")
                    Exit Sub
                End If
            End If

            ' Perform AutoFit on the rows of the selected range to ensure initial proper height
            selectedRange.Rows.AutoFit()

            ' Prepare dictionaries for tracking row heights
            Dim rowOriginalHeights As New Dictionary(Of Integer, Double)()
            Dim rowMaxHeights As New Dictionary(Of Integer, Double)()

            ' Initialize these dictionaries for each row in the selection
            For Each oneRow As Excel.Range In selectedRange.Rows
                Dim rowIndex As Integer = oneRow.Row
                Dim currentHeight As Double = CDbl(CType(activeSheet.Rows(rowIndex), Excel.Range).RowHeight)
                rowOriginalHeights(rowIndex) = currentHeight
                ' Start the max at whatever the row is currently
                rowMaxHeights(rowIndex) = currentHeight
            Next

            splash.Show()
            splash.Refresh()

            ' Iterate through each cell in the selection
            For Each cell As Excel.Range In selectedRange

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For
                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then Exit For

                If cell Is Nothing Then Continue For

                ' We'll always enable wrapping so AutoFit will compute multi-line height
                cell.WrapText = True

                Dim wasMerged As Boolean = CBool(cell.MergeCells)
                Dim mergeArea As Excel.Range = If(wasMerged, cell.MergeArea, cell)

                ' Temporarily store the row index for dictionary look-up
                Dim rowIndex As Integer = mergeArea.Row

                ' We'll measure how tall Excel wants to make this cell
                Dim newHeight As Double = 0

                If wasMerged Then
                    ' Store the original column widths for each column
                    Dim firstColIndex As Integer = mergeArea.Column
                    Dim totalCols As Integer = mergeArea.Columns.Count
                    Dim originalWidths As New List(Of Double)

                    For iCol As Integer = 0 To totalCols - 1
                        Dim colWidth As Double = CDbl(CType(activeSheet.Columns(firstColIndex + iCol), Excel.Range).ColumnWidth)
                        originalWidths.Add(colWidth)
                    Next

                    ' Sum the widths so we can set it on the first column after unmerging
                    Dim combinedWidth As Double = originalWidths.Sum()

                    ' Unmerge
                    mergeArea.UnMerge()

                    ' Set only the first column to the combined width so AutoFit sees the "true" width
                    CType(activeSheet.Columns(firstColIndex), Excel.Range).ColumnWidth = combinedWidth

                    ' Autofit (note: must do autofit on entire row(s) that the cell spans)
                    mergeArea.Rows.AutoFit()

                    ' Capture the new row height - handle DBNull for vertically merged cells
                    Dim rowHeightValue As Object = mergeArea.RowHeight
                    If rowHeightValue IsNot Nothing AndAlso Not IsDBNull(rowHeightValue) Then
                        newHeight = CDbl(rowHeightValue)
                    Else
                        ' For vertically merged cells, get height from first row
                        Dim firstRow As Excel.Range = CType(mergeArea.Rows(1), Excel.Range)
                        newHeight = CDbl(firstRow.RowHeight)
                    End If

                    ' Restore the original column widths
                    For iCol As Integer = 0 To totalCols - 1
                        CType(activeSheet.Columns(firstColIndex + iCol), Excel.Range).ColumnWidth = originalWidths(iCol)
                    Next

                    ' Re-merge
                    Dim remergeRange As Excel.Range = CType(activeSheet.Range(
                        CType(activeSheet.Cells(mergeArea.Row, firstColIndex), Excel.Range),
                        CType(activeSheet.Cells(mergeArea.Row, firstColIndex + totalCols - 1), Excel.Range)
                    ), Excel.Range)
                    remergeRange.Merge()

                Else
                    ' If not merged, simply use AutoFit
                    mergeArea.Rows.AutoFit()
                    Dim rowHeightValue As Object = mergeArea.RowHeight
                    If rowHeightValue IsNot Nothing AndAlso Not IsDBNull(rowHeightValue) Then
                        newHeight = CDbl(rowHeightValue)
                    End If
                End If

                ' Store the maximum needed height for this row so far
                If rowMaxHeights.ContainsKey(rowIndex) Then
                    ' Compare existing max with newly measured height
                    If newHeight > rowMaxHeights(rowIndex) Then
                        rowMaxHeights(rowIndex) = newHeight
                    End If
                End If

            Next

            ' Now set each row’s height to the maximum of:
            For Each rowIndex As Integer In rowMaxHeights.Keys.ToList()

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For
                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then Exit For

                Dim finalHeight As Double = Math.Max(rowMaxHeights(rowIndex), rowOriginalHeights(rowIndex))
                If finalHeight > 409 Then finalHeight = 409

                CType(activeSheet.Rows(rowIndex), Excel.Range).RowHeight = finalHeight
            Next

        Catch ex As System.Exception
            MessageBox.Show($"Error in AdjustHeight: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            splash.Close()
        End Try

    End Sub

    ''' <summary>
    ''' Resizes legacy comment (note) shapes in the selected range (or UsedRange if chosen) by constraining width
    ''' and computing required height from text length and font size. ESC aborts processing.
    ''' </summary>
    ''' <remarks>
    ''' Width constrained to 70–250 points; height based on AutoSize minimum and an estimated line height.
    ''' </remarks>
    Public Sub AdjustLegacyNotes()

        Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")

        Try
            ' Get the active Excel worksheet
            Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            ' Get the current selection
            Dim selectedRange As Excel.Range = CType(Globals.ThisAddIn.Application.Selection, Excel.Range)
            selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)

            ' Check if the selection is empty or null
            If selectedRange Is Nothing OrElse selectedRange.Count = 0 Then
                Dim result As Integer = ShowCustomYesNoBox(
                    "No cells are selected. Would you like to perform the operation on the entire worksheet?",
                    "Yes",
                    "No",
                    "Adjust Legacy Notes"
                )

                If result = 1 Then
                    selectedRange = activeSheet.UsedRange
                Else
                    ShowCustomMessageBox("Operation cancelled.")
                    Exit Sub
                End If
            End If

            ' Perform AutoFit on the rows of the selected range to ensure initial proper height
            selectedRange.Rows.AutoFit()

            splash.Show()
            splash.Refresh()

            For Each cell As Excel.Range In selectedRange

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For
                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then Exit For

                If cell Is Nothing Then Continue For

                If cell.Comment IsNot Nothing Then

                    ' Ensure the note box dimensions are at least 70 wide and 20 high, and no more than 250 wide
                    Dim comment As Excel.Comment = cell.Comment
                    With comment.Shape

                        .TextFrame.AutoSize = True
                        Dim MinimumHeight As Double = .Height

                        .TextFrame.AutoSize = False

                        ' Enforce width constraints
                        If .Width < 70 Then .Width = 70
                        If .Width > 250 Then .Width = 250

                        ' Dynamically calculate and set height
                        Dim textLength As Integer = Len(comment.Text)
                        Dim fontSize As Double = CDbl(.TextFrame.Characters.Font.Size)
                        Dim lines As Integer = CInt(Math.Ceiling(textLength / (250 / (fontSize - 2)))) ' Approximation based on average char width
                        Dim lineHeight As Double = fontSize + 2 ' Approximate height per line in points
                        Dim requiredHeight As Double = Math.Max(MinimumHeight, (lines * lineHeight)) + 10

                        If lines > 1 Then .Width = 250

                        .Height = CSng(requiredHeight)

                    End With
                End If

            Next

        Catch ex As System.Exception
            MessageBox.Show($"Error in AdjustLegacyNotes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            splash.Close()
        End Try

    End Sub

    ''' <summary>
    ''' Stores the last entered regex pattern list (multi-line, one pattern per line).
    ''' </summary>
    Private Shared LastRegexPattern As System.String = System.String.Empty

    ''' <summary>
    ''' Stores the last entered regex option flags for reuse.
    ''' </summary>
    Private Shared LastRegexOptions As System.String = System.String.Empty

    ''' <summary>
    ''' Stores the last entered replacement text lines aligned with patterns.
    ''' </summary>
    Private Shared LastRegexReplace As System.String = System.String.Empty

    Private NotInheritable Class RegexLlmSuggestion
        Public Property Patterns As System.Collections.Generic.List(Of System.String)
        Public Property Replacements As System.Collections.Generic.List(Of System.String)
        Public Property Options As System.String
        Public Property Feedback As System.String
        Public Property Warnings As System.Collections.Generic.List(Of System.String)
        Public Property Assumptions As System.Collections.Generic.List(Of System.String)
        Public Property Confidence As System.String
        Public Property RequiresReview As System.Boolean
        Public Property UsedFallback As System.Boolean

        Public Sub New()
            Patterns = New System.Collections.Generic.List(Of System.String)()
            Replacements = Nothing
            Options = System.String.Empty
            Feedback = System.String.Empty
            Warnings = New System.Collections.Generic.List(Of System.String)()
            Assumptions = New System.Collections.Generic.List(Of System.String)()
            Confidence = "low"
            RequiresReview = False
            UsedFallback = False
        End Sub
    End Class

    Private NotInheritable Class RegexPreviewStep
        Public Property Pattern As System.String
        Public Property Replacement As System.String
        Public Property MatchCount As System.Int32
        Public Property ExampleMatches As System.Collections.Generic.List(Of System.String)

        Public Sub New()
            ExampleMatches = New System.Collections.Generic.List(Of System.String)()
        End Sub
    End Class

    Private NotInheritable Class RegexPreviewInfo
        Public Property Steps As System.Collections.Generic.List(Of RegexPreviewStep)
        Public Property TotalMatches As System.Int32
        Public Property FirstMatchAddress As System.String
        Public Property FormulaTextCellsTouched As System.Boolean

        Public Sub New()
            Steps = New System.Collections.Generic.List(Of RegexPreviewStep)()
            FirstMatchAddress = System.String.Empty
            FormulaTextCellsTouched = False
        End Sub
    End Class

    ''' <summary>
    ''' Applies one or more regular expression search/replace operations across the selected range or entire worksheet.
    ''' Manual regex entry remains supported. An empty confirmed pattern input starts the natural-language workflow.
    ''' </summary>
    Public Async Sub RegexSearchReplace()

        Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")

        Try
            Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet =
                CType(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

            Dim usedRange As Excel.Range = activeSheet.UsedRange
            Dim selectedRange As Excel.Range = Nothing

            Try
                selectedRange = CType(Globals.ThisAddIn.Application.Selection, Excel.Range)
            Catch
            End Try

            If selectedRange IsNot Nothing Then
                selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)
            End If

            Dim scopeForLlm As System.String = "selected cells"
            Dim scopeDisplay As System.String = "selected cells"

            If selectedRange Is Nothing OrElse selectedRange.Count = 0 Then
                Dim result As System.Int32 =
                    ShowCustomYesNoBox(
                        "No cells are selected. Would you like to perform the operation on the entire worksheet?",
                        "Yes",
                        "No",
                        "Regex Search & Replace")

                If result = 1 Then
                    selectedRange = activeSheet.UsedRange
                    scopeForLlm = "entire worksheet"
                    scopeDisplay = "entire worksheet"
                Else
                    ShowCustomMessageBox("Operation cancelled.", "Regex Search & Replace")
                    Exit Sub
                End If
            ElseIf selectedRange.Cells.Count = 1 Then
                Dim result As System.Int32 =
                    ShowCustomYesNoBox(
                        "Only a single cell is selected. Would you like to perform the operation on the entire worksheet instead?",
                        "Yes",
                        "No",
                        "Regex Search & Replace")

                If result = 1 Then
                    selectedRange = activeSheet.UsedRange
                    scopeForLlm = "entire worksheet"
                    scopeDisplay = "entire worksheet"
                End If
            End If

            Dim llmSuggestion As RegexLlmSuggestion = Nothing
            Dim patternPrompt As System.String =
                "Step 1: Enter your regex pattern(s), one per line (or leave empty to let the AI generate a regex pattern):"

            Dim patternInput As System.String = ShowCustomInputBox(patternPrompt, "Regex Search & Replace", False, LastRegexPattern)
            If patternInput = "ESC" Then
                Exit Sub
            End If

            If System.String.IsNullOrWhiteSpace(patternInput) Then
                llmSuggestion = Await GetRegexSuggestionForExcel(selectedRange, scopeForLlm, scopeDisplay)
                If llmSuggestion Is Nothing Then
                    Exit Sub
                End If

                Dim suggestedPatterns As System.String = System.String.Join(System.Environment.NewLine, llmSuggestion.Patterns.ToArray())
                patternInput = ShowCustomInputBox(patternPrompt, "Regex Search & Replace", False, suggestedPatterns)
                If patternInput = "ESC" Then
                    Exit Sub
                End If
                If System.String.IsNullOrWhiteSpace(patternInput) Then
                    ShowCustomMessageBox("No regex pattern was confirmed. Aborting without changes.", "Regex Search & Replace")
                    Exit Sub
                End If
            End If

            Dim regexPattern As System.String = patternInput.Trim()
            Dim defaultOptions As System.String = If(llmSuggestion IsNot Nothing, llmSuggestion.Options, LastRegexOptions)
            Dim defaultReplacement As System.String =
                If(llmSuggestion IsNot Nothing,
                   If(llmSuggestion.Replacements IsNot Nothing,
                      System.String.Join(System.Environment.NewLine, llmSuggestion.Replacements.ToArray()),
                      System.String.Empty),
                   LastRegexReplace)

            Dim optionsInputRaw As System.String =
                ShowCustomInputBox(
                    "Step 2: Enter regex option(s) (i for IgnoreCase, m for Multiline, s for Singleline, c for Compiled, r for RightToLeft, e for ExplicitCapture):",
                    "Regex Search & Replace",
                    True,
                    defaultOptions)

            Dim optionsInput As System.String = NormalizeRegexOptionFlags(optionsInputRaw)
            Dim regexOptions As System.Text.RegularExpressions.RegexOptions = ParseRegexOptions(optionsInput)

            Dim replacementPrompt As System.String =
                ShowCustomInputBox(
                    "Step 3: Enter your replacement text(s), one per line, matching your pattern(s)." &
                    System.Environment.NewLine &
                    "Cancel = search only; OK with an empty replacement = delete matches.",
                    "Regex Search & Replace",
                    False,
                    defaultReplacement)

            Dim replacementText As System.String = Nothing
            Dim shouldCacheReplacement As System.Boolean = False

            If replacementPrompt <> "ESC" Then
                replacementText = If(replacementPrompt, System.String.Empty)
                shouldCacheReplacement = True
            End If

            Dim patterns() As System.String =
                regexPattern.Split(New System.String() {System.Environment.NewLine}, System.StringSplitOptions.RemoveEmptyEntries)

            If patterns.Length = 0 Then
                ShowCustomMessageBox("No valid regex patterns were entered. Aborting without changes.", "Regex Search & Replace")
                Exit Sub
            End If

            Dim replacements() As System.String =
                If(replacementText IsNot Nothing,
                   replacementText.Split(New System.String() {System.Environment.NewLine}, System.StringSplitOptions.None),
                   Nothing)

            If replacements IsNot Nothing AndAlso patterns.Length <> replacements.Length Then
                ShowCustomMessageBox(
                    "The number of regex patterns does not match the number of replacement lines. Aborting without changes.",
                    "Regex Search & Replace")
                Exit Sub
            End If

            Dim conflictingStepsMessage As System.String =
                SharedLibrary.SharedLibrary.SharedMethods.GetConflictingRegexStepMessage(patterns, replacements)

            If Not System.String.IsNullOrWhiteSpace(conflictingStepsMessage) Then
                ShowCustomMessageBox(conflictingStepsMessage, "Regex Search & Replace")
                Exit Sub
            End If

            Dim regexes As New System.Collections.Generic.List(Of System.Text.RegularExpressions.Regex)()

            For Each pattern As System.String In patterns
                Try
                    regexes.Add(
                        New System.Text.RegularExpressions.Regex(
                            pattern,
                            regexOptions,
                            System.TimeSpan.FromSeconds(2)))
                Catch ex As System.ArgumentException
                    ShowCustomMessageBox(
                        "The regex pattern is invalid:" &
                        System.Environment.NewLine &
                        pattern &
                        System.Environment.NewLine &
                        System.Environment.NewLine &
                        ex.Message &
                        System.Environment.NewLine &
                        System.Environment.NewLine &
                        "Aborting without changes.",
                        "Regex Search & Replace")
                    Exit Sub
                End Try
            Next

            splash.Show()
            splash.Refresh()

            Dim preview As RegexPreviewInfo = BuildRegexPreviewForRange(selectedRange, regexes, replacements, 3)

            If preview.TotalMatches = 0 Then
                ShowCustomMessageBox(
                    "No matches were found in the selected cells for the current regex pattern(s). Aborting without changes.",
                    "Regex Search & Replace")
                Exit Sub
            End If

            LastRegexPattern = regexPattern
            LastRegexOptions = optionsInput
            If shouldCacheReplacement Then
                LastRegexReplace = replacementText
            End If

            Dim previewMessage As System.String =
                BuildRegexPreviewMessage(scopeDisplay, optionsInput, preview, replacements, llmSuggestion)

            If ShowCustomYesNoBox(
                previewMessage,
                If(replacements Is Nothing, "Search", "Replace"),
                "Cancel",
                "Regex Search & Replace") <> 1 Then
                ShowCustomMessageBox("Operation cancelled. No changes were made.", "Regex Search & Replace")
                Exit Sub
            End If

            If replacements Is Nothing Then
                For Each cell As Excel.Range In selectedRange.Cells
                    System.Windows.Forms.Application.DoEvents()
                    If IsEscPressed() Then Exit For

                    If cell IsNot Nothing AndAlso cell.Value2 IsNot Nothing AndAlso TypeOf cell.Value2 Is System.String Then
                        Dim cellText As System.String = cell.Value2.ToString()

                        For i As System.Int32 = 0 To regexes.Count - 1
                            Dim firstMatch As System.Text.RegularExpressions.Match = regexes(i).Match(cellText)
                            If firstMatch.Success Then
                                cell.Select()
                                ShowCustomMessageBox(
                                    "The search completed." &
                                    System.Environment.NewLine &
                                    "The first matching cell was selected: " & cell.Address(False, False),
                                    "Regex Search & Replace")
                                Exit Sub
                            End If
                        Next
                    End If
                Next

                ShowCustomMessageBox(
                    "No match was found in the " & scopeDisplay & ".",
                    "Regex Search & Replace")
                Exit Sub
            End If

            Dim totalReplacements As System.Int32 = 0
            Dim changedCells As System.Int32 = 0

            For Each cell As Excel.Range In selectedRange.Cells
                System.Windows.Forms.Application.DoEvents()
                If IsEscPressed() Then Exit For

                If cell IsNot Nothing AndAlso cell.Value2 IsNot Nothing AndAlso TypeOf cell.Value2 Is System.String Then
                    Dim originalText As System.String = cell.Value2.ToString()
                    Dim workingText As System.String = originalText

                    For i As System.Int32 = 0 To regexes.Count - 1
                        Dim matches As System.Text.RegularExpressions.MatchCollection = regexes(i).Matches(workingText)
                        totalReplacements += matches.Count

                        If matches.Count > 0 Then
                            Dim localReplacement As System.String = replacements(i)
                            workingText = regexes(i).Replace(
                                workingText,
                                Function(match As System.Text.RegularExpressions.Match) match.Result(localReplacement))
                        End If
                    Next

                    If workingText <> originalText Then
                        changedCells += 1
                        cell.Value2 = workingText
                    End If
                End If
            Next

            Dim resultMessage As New System.Text.StringBuilder()
            resultMessage.AppendLine(totalReplacements.ToString() & " replacement(s) were made in the " & scopeDisplay & ".")
            resultMessage.AppendLine(changedCells.ToString() & " cell(s) were changed.")

            If preview.FormulaTextCellsTouched Then
                resultMessage.AppendLine()
                resultMessage.AppendLine("Note: At least one affected text cell was a formula result. This follows the existing Excel behavior and should be reviewed manually.")
            End If

            ShowCustomMessageBox(resultMessage.ToString(), "Regex Search & Replace")

        Catch ex As System.Text.RegularExpressions.RegexMatchTimeoutException
            ShowCustomMessageBox(
                "The regex operation exceeded the time limit. No further changes were made." &
                System.Environment.NewLine &
                ex.Message,
                "Regex Search & Replace")
        Catch ex As System.Exception
            MessageBox.Show("Error in RegexSearchReplace: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            splash.Close()
        End Try
    End Sub

    Private Async Function GetRegexSuggestionForExcel(targetRange As Excel.Range,
                                                      scopeForLlm As System.String,
                                                      scopeDisplay As System.String) As System.Threading.Tasks.Task(Of RegexLlmSuggestion)
        Dim nlInstruction As System.String =
            ShowCustomInputBox(
                "Please describe, in multiple lines, what you want to search and replace.",
                "Regex Search & Replace",
                False,
                System.String.Empty)

        If nlInstruction = "ESC" Then
            Return Nothing
        End If

        If System.String.IsNullOrWhiteSpace(nlInstruction) Then
            ShowCustomMessageBox("No instruction was entered. Aborting without changes.", "Regex Search & Replace")
            Return Nothing
        End If

        Dim sampleWasTruncated As System.Boolean = False
        Dim sampleText As System.String = BuildExcelRangeSample(targetRange, 6000, sampleWasTruncated)

        Dim suggestion As RegexLlmSuggestion =
            Await RequestRegexSuggestionFromLlm(
                nlInstruction,
                "Excel",
                scopeForLlm,
                sampleText,
                sampleWasTruncated,
                LastRegexPattern,
                LastRegexOptions,
                LastRegexReplace)

        If suggestion Is Nothing Then
            Return Nothing
        End If

        If ShowCustomYesNoBox(
            BuildRegexLlmFeedbackMessage("Excel", scopeDisplay, suggestion),
            "Continue",
            "Abort",
            "Regex Search & Replace") <> 1 Then
            Return Nothing
        End If

        Return suggestion
    End Function

    Private Async Function RequestRegexSuggestionFromLlm(nlInstruction As System.String,
                                                         hostApplicationName As System.String,
                                                         scopeForLlm As System.String,
                                                         sampleText As System.String,
                                                         sampleWasTruncated As System.Boolean,
                                                         lastPattern As System.String,
                                                         lastOptions As System.String,
                                                         lastReplace As System.String) As System.Threading.Tasks.Task(Of RegexLlmSuggestion)
        Try
            Dim prompt As System.String =
                SharedLibrary.SharedLibrary.SharedMethods.BuildRegexLlmRequestPrompt(
                    nlInstruction,
                    hostApplicationName,
                    scopeForLlm,
                    sampleText,
                    sampleWasTruncated,
                    lastPattern,
                    lastOptions,
                    lastReplace,
                    "This feature operates on cell text values only. It does not support matching or filtering by bold, italic, underline, highlight, font color, cell formatting, comments, notes, formulas, or other non-text properties.")

            Dim llmResponse As System.String =
                Await LLM(SP_Regex, prompt, "", "", 0, False, False)

            llmResponse = SharedLibrary.SharedLibrary.WebAgentInterpreter.SanitizeLlmResult(llmResponse)

            Dim suggestion As RegexLlmSuggestion = Nothing
            Dim parseError As System.String = System.String.Empty

            If Not TryParseRegexSuggestion(llmResponse, suggestion, parseError) Then
                ShowCustomMessageBox(
                    "The AI response could not be parsed reliably as JSON for regex suggestions." &
                    System.Environment.NewLine &
                    System.Environment.NewLine &
                    parseError,
                    "Regex Search & Replace")
                Return Nothing
            End If

            If suggestion.Patterns.Count = 0 Then
                ShowCustomMessageBox("The AI did not return any usable regex patterns. Aborting without changes.", "Regex Search & Replace")
                Return Nothing
            End If

            If suggestion.Patterns.Count > 5 Then
                suggestion.Warnings.Add("The AI suggestion contains more than 5 steps. Please review it especially carefully.")
                suggestion.RequiresReview = True
            End If

            If suggestion.Replacements IsNot Nothing AndAlso suggestion.Replacements.Count <> suggestion.Patterns.Count Then
                suggestion.Warnings.Add("The AI returned a different number of replacement lines. The replacement field was not prefilled.")
                suggestion.Replacements = Nothing
                suggestion.RequiresReview = True
            End If

            Return suggestion

        Catch ex As System.Exception
            ShowCustomMessageBox(
                "The AI suggestion could not be created." &
                System.Environment.NewLine &
                ex.Message,
                "Regex Search & Replace")
            Return Nothing
        End Try
    End Function

    Private Function TryParseRegexSuggestion(rawResponse As System.String,
                                             ByRef suggestion As RegexLlmSuggestion,
                                             ByRef errorMessage As System.String) As System.Boolean
        suggestion = Nothing
        errorMessage = System.String.Empty

        If System.String.IsNullOrWhiteSpace(rawResponse) Then
            errorMessage = "Empty AI response."
            Return False
        End If

        Dim jsonText As System.String = rawResponse.Trim()
        Dim usedFallback As System.Boolean = False
        Dim parsedObject As Newtonsoft.Json.Linq.JObject = Nothing

        Try
            parsedObject = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
        Catch
            Dim firstBrace As System.Int32 = jsonText.IndexOf("{"c)
            Dim lastBrace As System.Int32 = jsonText.LastIndexOf("}"c)

            If firstBrace >= 0 AndAlso lastBrace > firstBrace Then
                Dim candidate As System.String = jsonText.Substring(firstBrace, lastBrace - firstBrace + 1)
                Try
                    parsedObject = Newtonsoft.Json.Linq.JObject.Parse(candidate)
                    usedFallback = True
                Catch ex As System.Exception
                    errorMessage = ex.Message
                    Return False
                End Try
            Else
                errorMessage = "No JSON object was found."
                Return False
            End If
        End Try

        Dim parsed As New RegexLlmSuggestion()
        parsed.UsedFallback = usedFallback

        Dim patternsToken As Newtonsoft.Json.Linq.JToken = parsedObject("patterns")
        If patternsToken Is Nothing OrElse patternsToken.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then
            errorMessage = "The 'patterns' field is missing or is not an array."
            Return False
        End If

        For Each item As Newtonsoft.Json.Linq.JToken In patternsToken
            If item IsNot Nothing AndAlso item.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                Dim value As System.String = item.ToString()
                If Not System.String.IsNullOrWhiteSpace(value) Then
                    parsed.Patterns.Add(value)
                End If
            End If
        Next

        If parsed.Patterns.Count = 0 Then
            errorMessage = "The 'patterns' field does not contain any usable patterns."
            Return False
        End If

        Dim replacementsToken As Newtonsoft.Json.Linq.JToken = parsedObject("replacements")
        If replacementsToken Is Nothing OrElse replacementsToken.Type = Newtonsoft.Json.Linq.JTokenType.Null Then
            parsed.Replacements = Nothing
        ElseIf replacementsToken.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
            parsed.Replacements = New System.Collections.Generic.List(Of System.String)()
            For Each item As Newtonsoft.Json.Linq.JToken In replacementsToken
                If item Is Nothing OrElse item.Type = Newtonsoft.Json.Linq.JTokenType.Null Then
                    parsed.Replacements.Add(System.String.Empty)
                Else
                    parsed.Replacements.Add(item.ToString())
                End If
            Next
        Else
            errorMessage = "The 'replacements' field is neither null nor an array."
            Return False
        End If

        parsed.Options = NormalizeRegexOptionFlags(parsedObject.Value(Of System.String)("options"))
        parsed.Feedback = If(parsedObject.Value(Of System.String)("feedback"), System.String.Empty)
        parsed.Confidence = If(parsedObject.Value(Of System.String)("confidence"), "low").Trim().ToLowerInvariant()
        parsed.RequiresReview = SharedLibrary.SharedLibrary.SharedMethods.GetJsonBooleanValue(parsedObject("requiresReview"), True)

        Dim warningsToken As Newtonsoft.Json.Linq.JToken = parsedObject("warnings")
        If warningsToken IsNot Nothing AndAlso warningsToken.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
            For Each item As Newtonsoft.Json.Linq.JToken In warningsToken
                If item IsNot Nothing AndAlso item.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                    parsed.Warnings.Add(item.ToString())
                End If
            Next
        End If

        Dim assumptionsToken As Newtonsoft.Json.Linq.JToken = parsedObject("assumptions")
        If assumptionsToken IsNot Nothing AndAlso assumptionsToken.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
            For Each item As Newtonsoft.Json.Linq.JToken In assumptionsToken
                If item IsNot Nothing AndAlso item.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                    parsed.Assumptions.Add(item.ToString())
                End If
            Next
        End If

        If parsed.UsedFallback Then
            parsed.Warnings.Add("The AI response had to be reduced to a JSON object first. Please review the result manually.")
            parsed.RequiresReview = True
        End If

        suggestion = parsed
        Return True
    End Function

    Private Function BuildExcelRangeSample(targetRange As Excel.Range,
                                           maxLength As System.Int32,
                                           ByRef wasTruncated As System.Boolean) As System.String
        Dim sb As New System.Text.StringBuilder()
        wasTruncated = False

        For Each cell As Excel.Range In targetRange.Cells
            If cell IsNot Nothing AndAlso cell.Value2 IsNot Nothing Then
                Dim cellText As System.String = cell.Value2.ToString()
                If Not System.String.IsNullOrWhiteSpace(cellText) Then
                    Dim entry As System.String = cell.Address(False, False) & ": " & cellText
                    Dim separator As System.String = If(sb.Length = 0, System.String.Empty, System.Environment.NewLine)

                    If sb.Length + separator.Length + entry.Length > maxLength Then
                        Dim remaining As System.Int32 = maxLength - sb.Length - separator.Length
                        If remaining > 0 Then
                            sb.Append(separator)
                            sb.Append(entry.Substring(0, System.Math.Min(remaining, entry.Length)))
                        End If
                        wasTruncated = True
                        Exit For
                    End If

                    sb.Append(separator)
                    sb.Append(entry)
                End If
            End If
        Next

        If sb.Length = 0 Then
            Return "(no text content in scope)"
        End If

        Return sb.ToString()
    End Function

    Private Function BuildRegexPreviewForRange(targetRange As Excel.Range,
                                               regexes As System.Collections.Generic.List(Of System.Text.RegularExpressions.Regex),
                                               replacements() As System.String,
                                               maxExamplesPerStep As System.Int32) As RegexPreviewInfo
        Dim preview As New RegexPreviewInfo()

        For i As System.Int32 = 0 To regexes.Count - 1
            Dim stepInfo As New RegexPreviewStep()
            stepInfo.Pattern = regexes(i).ToString()
            stepInfo.Replacement = If(replacements IsNot Nothing, replacements(i), Nothing)
            preview.Steps.Add(stepInfo)
        Next

        For Each cell As Excel.Range In targetRange.Cells
            If cell IsNot Nothing AndAlso cell.Value2 IsNot Nothing AndAlso TypeOf cell.Value2 Is System.String Then
                Dim workingText As System.String = cell.Value2.ToString()

                If CBool(cell.HasFormula) Then
                    preview.FormulaTextCellsTouched = True
                End If

                For i As System.Int32 = 0 To regexes.Count - 1
                    Dim matches As System.Text.RegularExpressions.MatchCollection = regexes(i).Matches(workingText)
                    preview.Steps(i).MatchCount += matches.Count
                    preview.TotalMatches += matches.Count

                    If System.String.IsNullOrWhiteSpace(preview.FirstMatchAddress) AndAlso matches.Count > 0 Then
                        preview.FirstMatchAddress = cell.Address(False, False)
                    End If

                    Dim remainingExampleSlots As System.Int32 = maxExamplesPerStep - preview.Steps(i).ExampleMatches.Count
                    If remainingExampleSlots > 0 Then
                        Dim exampleCount As System.Int32 = System.Math.Min(matches.Count, remainingExampleSlots)
                        For exampleIndex As System.Int32 = 0 To exampleCount - 1
                            preview.Steps(i).ExampleMatches.Add(PreviewDisplayText(matches(exampleIndex).Value, 120))
                        Next
                    End If

                    If replacements IsNot Nothing AndAlso matches.Count > 0 Then
                        Dim localReplacement As System.String = replacements(i)
                        workingText = regexes(i).Replace(
                            workingText,
                            Function(match As System.Text.RegularExpressions.Match) match.Result(localReplacement))
                    End If
                Next
            End If
        Next

        Return preview
    End Function

    Private Function BuildRegexPreviewMessage(scopeDisplay As System.String,
                                              optionsInput As System.String,
                                              preview As RegexPreviewInfo,
                                              replacements() As System.String,
                                              llmSuggestion As RegexLlmSuggestion) As System.String
        Dim msg As New System.Text.StringBuilder()

        msg.AppendLine("Regex operation preview")
        msg.AppendLine("Scope: " & scopeDisplay)
        msg.AppendLine("Regex steps: " & preview.Steps.Count.ToString())
        msg.AppendLine("Options: " & If(System.String.IsNullOrWhiteSpace(optionsInput), "(none)", optionsInput))
        msg.AppendLine("Total matches: " & preview.TotalMatches.ToString())

        If Not System.String.IsNullOrWhiteSpace(preview.FirstMatchAddress) Then
            msg.AppendLine("First match: " & preview.FirstMatchAddress)
        End If

        For i As System.Int32 = 0 To preview.Steps.Count - 1
            Dim stepInfo As RegexPreviewStep = preview.Steps(i)
            msg.AppendLine()
            msg.AppendLine("Step " & (i + 1).ToString() & ": " & stepInfo.Pattern)
            msg.AppendLine("Replacement: " & DescribeReplacement(stepInfo.Replacement, replacements Is Nothing))
            msg.AppendLine("Matches: " & stepInfo.MatchCount.ToString())

            If stepInfo.ExampleMatches.Count > 0 Then
                msg.AppendLine("Examples:")
                For Each exampleMatch As System.String In stepInfo.ExampleMatches
                    msg.AppendLine("  • " & exampleMatch)
                Next
            End If
        Next

        If preview.FormulaTextCellsTouched Then
            msg.AppendLine()
            msg.AppendLine("Warning: The scope contains text cells that are formula results. This follows the existing Excel behavior and should be reviewed manually.")
        End If

        If llmSuggestion IsNot Nothing Then
            If llmSuggestion.Warnings.Count > 0 Then
                msg.AppendLine()
                msg.AppendLine("Warnings:")
                For Each warning As System.String In llmSuggestion.Warnings
                    msg.AppendLine("  • " & warning)
                Next
            End If

            If llmSuggestion.Assumptions.Count > 0 Then
                msg.AppendLine()
                msg.AppendLine("Assumptions:")
                For Each assumption As System.String In llmSuggestion.Assumptions
                    msg.AppendLine("  • " & assumption)
                Next
            End If

            msg.AppendLine()
            msg.AppendLine("Confidence: " & llmSuggestion.Confidence)
            msg.AppendLine("Manual review required: " & If(llmSuggestion.RequiresReview, "Yes", "No"))
        End If

        msg.AppendLine()
        msg.AppendLine("Continue?")

        Return msg.ToString()
    End Function

    Private Function BuildRegexLlmFeedbackMessage(hostApplicationName As System.String,
                                                  scopeDisplay As System.String,
                                                  suggestion As RegexLlmSuggestion) As System.String
        Dim msg As New System.Text.StringBuilder()

        msg.AppendLine("AI suggestion for Regex Search && Replace")
        msg.AppendLine("Application: " & hostApplicationName)
        msg.AppendLine("Scope: " & scopeDisplay)
        msg.AppendLine("Steps: " & suggestion.Patterns.Count.ToString())
        msg.AppendLine("Options: " & If(System.String.IsNullOrWhiteSpace(suggestion.Options), "(none)", suggestion.Options))
        msg.AppendLine()
        msg.AppendLine("What the regex does:")
        msg.AppendLine(If(System.String.IsNullOrWhiteSpace(suggestion.Feedback), "(no explanation provided)", suggestion.Feedback.Trim()))
        msg.AppendLine()
        msg.AppendLine("What will be replaced:")

        If suggestion.Replacements Is Nothing Then
            msg.AppendLine("Search only; no replacement was proposed.")
        Else
            For i As System.Int32 = 0 To System.Math.Min(suggestion.Patterns.Count, suggestion.Replacements.Count) - 1
                msg.AppendLine("  Step " & (i + 1).ToString() & ": " & DescribeReplacement(suggestion.Replacements(i), False))
            Next
        End If

        If suggestion.Warnings.Count > 0 Then
            msg.AppendLine()
            msg.AppendLine("Warnings:")
            For Each warning As System.String In suggestion.Warnings
                msg.AppendLine("  • " & warning)
            Next
        End If

        If suggestion.Assumptions.Count > 0 Then
            msg.AppendLine()
            msg.AppendLine("Assumptions:")
            For Each assumption As System.String In suggestion.Assumptions
                msg.AppendLine("  • " & assumption)
            Next
        End If

        msg.AppendLine()
        msg.AppendLine("Confidence: " & suggestion.Confidence)
        msg.AppendLine("Manual review required: " & If(suggestion.RequiresReview, "Yes", "No"))

        Return msg.ToString()
    End Function

    Private Function ParseRegexOptions(optionChars As System.String) As System.Text.RegularExpressions.RegexOptions
        Return SharedLibrary.SharedLibrary.SharedMethods.ParseRegexOptionFlags(optionChars)
    End Function

    Private Function NormalizeRegexOptionFlags(optionChars As System.String) As System.String
        Return SharedLibrary.SharedLibrary.SharedMethods.NormalizeRegexOptionFlags(optionChars)
    End Function

    Private Function PreviewDisplayText(value As System.String, maxLength As System.Int32) As System.String
        Return SharedLibrary.SharedLibrary.SharedMethods.PreviewDisplayText(value, maxLength)
    End Function

    Private Function DescribeReplacement(replacement As System.String, searchOnly As System.Boolean) As System.String
        Return SharedLibrary.SharedLibrary.SharedMethods.DescribeRegexReplacement(replacement, searchOnly)
    End Function

    Private Function IsEscPressed() As System.Boolean
        Return (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 OrElse
               (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0
    End Function
End Class
