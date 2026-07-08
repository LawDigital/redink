' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPIlot.Tools.Office.Interop.vb
' Purpose:
'   Outlook-only AutoPilot / Local Agent Excel Interop tools for existing
'   workbooks that must be inspected and completed live rather than through
'   OpenXML/XML-only access.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Partial Public Class ThisAddIn

    Private Structure APExcelLiftlockInfo
        Public Found As Boolean
        Public Password As String
        Public WasUnprotected As Boolean

        Public DrawingObjects As Boolean
        Public Contents As Boolean
        Public Scenarios As Boolean
        Public AllowFormattingCells As Boolean
        Public AllowFormattingColumns As Boolean
        Public AllowFormattingRows As Boolean
        Public AllowInsertingColumns As Boolean
        Public AllowInsertingRows As Boolean
        Public AllowInsertingHyperlinks As Boolean
        Public AllowDeletingColumns As Boolean
        Public AllowDeletingRows As Boolean
        Public AllowSorting As Boolean
        Public AllowFiltering As Boolean
        Public AllowUsingPivotTables As Boolean
    End Structure

    Private Structure APExcelApplyUpdatesResult
        Public AppliedCount As Integer
        Public FailedCount As Integer
        Public SkippedProtectedCount As Integer
        Public Issues As JArray
    End Structure

    Private Async Function ExecuteExcelListLiveWorksheetsTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow
        }

        Try
            ct.ThrowIfCancellationRequested()

            Dim fileName As String = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att As AutoPilotAttachmentInfo = FindAttachment(fileName)
            If att Is Nothing Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' not found."
                Return response
            End If

            If att.IsOverSizeLimit Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' exceeds the size limit."
                Return response
            End If

            context.Log($"Listing live Excel worksheets: {fileName}")
            ApDashboardLog($"📊 Listing live Excel worksheets: {fileName}", "step")

            response = Await SwitchToUi(Function()
                                            Return APExcelListLiveWorksheetsCore(toolCall, att)
                                        End Function)

            If response.Success Then
                ApDashboardLog($"✓ Live Excel worksheets listed: {fileName}", "info")
            Else
                ApDashboardLog($"⚠ Could not list live Excel worksheets: {fileName}", "warn")
            End If
        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error listing live Excel worksheets: {ex.Message}"
        End Try

        Return response
    End Function

    Private Async Function ExecuteExcelReadLiveRangeTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow
        }

        Try
            ct.ThrowIfCancellationRequested()

            Dim fileName As String = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att As AutoPilotAttachmentInfo = FindAttachment(fileName)
            If att Is Nothing Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' not found."
                Return response
            End If

            If att.IsOverSizeLimit Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' exceeds the size limit."
                Return response
            End If

            context.Log($"Reading live Excel range: {fileName}")
            ApDashboardLog($"📊 Reading live Excel range: {fileName}", "step")

            response = Await SwitchToUi(Function()
                                            Return APExcelReadLiveRangeCore(toolCall, att)
                                        End Function)

            If response.Success Then
                ApDashboardLog($"✓ Live Excel range read: {fileName}", "info")
            Else
                ApDashboardLog($"⚠ Could not read live Excel range: {fileName}", "warn")
            End If
        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error reading live Excel range: {ex.Message}"
        End Try

        Return response
    End Function

    Private Async Function ExecuteExcelCompleteLiveWorkbookTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow
        }

        Try
            ct.ThrowIfCancellationRequested()

            Dim fileName As String = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim updatesToken As Object = Nothing
            Dim updates As JArray = Nothing

            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("updates") Then
                updatesToken = toolCall.Arguments("updates")
            End If

            If TypeOf updatesToken Is JArray Then
                updates = DirectCast(updatesToken, JArray)
            End If

            If updates Is Nothing OrElse updates.Count = 0 Then
                response.Success = False
                response.Response = "Missing required parameter: updates (must be a non-empty array)"
                Return response
            End If

            Dim att As AutoPilotAttachmentInfo = FindAttachment(fileName)
            If att Is Nothing Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' not found."
                Return response
            End If

            If att.IsOverSizeLimit Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' exceeds the size limit."
                Return response
            End If

            context.Log($"Completing live Excel workbook: {fileName}")
            ApDashboardLog($"📊 Completing live Excel workbook: {fileName}", "step")

            response = Await SwitchToUi(Function()
                                            Return APExcelCompleteLiveWorkbookCore(toolCall, att)
                                        End Function)

            If response.Success Then
                ApDashboardLog($"✓ Live Excel workbook completed: {fileName}", "info")
            Else
                ApDashboardLog($"⚠ Could not complete live Excel workbook: {fileName}", "warn")
            End If
        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error completing live Excel workbook: {ex.Message}"
        End Try

        Return response
    End Function

    Private Function APExcelListLiveWorksheetsCore(
            toolCall As ToolCall,
            att As AutoPilotAttachmentInfo) As ToolResponse

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow
        }

        Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing

        Try
            excelApp = APExcelCreateApplication()
            wb = APExcelOpenWorkbook(excelApp, att.TempFilePath, True)

            Dim worksheets As New JArray()
            Dim defaultWorksheetName As String = Nothing

            For i As Integer = 1 To wb.Worksheets.Count
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim used As Microsoft.Office.Interop.Excel.Range = Nothing

                Try
                    ws = CType(wb.Worksheets(i), Microsoft.Office.Interop.Excel.Worksheet)

                    If i = 1 Then
                        defaultWorksheetName = ws.Name
                    End If

                    used = ws.UsedRange

                    Dim item As New JObject()
                    item("index") = i
                    item("name") = ws.Name
                    item("visible") = (CLng(ws.Visible) = CLng(Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible))
                    item("used_range") = APExcelGetRangeAddress(used)
                    item("used_row_count") = APExcelGetUsedRowCount(used)
                    item("used_column_count") = APExcelGetUsedColumnCount(used)
                    item("protected") = ws.ProtectContents
                    item("has_liftlock") = APExcelHasLiftlockToken(ws)

                    worksheets.Add(item)
                Finally
                    APExcelReleaseComObject(used)
                    APExcelReleaseComObject(ws)
                End Try
            Next

            Dim payload As New JObject()
            payload("attachment_name") = att.OriginalFileName
            payload("workbook_path") = att.TempFilePath
            payload("worksheet_count") = worksheets.Count
            If defaultWorksheetName Is Nothing Then
                payload("default_worksheet_name") = JValue.CreateNull()
            Else
                payload("default_worksheet_name") = defaultWorksheetName
            End If
            payload("worksheets") = worksheets

            response.Success = True
            response.Response = payload.ToString(Formatting.None)
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error listing live Excel worksheets: {ex.Message}"
        Finally
            SafeCloseExcel(wb, excelApp, True)
        End Try

        Return response
    End Function

    Private Function APExcelReadLiveRangeCore(
            toolCall As ToolCall,
            att As AutoPilotAttachmentInfo) As ToolResponse

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow
        }

        Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim targetRange As Microsoft.Office.Interop.Excel.Range = Nothing

        Try
            Dim worksheetName As String = GetArgString(toolCall.Arguments, "worksheet_name")
            Dim rangeAddress As String = GetArgString(toolCall.Arguments, "range_address")
            Dim includeFormulas As Boolean = GetArgBool(toolCall.Arguments, "include_formulas", False)
            Dim includeColor As Boolean = GetArgBool(toolCall.Arguments, "include_color", True)

            excelApp = APExcelCreateApplication()
            wb = APExcelOpenWorkbook(excelApp, att.TempFilePath, True)

            ws = APExcelFindWorksheet(wb, worksheetName)
            If ws Is Nothing Then
                response.Success = False
                response.Response = If(
                    String.IsNullOrWhiteSpace(worksheetName),
                    "No worksheet found in the workbook.",
                    $"Worksheet '{worksheetName}' was not found.")
                Return response
            End If

            targetRange = APExcelResolveTargetRange(ws, rangeAddress)
            If targetRange Is Nothing Then
                response.Success = False
                response.Response = If(
                    String.IsNullOrWhiteSpace(rangeAddress),
                    $"Could not resolve the used range on worksheet '{ws.Name}'.",
                    $"Range '{rangeAddress}' could not be resolved on worksheet '{ws.Name}'.")
                Return response
            End If

            Dim worksheetProtectedBeforeLift As Boolean = False
            Try
                worksheetProtectedBeforeLift = ws.ProtectContents
            Catch
            End Try

            Dim lockInfo As APExcelLiftlockInfo = APExcelTryLiftProtection(ws)

            Try
                Dim nonWritableCells As JArray = APExcelGetNonWritableCells(targetRange, worksheetProtectedBeforeLift)
                Dim nonWritableCellAddresses As New JArray()

                For Each item As JToken In nonWritableCells
                    Dim obj As JObject = TryCast(item, JObject)
                    If obj Is Nothing Then Continue For

                    Dim cellAddress As String = If(obj("cell")?.ToString(), "")
                    If Not String.IsNullOrWhiteSpace(cellAddress) Then
                        nonWritableCellAddresses.Add(cellAddress)
                    End If
                Next

                Dim payload As New JObject()
                payload("attachment_name") = att.OriginalFileName
                payload("workbook_path") = att.TempFilePath
                payload("worksheet_name") = ws.Name
                payload("used_range_address") = APExcelGetRangeAddress(ws.UsedRange)
                payload("range_address") = APExcelGetRangeAddress(targetRange)
                payload("include_formulas") = includeFormulas
                payload("include_color") = includeColor
                payload("worksheet_protected_before_liftlock") = worksheetProtectedBeforeLift
                payload("liftlock_found") = lockInfo.Found
                payload("liftlock_unprotected") = lockInfo.WasUnprotected
                payload("non_writable_cell_count") = nonWritableCells.Count
                payload("non_writable_cells") = nonWritableCells
                payload("non_writable_cell_addresses") = nonWritableCellAddresses
                payload("content") = APExcelConvertRangeToString(
                    targetRange,
                    includeFormulas,
                    includeColor,
                    worksheetProtectedBeforeLift)

                response.Success = True
                response.Response = payload.ToString(Formatting.None)
            Finally
                APExcelReprotectWorksheet(ws, lockInfo)
            End Try
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error reading live Excel range: {ex.Message}"
        Finally
            APExcelReleaseComObject(targetRange)
            APExcelReleaseComObject(ws)
            SafeCloseExcel(wb, excelApp, True)
        End Try

        Return response
    End Function

    Private Function APExcelCompleteLiveWorkbookCore(
            toolCall As ToolCall,
            att As AutoPilotAttachmentInfo) As ToolResponse

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow
        }

        Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing

        Try
            Dim updates As JArray = DirectCast(toolCall.Arguments("updates"), JArray)
            Dim defaultWorksheetName As String = GetArgString(toolCall.Arguments, "worksheet_name")
            Dim outputPath As String = APExcelBuildCompletedOutputPath(att.OriginalFileName)

            excelApp = APExcelCreateApplication()
            wb = APExcelOpenWorkbook(excelApp, att.TempFilePath, False)

            If String.IsNullOrWhiteSpace(defaultWorksheetName) Then
                Dim firstWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Try
                    firstWs = CType(wb.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                    If firstWs IsNot Nothing Then
                        defaultWorksheetName = firstWs.Name
                    End If
                Finally
                    APExcelReleaseComObject(firstWs)
                End Try
            End If

            Dim updateResult As APExcelApplyUpdatesResult =
                APExcelApplyUpdates(wb, excelApp, updates, defaultWorksheetName)

            Try
                wb.Calculate()
            Catch
            End Try

            Try
                excelApp.CalculateFull()
            Catch
            End Try

            Dim payload As New JObject()
            payload("attachment_name") = att.OriginalFileName
            payload("worksheet_name_default") = If(defaultWorksheetName, "")
            payload("applied_update_count") = updateResult.AppliedCount
            payload("failed_update_count") = updateResult.FailedCount
            payload("skipped_non_writable_count") = updateResult.SkippedProtectedCount
            payload("partial_success") =
                (updateResult.AppliedCount > 0 AndAlso
                 (updateResult.FailedCount > 0 OrElse updateResult.SkippedProtectedCount > 0))
            payload("write_blocked_by_protection") =
                (updateResult.AppliedCount = 0 AndAlso
                 updateResult.FailedCount = 0 AndAlso
                 updateResult.SkippedProtectedCount > 0)
            payload("issues") = updateResult.Issues

            If updateResult.AppliedCount > 0 Then
                Const xlOpenXMLWorkbook As Integer = 51
                wb.SaveAs(outputPath, xlOpenXMLWorkbook)

                Dim registrationTarget As AutoPilotAttachmentInfo = APExcelResolveRegistrationTarget(att)
                If registrationTarget IsNot Nothing Then
                    registrationTarget.OutputFiles.Add(outputPath)
                End If

                payload("output_file") = Path.GetFileName(outputPath)
                payload("output_path") = outputPath
                payload("message") =
                    $"Workbook saved as '{Path.GetFileName(outputPath)}'. Applied {updateResult.AppliedCount} update(s), " &
                    $"{updateResult.FailedCount} failed, {updateResult.SkippedProtectedCount} skipped as non-writable."

                response.Success = True
                response.Response = payload.ToString(Formatting.None)
            Else
                payload("output_file") = JValue.CreateNull()
                payload("output_path") = JValue.CreateNull()
                payload("message") =
                    $"No updates could be applied. {updateResult.FailedCount} failed and " &
                    $"{updateResult.SkippedProtectedCount} were skipped as non-writable."

                response.Success = False
                response.ErrorMessage = CStr(payload("message"))
                response.Response = payload.ToString(Formatting.None)
            End If
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error completing live Excel workbook: {ex.Message}"
        Finally
            SafeCloseExcel(wb, excelApp, True)
        End Try

        Return response
    End Function

    Private Function APExcelCreateApplication() As Microsoft.Office.Interop.Excel.Application
        Dim app As New Microsoft.Office.Interop.Excel.Application()

        app.Visible = False
        app.DisplayAlerts = False
        app.ScreenUpdating = False

        Try
            app.EnableEvents = False
        Catch
        End Try

        Try
            app.AskToUpdateLinks = False
        Catch
        End Try

        Try
            app.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic
        Catch
        End Try

        Return app
    End Function

    Private Function APExcelOpenWorkbook(
            app As Microsoft.Office.Interop.Excel.Application,
            path As String,
            isReadOnly As Boolean) As Microsoft.Office.Interop.Excel.Workbook

        Return app.Workbooks.Open(
            Filename:=path,
            UpdateLinks:=0,
            ReadOnly:=isReadOnly,
            AddToMru:=False,
            IgnoreReadOnlyRecommended:=True)
    End Function

    Private Function APExcelFindWorksheet(
            wb As Microsoft.Office.Interop.Excel.Workbook,
            worksheetName As String) As Microsoft.Office.Interop.Excel.Worksheet

        If wb Is Nothing OrElse wb.Worksheets.Count = 0 Then
            Return Nothing
        End If

        If String.IsNullOrWhiteSpace(worksheetName) Then
            Return CType(wb.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
        End If

        For i As Integer = 1 To wb.Worksheets.Count
            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing

            Try
                ws = CType(wb.Worksheets(i), Microsoft.Office.Interop.Excel.Worksheet)
                If ws IsNot Nothing AndAlso
                   ws.Name.Equals(worksheetName.Trim(), StringComparison.OrdinalIgnoreCase) Then
                    Return ws
                End If
            Catch
            End Try

            APExcelReleaseComObject(ws)
        Next

        Return Nothing
    End Function

    Private Function APExcelResolveTargetRange(
            ws As Microsoft.Office.Interop.Excel.Worksheet,
            rangeAddress As String) As Microsoft.Office.Interop.Excel.Range

        If ws Is Nothing Then
            Return Nothing
        End If

        If String.IsNullOrWhiteSpace(rangeAddress) Then
            Return ws.UsedRange
        End If

        Try
            Return ws.Range(rangeAddress.Trim())
        Catch
            Return Nothing
        End Try
    End Function

    Private Function APExcelBuildCompletedOutputPath(originalFileName As String) As String
        Dim baseName As String = Path.GetFileNameWithoutExtension(originalFileName)
        Dim outputName As String = baseName & "_completed.xlsx"
        Dim outputPath As String = Path.Combine(_apCurrentTempDir, outputName)
        Dim counter As Integer = 1

        While File.Exists(outputPath)
            outputName = $"{baseName}_completed_{counter}.xlsx"
            outputPath = Path.Combine(_apCurrentTempDir, outputName)
            counter += 1
        End While

        Return outputPath
    End Function

    Private Function APExcelResolveRegistrationTarget(att As AutoPilotAttachmentInfo) As AutoPilotAttachmentInfo
        If att Is Nothing Then
            Return Nothing
        End If

        If Not att.IsToolOutput Then
            Return att
        End If

        If _apCurrentAttachments IsNot Nothing Then
            For Each item As AutoPilotAttachmentInfo In _apCurrentAttachments
                If item Is Nothing OrElse item.OutputFiles Is Nothing Then Continue For

                For Each outputPath As String In item.OutputFiles
                    If String.IsNullOrWhiteSpace(outputPath) Then Continue For

                    If Path.GetFileName(outputPath).Equals(att.OriginalFileName, StringComparison.OrdinalIgnoreCase) Then
                        Return item
                    End If
                Next
            Next

            If _apCurrentAttachments.Count > 0 Then
                Return _apCurrentAttachments(0)
            End If
        End If

        Return Nothing
    End Function

    Private Function APExcelApplyUpdates(
            wb As Microsoft.Office.Interop.Excel.Workbook,
            excelApp As Microsoft.Office.Interop.Excel.Application,
            updates As JArray,
            defaultWorksheetName As String) As APExcelApplyUpdatesResult

        Dim result As New APExcelApplyUpdatesResult With {
            .AppliedCount = 0,
            .FailedCount = 0,
            .SkippedProtectedCount = 0,
            .Issues = New JArray()
        }

        For Each token As JToken In updates
            If Not TypeOf token Is JObject Then
                Continue For
            End If

            Dim updateObj As JObject = DirectCast(token, JObject)
            Dim cellAddress As String = APExcelGetJsonString(updateObj, "cell")
            Dim worksheetName As String = APExcelGetJsonString(updateObj, "worksheet_name")

            If String.IsNullOrWhiteSpace(worksheetName) Then
                worksheetName = defaultWorksheetName
            End If

            Dim issue As New JObject()
            issue("cell") = If(cellAddress, "")
            issue("worksheet_name") = If(worksheetName, "")

            If String.IsNullOrWhiteSpace(cellAddress) Then
                issue("status") = "failed"
                issue("message") = "Each update requires a 'cell' value."
                result.FailedCount += 1
                result.Issues.Add(issue)
                Continue For
            End If

            If Not Regex.IsMatch(cellAddress, "^[A-Za-z]+\d+$") Then
                issue("status") = "failed"
                issue("message") = $"Invalid cell address '{cellAddress}'."
                result.FailedCount += 1
                result.Issues.Add(issue)
                Continue For
            End If

            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim targetRange As Microsoft.Office.Interop.Excel.Range = Nothing
            Dim mergeArea As Microsoft.Office.Interop.Excel.Range = Nothing
            Dim anchorRange As Microsoft.Office.Interop.Excel.Range = Nothing
            Dim worksheetProtectedBeforeLift As Boolean = False
            Dim lockInfo As APExcelLiftlockInfo = Nothing
            Dim valueOrFormulaApplied As Boolean = False
            Dim commentApplied As Boolean = False

            Try
                ws = APExcelFindWorksheet(wb, worksheetName)
                If ws Is Nothing Then
                    Throw New InvalidOperationException($"Worksheet '{worksheetName}' was not found.")
                End If

                Try
                    worksheetProtectedBeforeLift = ws.ProtectContents
                Catch
                End Try

                lockInfo = APExcelTryLiftProtection(ws)
                targetRange = ws.Range(cellAddress)

                If CBool(targetRange.MergeCells) Then
                    mergeArea = targetRange.MergeArea
                    anchorRange = CType(mergeArea.Cells(1, 1), Microsoft.Office.Interop.Excel.Range)
                Else
                    anchorRange = targetRange
                End If

                issue("cell") = APExcelGetRangeAddress(anchorRange)

                Dim formulaText As String = APExcelNormalizeFormulaText(APExcelGetJsonString(updateObj, "formula"))
                Dim commentText As String = APExcelGetJsonString(updateObj, "comment")
                Dim hasValue As Boolean = (updateObj.Property("value") IsNot Nothing)
                Dim valueToken As JToken = updateObj("value")

                If String.IsNullOrWhiteSpace(formulaText) AndAlso
                   String.IsNullOrWhiteSpace(commentText) AndAlso
                   Not hasValue Then
                    Throw New InvalidOperationException(
                        $"Update for cell '{cellAddress}' on worksheet '{worksheetName}' does not contain value, formula, or comment.")
                End If

                Try
                    If Not String.IsNullOrWhiteSpace(formulaText) Then
                        anchorRange.Value = ""
                        anchorRange.NumberFormat = "General"

                        Dim formulaError As String = Nothing
                        If Not APExcelSetFormulaSafe(anchorRange, formulaText, excelApp, formulaError) Then
                            Throw New System.InvalidOperationException(
                If(formulaError, $"Excel rejected the formula '{formulaText}' for {worksheetName}!{cellAddress}."))
                        End If

                        valueOrFormulaApplied = True
                    ElseIf hasValue Then
                        If valueToken Is Nothing OrElse valueToken.Type = JTokenType.Null Then
                            anchorRange.ClearContents()
                        Else
                            Select Case valueToken.Type
                                Case JTokenType.Integer
                                    anchorRange.Value2 = valueToken.Value(Of Long)()
                                Case JTokenType.Float
                                    anchorRange.Value2 = valueToken.Value(Of Double)()
                                Case JTokenType.Boolean
                                    anchorRange.Value = valueToken.Value(Of Boolean)()
                                Case JTokenType.Date
                                    anchorRange.Value = valueToken.Value(Of DateTime)()
                                Case Else
                                    anchorRange.NumberFormat = "@"
                                    anchorRange.Value = valueToken.ToString()
                            End Select
                        End If

                        valueOrFormulaApplied = True
                    End If

                    If Not String.IsNullOrWhiteSpace(commentText) Then
                        Dim commentError As String = Nothing

                        If APExcelTryAddComment(anchorRange, commentText, commentError) Then
                            commentApplied = True
                        ElseIf valueOrFormulaApplied Then
                            issue("status") = "partial"
                            issue("message") =
                $"Cell '{APExcelGetRangeAddress(anchorRange)}' on worksheet '{worksheetName}' was updated, but the comment could not be added: {commentError}"
                            result.AppliedCount += 1
                            result.FailedCount += 1
                            result.Issues.Add(issue)
                            Continue For
                        Else
                            Throw New System.InvalidOperationException(
                $"Comment could not be added to cell '{APExcelGetRangeAddress(anchorRange)}' on worksheet '{worksheetName}': {commentError}")
                        End If
                    End If

                Catch ex As System.Exception When APExcelIsProtectionWriteFailure(ex, ws)
                    issue("status") = "skipped_non_writable"
                    issue("message") =
        $"Excel rejected write to cell '{APExcelGetRangeAddress(anchorRange)}' on worksheet '{worksheetName}' while worksheet protection applies."
                    issue("excel_error") = ex.Message
                    issue("requested_cell") = cellAddress
                    issue("write_cell") = APExcelGetRangeAddress(anchorRange)
                    result.SkippedProtectedCount += 1
                    result.Issues.Add(issue)
                    Continue For
                End Try

                If valueOrFormulaApplied OrElse commentApplied Then
                    issue("status") = "applied"
                    issue("message") =
                        $"Applied update to cell '{APExcelGetRangeAddress(anchorRange)}' on worksheet '{worksheetName}'."
                    result.AppliedCount += 1
                    result.Issues.Add(issue)
                Else
                    issue("status") = "failed"
                    issue("message") =
                        $"No change was applied to cell '{APExcelGetRangeAddress(anchorRange)}' on worksheet '{worksheetName}'."
                    result.FailedCount += 1
                    result.Issues.Add(issue)
                End If
            Catch ex As Exception
                issue("status") = "failed"
                issue("message") = ex.Message
                result.FailedCount += 1
                result.Issues.Add(issue)
            Finally
                APExcelReprotectWorksheet(ws, lockInfo)
                APExcelReleaseComObject(anchorRange)
                APExcelReleaseComObject(mergeArea)
                APExcelReleaseComObject(targetRange)
                APExcelReleaseComObject(ws)
            End Try
        Next

        Return result
    End Function


    Private Function APExcelIsProtectionWriteFailure(
        ex As System.Exception,
        ws As Microsoft.Office.Interop.Excel.Worksheet) As Boolean

        If ex Is Nothing Then
            Return False
        End If

        Dim comEx As System.Runtime.InteropServices.COMException =
        TryCast(ex, System.Runtime.InteropServices.COMException)

        If comEx Is Nothing Then
            Return False
        End If

        If comEx.ErrorCode <> &H800A03EC Then
            Return False
        End If

        Try
            If ws IsNot Nothing AndAlso ws.ProtectContents Then
                Return True
            End If
        Catch
        End Try

        Return False
    End Function

    Private Sub APExcelAddThreadedComment(targetRange As Microsoft.Office.Interop.Excel.Range, commentText As String)
        Dim cellObj As Object = targetRange

        Try
            If cellObj.CommentThreaded Is Nothing Then
                targetRange.AddCommentThreaded(Text:=commentText)
            Else
                cellObj.CommentThreaded.AddReply(Text:=commentText)
            End If
        Catch
            Try
                If targetRange.Comment Is Nothing Then
                    targetRange.AddComment(commentText)
                Else
                    targetRange.Comment.Text(Text:=commentText, Start:=1, Overwrite:=False)
                End If
            Catch
            End Try
        End Try
    End Sub

    Private Function APExcelConvertRangeToString(
            cellRange As Microsoft.Office.Interop.Excel.Range,
            includeFormulas As Boolean,
            doColor As Boolean,
            sheetProtectedBeforeLift As Boolean) As String

        If cellRange Is Nothing Then
            Return String.Empty
        End If

        Dim sb As New StringBuilder()
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing

        Try
            ws = CType(cellRange.Worksheet, Microsoft.Office.Interop.Excel.Worksheet)
            wb = CType(ws.Parent, Microsoft.Office.Interop.Excel.Workbook)

            sb.AppendLine($"From Worksheet: {ws.Name}, File: {wb.FullName}")

            If sheetProtectedBeforeLift Then
                Dim nonWritableCells As JArray = APExcelGetNonWritableCells(cellRange, True)

                sb.AppendLine($"Worksheet protection before lift-lock: yes")
                sb.AppendLine($"Non-writable cells after re-protection: {nonWritableCells.Count}")

                If nonWritableCells.Count > 0 Then
                    sb.AppendLine("Non-writable cells:")

                    For Each item As JToken In nonWritableCells
                        Dim obj As JObject = TryCast(item, JObject)
                        If obj Is Nothing Then Continue For

                        Dim cellAddress As String = If(obj("cell")?.ToString(), "")
                        Dim mergedArea As String = If(obj("merged_area")?.ToString(), "")

                        If String.IsNullOrWhiteSpace(mergedArea) Then
                            sb.AppendLine($"- {cellAddress}")
                        Else
                            sb.AppendLine($"- {cellAddress} (merged area {mergedArea})")
                        End If
                    Next
                End If

                sb.AppendLine()
            End If

            Dim rawVals As Object = cellRange.Value2
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
                    Dim raw As Object = vals(r, c)
                    Dim relativeRow As Integer = r - rowLB + 1
                    Dim relativeCol As Integer = c - colLB + 1
                    Dim cell As Microsoft.Office.Interop.Excel.Range = Nothing

                    Try
                        cell = CType(cellRange.Cells(relativeRow, relativeCol), Microsoft.Office.Interop.Excel.Range)

                        Dim addr As String = cell.Address(False, False)
                        Dim shouldProcess As Boolean = (raw IsNot Nothing)
                        Dim hasCustomFontColor As Boolean = False
                        Dim hasCustomFillColor As Boolean = False

                        If Not shouldProcess AndAlso cell.Comment IsNot Nothing Then
                            shouldProcess = True
                        End If

                        If Not shouldProcess Then
                            Try
                                Dim threaded As Object = CType(cell, Object).CommentThreaded
                                If threaded IsNot Nothing Then
                                    shouldProcess = True
                                End If
                            Catch ex As COMException When ex.ErrorCode = &H800A03EC
                            End Try
                        End If

                        If Not shouldProcess Then
                            Try
                                If cell.Validation.Type = Microsoft.Office.Interop.Excel.XlDVType.xlValidateList Then
                                    shouldProcess = True
                                End If
                            Catch
                            End Try
                        End If

                        If doColor Then
                            Try
                                hasCustomFontColor = (CLng(cell.Font.ColorIndex) <> CLng(Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic))
                            Catch
                            End Try

                            Try
                                hasCustomFillColor = (CLng(cell.Interior.ColorIndex) <> CLng(Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone))
                            Catch
                            End Try

                            If Not shouldProcess AndAlso (hasCustomFontColor OrElse hasCustomFillColor) Then
                                shouldProcess = True
                            End If
                        End If

                        If shouldProcess Then
                            sb.AppendLine($"Cell {addr} has")
                            sb.AppendLine($"- Value {APExcelValueToInvariantString(raw)}")

                            If sheetProtectedBeforeLift Then
                                sb.AppendLine($"- Locked flag after re-protection {If(APExcelCanWriteToCell(cell, True), "not locked", "locked or ambiguous")}")
                            End If

                            If includeFormulas AndAlso CBool(cell.HasFormula) Then
                                Dim formulaText As String = String.Empty

                                Try
                                    formulaText = CStr(cell.Formula2)
                                Catch ex As COMException When ex.ErrorCode = &H800A03EC
                                    Try
                                        formulaText = CStr(cell.Formula)
                                    Catch
                                        formulaText = String.Empty
                                    End Try
                                End Try

                                sb.AppendLine($"- Formula {If(String.IsNullOrWhiteSpace(formulaText), "none", formulaText)}")
                            End If

                            If cell.Comment IsNot Nothing Then
                                Try
                                    sb.AppendLine($"- Comment {cell.Comment.Text()}")
                                Catch
                                End Try
                            End If

                            Try
                                Dim threaded As Object = CType(cell, Object).CommentThreaded
                                If threaded IsNot Nothing Then
                                    Try
                                        sb.AppendLine($"- Threaded comment {threaded.Text} (by {threaded.Author.Name})")
                                    Catch
                                    End Try

                                    Try
                                        For Each reply As Object In threaded.Replies
                                            Try
                                                sb.AppendLine($"- Reply comment {reply.Text} (by {reply.Author.Name})")
                                            Catch
                                            End Try
                                        Next
                                    Catch
                                    End Try
                                End If
                            Catch ex As COMException When ex.ErrorCode = &H800A03EC
                            End Try

                            Dim options As List(Of String) = APExcelGetValidationOptions(cell)
                            If options.Count > 0 Then
                                sb.AppendLine($"- Dropdown options {String.Join(" | ", options)}")
                            End If

                            If doColor Then
                                If hasCustomFontColor Then
                                    Try
                                        sb.AppendLine($"- Font color {APExcelOleColorToHex(CInt(cell.Font.Color))}")
                                    Catch
                                    End Try
                                End If

                                If hasCustomFillColor Then
                                    Try
                                        sb.AppendLine($"- Background color {APExcelOleColorToHex(CInt(cell.Interior.Color))}")
                                    Catch
                                    End Try
                                End If
                            End If
                        End If
                    Finally
                        APExcelReleaseComObject(cell)
                    End Try
                Next
            Next
        Finally
            APExcelReleaseComObject(wb)
            APExcelReleaseComObject(ws)
        End Try

        Return sb.ToString().TrimEnd()
    End Function

    Private Function APExcelGetNonWritableCells(
            cellRange As Microsoft.Office.Interop.Excel.Range,
            sheetProtectedBeforeLift As Boolean) As JArray

        Dim result As New JArray()

        If cellRange Is Nothing OrElse Not sheetProtectedBeforeLift Then
            Return result
        End If

        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each cellObj As Object In cellRange.Cells
            Dim cell As Microsoft.Office.Interop.Excel.Range = Nothing
            Dim mergeArea As Microsoft.Office.Interop.Excel.Range = Nothing
            Dim anchor As Microsoft.Office.Interop.Excel.Range = Nothing

            Try
                cell = CType(cellObj, Microsoft.Office.Interop.Excel.Range)

                Dim key As String = Nothing
                Dim isLocked As Boolean = False
                Dim mergedAreaAddress As String = Nothing

                If CBool(cell.MergeCells) Then
                    mergeArea = cell.MergeArea
                    anchor = CType(mergeArea.Cells(1, 1), Microsoft.Office.Interop.Excel.Range)
                    key = APExcelGetRangeAddress(mergeArea)
                    mergedAreaAddress = key

                    Try
                        isLocked = CBool(anchor.Locked)
                    Catch
                    End Try
                Else
                    key = APExcelGetRangeAddress(cell)

                    Try
                        isLocked = CBool(cell.Locked)
                    Catch
                    End Try
                End If

                If isLocked AndAlso Not String.IsNullOrWhiteSpace(key) AndAlso seen.Add(key) Then
                    Dim item As New JObject()
                    item("cell") = If(anchor IsNot Nothing, APExcelGetRangeAddress(anchor), APExcelGetRangeAddress(cell))
                    item("reason") = "locked while worksheet protection applies"

                    If Not String.IsNullOrWhiteSpace(mergedAreaAddress) Then
                        item("merged_area") = mergedAreaAddress
                    End If

                    result.Add(item)
                End If
            Finally
                APExcelReleaseComObject(anchor)
                APExcelReleaseComObject(mergeArea)
                APExcelReleaseComObject(cell)
            End Try
        Next

        Return result
    End Function

    Private Function APExcelCanWriteToCell(
            targetRange As Microsoft.Office.Interop.Excel.Range,
            sheetProtectedBeforeLift As Boolean) As Boolean

        If targetRange Is Nothing Then
            Return False
        End If

        If Not sheetProtectedBeforeLift Then
            Return True
        End If

        Dim mergeArea As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim anchor As Microsoft.Office.Interop.Excel.Range = Nothing

        Try
            If CBool(targetRange.MergeCells) Then
                mergeArea = targetRange.MergeArea
                anchor = CType(mergeArea.Cells(1, 1), Microsoft.Office.Interop.Excel.Range)
            Else
                anchor = targetRange
            End If

            Try
                Return Not CBool(anchor.Locked)
            Catch
                Return True
            End Try
        Finally
            If anchor IsNot Nothing AndAlso Not Object.ReferenceEquals(anchor, targetRange) Then
                APExcelReleaseComObject(anchor)
            End If

            APExcelReleaseComObject(mergeArea)
        End Try
    End Function

    Private Function APExcelTryAddComment(
            targetRange As Microsoft.Office.Interop.Excel.Range,
            commentText As String,
            ByRef errorMessage As String) As Boolean

        Dim cellObj As Object = targetRange

        Try
            If cellObj.CommentThreaded Is Nothing Then
                targetRange.AddCommentThreaded(Text:=commentText)
            Else
                cellObj.CommentThreaded.AddReply(Text:=commentText)
            End If

            Return True
        Catch exThreaded As Exception
            Try
                If targetRange.Comment Is Nothing Then
                    targetRange.AddComment(commentText)
                Else
                    targetRange.Comment.Text(Text:=commentText, Start:=1, Overwrite:=False)
                End If

                Return True
            Catch exClassic As Exception
                errorMessage = exClassic.Message
                Return False
            End Try
        End Try
    End Function

    Private Function APExcelGetValidationOptions(cell As Microsoft.Office.Interop.Excel.Range) As List(Of String)
        Dim options As New List(Of String)()

        If cell Is Nothing Then
            Return options
        End If

        Dim formula1 As String = Nothing

        Try
            If cell.Validation.Type <> Microsoft.Office.Interop.Excel.XlDVType.xlValidateList Then
                Return options
            End If

            formula1 = CStr(cell.Validation.Formula1)
        Catch
            Return options
        End Try

        If String.IsNullOrWhiteSpace(formula1) Then
            Return options
        End If

        If Not formula1.StartsWith("="c) Then
            For Each part As String In formula1.Split(New Char() {","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                APExcelAddUniqueOption(options, part)
            Next

            Return options
        End If

        Dim evaluated As Object = Nothing

        Try
            evaluated = cell.Worksheet.Evaluate(formula1)
            APExcelCollectValidationValues(evaluated, options)
        Catch
        End Try

        If options.Count = 0 Then
            Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
            Dim nm As Microsoft.Office.Interop.Excel.Name = Nothing
            Dim refersToRange As Microsoft.Office.Interop.Excel.Range = Nothing

            Try
                wb = CType(cell.Worksheet.Parent, Microsoft.Office.Interop.Excel.Workbook)
                nm = wb.Names.Item(formula1.Trim().TrimStart("="c))
                refersToRange = nm.RefersToRange
                APExcelCollectRangeValues(refersToRange, options)
            Catch
            Finally
                APExcelReleaseComObject(refersToRange)
                APExcelReleaseComObject(nm)
                APExcelReleaseComObject(wb)
            End Try
        End If

        If options.Count = 0 Then
            Dim refRange As Microsoft.Office.Interop.Excel.Range = Nothing

            Try
                refRange = cell.Worksheet.Range(formula1.Trim().TrimStart("="c))
                APExcelCollectRangeValues(refRange, options)
            Catch
            Finally
                APExcelReleaseComObject(refRange)
            End Try
        End If

        If options.Count = 0 Then
            Dim localSep As String = ","

            Try
                localSep = CStr(cell.Application.International(Microsoft.Office.Interop.Excel.XlApplicationInternational.xlListSeparator))
            Catch
            End Try

            For Each part As String In formula1.Trim().TrimStart("="c).Split(New String() {localSep, ",", ";"}, StringSplitOptions.RemoveEmptyEntries)
                APExcelAddUniqueOption(options, part)
            Next
        End If

        Return options
    End Function

    Private Sub APExcelCollectValidationValues(value As Object, options As List(Of String))
        If value Is Nothing Then
            Return
        End If

        If TypeOf value Is Microsoft.Office.Interop.Excel.Range Then
            APExcelCollectRangeValues(CType(value, Microsoft.Office.Interop.Excel.Range), options)
            Return
        End If

        If TypeOf value Is Object(,) Then
            Dim arr As Object(,) = CType(value, Object(,))

            For r As Integer = arr.GetLowerBound(0) To arr.GetUpperBound(0)
                For c As Integer = arr.GetLowerBound(1) To arr.GetUpperBound(1)
                    APExcelAddUniqueOption(options, APExcelValueToInvariantString(arr(r, c)))
                Next
            Next

            Return
        End If

        Dim asText As String = APExcelValueToInvariantString(value)
        If asText <> "" Then
            For Each part As String In asText.Split(New Char() {","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                APExcelAddUniqueOption(options, part)
            Next
        End If
    End Sub

    Private Sub APExcelCollectRangeValues(refRange As Microsoft.Office.Interop.Excel.Range, options As List(Of String))
        If refRange Is Nothing Then
            Return
        End If

        Dim raw As Object = Nothing

        Try
            raw = refRange.Value2
        Catch
            raw = Nothing
        End Try

        If raw Is Nothing Then
            Return
        End If

        If TypeOf raw Is Object(,) Then
            Dim arr As Object(,) = CType(raw, Object(,))

            For r As Integer = arr.GetLowerBound(0) To arr.GetUpperBound(0)
                For c As Integer = arr.GetLowerBound(1) To arr.GetUpperBound(1)
                    APExcelAddUniqueOption(options, APExcelValueToInvariantString(arr(r, c)))
                Next
            Next
        Else
            APExcelAddUniqueOption(options, APExcelValueToInvariantString(raw))
        End If
    End Sub

    Private Sub APExcelAddUniqueOption(options As List(Of String), text As String)
        Dim trimmed As String = If(text, "").Trim()

        If trimmed = "" Then
            Return
        End If

        If Not options.Any(Function(x) x.Equals(trimmed, StringComparison.OrdinalIgnoreCase)) Then
            options.Add(trimmed)
        End If
    End Sub

    Private Shared Function APExcelValueToInvariantString(value As Object) As String
        If value Is Nothing Then
            Return ""
        End If

        If TypeOf value Is Double Then
            Return DirectCast(value, Double).ToString(CultureInfo.InvariantCulture)
        End If

        If TypeOf value Is Single Then
            Return DirectCast(value, Single).ToString(CultureInfo.InvariantCulture)
        End If

        If TypeOf value Is Decimal Then
            Return DirectCast(value, Decimal).ToString(CultureInfo.InvariantCulture)
        End If

        If TypeOf value Is DateTime Then
            Return DirectCast(value, DateTime).ToString("o", CultureInfo.InvariantCulture)
        End If

        Return Convert.ToString(value, CultureInfo.InvariantCulture)
    End Function

    Private Function APExcelSetFormulaSafe(
            cell As Microsoft.Office.Interop.Excel.Range,
            formulaText As String,
            excelApp As Microsoft.Office.Interop.Excel.Application,
            ByRef errorMessage As String) As Boolean

        Dim localSep As String = ","
        Try
            localSep = CStr(excelApp.International(Microsoft.Office.Interop.Excel.XlApplicationInternational.xlListSeparator))
        Catch
        End Try

        Dim englishFormula As String = APExcelNormalizeFormulaText(formulaText)

        Try
            Try
                cell.Formula2 = englishFormula
            Catch ex1 As System.Runtime.InteropServices.COMException When ex1.ErrorCode = &H800A03EC
                Try
                    cell.Formula2Local = englishFormula
                Catch
                End Try
            End Try

            If Not CBool(cell.HasFormula) Then
                Try
                    cell.FormulaLocal = englishFormula
                Catch
                End Try
            End If

            If CBool(cell.HasFormula) AndAlso
               String.Equals(Trim(CStr(cell.Text)), "#NAME?", StringComparison.OrdinalIgnoreCase) Then

                Try
                    cell.FormulaLocal = englishFormula.Replace(",", localSep)
                Catch
                End Try

                If String.Equals(Trim(CStr(cell.Text)), "#NAME?", StringComparison.OrdinalIgnoreCase) Then
                    Dim converted As String = Trim(APExcelConvertFormulaToLocale(englishFormula, excelApp))

                    If converted <> "" Then
                        converted = converted.Replace(",", localSep)

                        Try
                            cell.FormulaLocal = converted
                        Catch ex4 As System.Runtime.InteropServices.COMException
                            errorMessage = $"Failed to set converted formula: {ex4.Message}"
                            Return False
                        End Try
                    End If

                    If String.Equals(Trim(CStr(cell.Text)), "#NAME?", StringComparison.OrdinalIgnoreCase) Then
                        errorMessage = $"Excel rejected the formula '{englishFormula}' for cell {cell.Address}. Resulted in #NAME?."
                        Return False
                    End If
                End If
            End If

            If Not CBool(cell.HasFormula) Then
                errorMessage = $"Excel did not accept the formula '{englishFormula}' for cell {cell.Address}."
                Return False
            End If

            Return True
        Catch comEx As System.Runtime.InteropServices.COMException
            errorMessage = $"COM Error setting formula: {comEx.Message}"
            Return False
        Catch ex As Exception
            errorMessage = $"General error setting formula: {ex.Message}"
            Return False
        End Try
    End Function

    Private Function APExcelConvertFormulaToLocale(
            englishFormula As String,
            excelApp As Microsoft.Office.Interop.Excel.Application) As String

        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim tempRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim localizedFormula As String = englishFormula

        Dim previousScreenUpdating As Boolean = excelApp.ScreenUpdating
        Dim previousDisplayAlerts As Boolean = excelApp.DisplayAlerts

        Try
            excelApp.ScreenUpdating = False
            excelApp.DisplayAlerts = False

            wb = excelApp.Workbooks.Add()
            ws = CType(wb.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)
            tempRange = ws.Range("A1")
            tempRange.Formula = englishFormula
            localizedFormula = CStr(tempRange.FormulaLocal)
        Catch
        Finally
            Try
                If wb IsNot Nothing Then
                    wb.Close(SaveChanges:=False)
                End If
            Catch
            End Try

            excelApp.DisplayAlerts = previousDisplayAlerts
            excelApp.ScreenUpdating = previousScreenUpdating

            APExcelReleaseComObject(tempRange)
            APExcelReleaseComObject(ws)
            APExcelReleaseComObject(wb)
        End Try

        Return localizedFormula
    End Function

    Private Shared Function APExcelNormalizeFormulaText(formulaText As String) As String
        Dim result As String = If(formulaText, "").Trim()

        If result.Length >= 3 Then
            If (result.StartsWith("'") AndAlso result.EndsWith("'")) OrElse
               (result.StartsWith("""") AndAlso result.EndsWith("""")) Then

                Dim inner As String = result.Substring(1, result.Length - 2).Trim()
                If inner.StartsWith("=") Then
                    result = inner
                End If
            End If
        End If

        Return result
    End Function

    Private Shared Function APExcelGetJsonString(obj As JObject, propertyName As String) As String
        If obj Is Nothing Then
            Return Nothing
        End If

        Dim token As JToken = obj(propertyName)
        If token Is Nothing OrElse token.Type = JTokenType.Null Then
            Return Nothing
        End If

        Return token.ToString()
    End Function

    Private Function APExcelTryLiftProtection(ws As Microsoft.Office.Interop.Excel.Worksheet) As APExcelLiftlockInfo
        Dim info As New APExcelLiftlockInfo With {
            .Found = False,
            .Password = "",
            .WasUnprotected = False
        }

        If ws Is Nothing Then Return info
        If Not ws.ProtectContents Then Return info

        Try
            Dim prot As Microsoft.Office.Interop.Excel.Protection = ws.Protection
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
        End Try

        Dim used As Microsoft.Office.Interop.Excel.Range = Nothing

        Try
            used = ws.UsedRange
        Catch
            Return info
        End Try

        If used Is Nothing Then Return info

        Dim prefixes() As String = {AN2 & "_liftlock", AN5 & "_liftlock"}

        For Each prefix As String In prefixes
            Dim found As Microsoft.Office.Interop.Excel.Range = Nothing

            Try
                found = used.Find(
                    What:=prefix,
                    LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                    LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
                    SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                    SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext,
                    MatchCase:=False)
            Catch
                Continue For
            End Try

            If found IsNot Nothing Then
                info.Found = True

                Dim cellText As String = ""
                Try
                    cellText = CStr(found.Value).Trim()
                Catch
                End Try

                Dim idx As Integer = cellText.IndexOf(prefix, StringComparison.OrdinalIgnoreCase)
                If idx >= 0 Then
                    Dim remainder As String = cellText.Substring(idx + prefix.Length).Trim()
                    If remainder.StartsWith("=") Then
                        info.Password = remainder.Substring(1).Trim()
                    End If
                End If

                Try
                    If String.IsNullOrEmpty(info.Password) Then
                        ws.Unprotect()
                    Else
                        ws.Unprotect(info.Password)
                    End If

                    info.WasUnprotected = True
                Catch
                    info.WasUnprotected = False
                End Try

                APExcelReleaseComObject(found)
                Exit For
            End If

            APExcelReleaseComObject(found)
        Next

        APExcelReleaseComObject(used)
        Return info
    End Function

    Private Sub APExcelReprotectWorksheet(
            ws As Microsoft.Office.Interop.Excel.Worksheet,
            info As APExcelLiftlockInfo)

        If ws Is Nothing OrElse Not info.WasUnprotected Then
            Return
        End If

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
        End Try
    End Sub

    Private Function APExcelHasLiftlockToken(ws As Microsoft.Office.Interop.Excel.Worksheet) As Boolean
        If ws Is Nothing Then
            Return False
        End If

        Dim used As Microsoft.Office.Interop.Excel.Range = Nothing

        Try
            used = ws.UsedRange
            If used Is Nothing Then Return False

            Dim prefixes() As String = {AN2 & "_liftlock", AN5 & "_liftlock"}

            For Each prefix As String In prefixes
                Dim found As Microsoft.Office.Interop.Excel.Range = Nothing

                Try
                    found = used.Find(
                        What:=prefix,
                        LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                        LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
                        SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                        SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext,
                        MatchCase:=False)

                    If found IsNot Nothing Then
                        APExcelReleaseComObject(found)
                        Return True
                    End If
                Catch
                Finally
                    APExcelReleaseComObject(found)
                End Try
            Next
        Catch
        Finally
            APExcelReleaseComObject(used)
        End Try

        Return False
    End Function

    Private Shared Function APExcelGetRangeAddress(rng As Microsoft.Office.Interop.Excel.Range) As String
        If rng Is Nothing Then
            Return ""
        End If

        Try
            Return rng.Address(False, False)
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function APExcelGetUsedRowCount(rng As Microsoft.Office.Interop.Excel.Range) As Integer
        If rng Is Nothing Then
            Return 0
        End If

        Try
            Return rng.Rows.Count
        Catch
            Return 0
        End Try
    End Function

    Private Shared Function APExcelGetUsedColumnCount(rng As Microsoft.Office.Interop.Excel.Range) As Integer
        If rng Is Nothing Then
            Return 0
        End If

        Try
            Return rng.Columns.Count
        Catch
            Return 0
        End Try
    End Function

    Private Shared Function APExcelOleColorToHex(oleColor As Integer) As String
        Dim color As System.Drawing.Color = System.Drawing.ColorTranslator.FromOle(oleColor)
        Return $"#{color.R:X2}{color.G:X2}{color.B:X2}"
    End Function

    Private Shared Sub APExcelReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing AndAlso Marshal.IsComObject(obj) Then
                Marshal.FinalReleaseComObject(obj)
            End If
        Catch
        End Try
    End Sub

End Class
