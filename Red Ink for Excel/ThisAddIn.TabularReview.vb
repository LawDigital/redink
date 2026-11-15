' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. For license to use see https://redink.ai.
' Office adapter for the shared review engine/UI. Capture one workbook at launch;
' never redirect a later asynchronous result to whichever workbook is active.
Option Explicit On
Option Strict On

Partial Public Class ThisAddIn
    Private _tabularReview As Global.SharedLibrary.TabularReview.TabularReviewForm
    Private _tabularReviewStore As TabularReviewWorkbookStore
    Private _tabularCloseInProgress As System.Boolean

    Public Sub OpenTabularReview()
        Try
            Global.SharedLibrary.SharedLibrary.SharedMethods.RequireInteractiveExecution("tabular_review")
            If _tabularReview IsNot Nothing AndAlso Not _tabularReview.IsDisposed Then
                _tabularReview.WindowState = System.Windows.Forms.FormWindowState.Normal
                _tabularReview.BringToFront()
                _tabularReview.Activate()
                Return
            End If
            Dim workbook As Microsoft.Office.Interop.Excel.Workbook = Me.Application.ActiveWorkbook
            If workbook Is Nothing Then workbook = Me.Application.Workbooks.Add()
            If workbook.ReadOnly Then Throw New System.InvalidOperationException("Save a writable copy of this workbook, then open Tabular Review again.")
            If workbook.ProtectStructure Then Throw New System.InvalidOperationException("Unprotect the workbook structure before opening Tabular Review.")
            Dim fileFormat As System.Int32 = CInt(workbook.FileFormat)
            If fileFormat <> 51 AndAlso fileFormat <> 52 AndAlso fileFormat <> 50 Then Throw New System.InvalidOperationException("Save this workbook as .xlsx, .xlsm or .xlsb, then open Tabular Review again so Excel can retain the complete review.")
            Dim store As New TabularReviewWorkbookStore(workbook)
            Dim project As Global.SharedLibrary.TabularReview.ReviewProject = Nothing
            Dim recoveredFirstWrite As System.Boolean = False
            Try
                project = store.Load()
            Catch ex As TabularReviewWorkbookStore.InvalidInitialSnapshotException
                System.Diagnostics.Trace.WriteLine("[TabularReview:initial_load_failed] " & ex.ToString())
                ' An old interrupted first write must never block all further folder
                ' imports. Explicitly preserve the old VeryHidden sheet and visible
                ' results; do not silently overwrite either.
                Dim answer As System.Int32 = Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomYesNoBox(
                    "An earlier review could not be reopened. Its existing data will be kept in this workbook." & System.Environment.NewLine &
                    "Start a new review?", "Start new review", "Cancel", "Red Ink — Review recovery")
                If answer <> 1 Then Return
                store.ArchiveInvalidInitialSnapshot()
                recoveredFirstWrite = True
            End Try
            If project Is Nothing Then project = New Global.SharedLibrary.TabularReview.ReviewProject()
            Dim callbacks As New Global.SharedLibrary.TabularReview.ReviewHostCallbacks With {
                .EnsureAvailable = Sub() store.EnsureAvailable(),
                .SaveProject = Sub(value As Global.SharedLibrary.TabularReview.ReviewProject) store.Save(value),
                .ReloadProject = Function() store.RefreshFromWorksheet(),
                .ImportSelection = Function() ImportTabularReviewSelection(store),
                .SaveWorkbook = Function() SaveTabularReviewWorkbook(store),
                .SaveWorkbookAs = Function() SaveTabularReviewWorkbook(store, saveAs:=True),
                .AutoSaveWorkbook = Function() SaveTabularReviewWorkbook(store, automatic:=True),
                .WorkbookSaveState = Function() TabularReviewWorkbookSaveState(store),
                .WorkbookLabel = Function() TabularReviewWorkbookLabel(store),
                .WorkbookPath = Function() If(System.String.IsNullOrEmpty(workbook.Path), "", workbook.FullName)
            }
            Dim form As New Global.SharedLibrary.TabularReview.TabularReviewForm(_context, project, callbacks, Me.HostWindowLifetimeToken)
            _tabularReviewStore = store
            _tabularReview = form
            AddHandler Me.Application.WorkbookBeforeClose, AddressOf TabularReviewWorkbookBeforeClose
            AddHandler form.FormClosed,
                Sub(sender As System.Object, args As System.Windows.Forms.FormClosedEventArgs)
                    RemoveHandler Me.Application.WorkbookBeforeClose, AddressOf TabularReviewWorkbookBeforeClose
                    If System.Object.ReferenceEquals(_tabularReview, form) Then
                        _tabularReview = Nothing
                        _tabularReviewStore = Nothing
                    End If
                End Sub
            Me.ShowHostModelessForm(form)
            If Not recoveredFirstWrite AndAlso Not System.String.IsNullOrWhiteSpace(store.LastNotice) Then
                Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(form)
                    Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox(store.LastNotice, "Red Ink — Tabular Review")
                End Using
            End If
        Catch ex As System.Exception
            Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Tabular Review could not start: " & ex.Message, "Red Ink — Tabular Review")
        End Try
    End Sub

    Private Function ImportTabularReviewSelection(store As TabularReviewWorkbookStore) As Global.SharedLibrary.TabularReview.ReviewProject
        store.EnsureAvailable()
        Dim form As Global.SharedLibrary.TabularReview.TabularReviewForm = _tabularReview
        Dim selected As Microsoft.Office.Interop.Excel.Range = Nothing
        ' Excel's native range selector remains on its owning STA. Hide only our
        ' own modeless form while the user selects the source header and data.
        If form IsNot Nothing Then form.Hide()
        Try
            Dim picked As System.Object = Me.Application.InputBox(
                Prompt:="Select the existing result including its header row and all data rows. Exclude totals and diagnostic blocks. The original cells will stay unchanged.",
                Title:="Import result into Tabular Review", Type:=8)
            selected = TryCast(picked, Microsoft.Office.Interop.Excel.Range)
        Finally
            If form IsNot Nothing AndAlso Not form.IsDisposed Then
                form.Show()
                form.Activate()
            End If
        End Try
        If selected Is Nothing Then Return Nothing
        If selected.Areas.Count <> 1 Then Throw New System.InvalidOperationException("Choose one rectangular range containing the header row and data.")
        If selected.Rows.Count < 2 Then Throw New System.InvalidOperationException("Include the header row and at least one result row.")
        Dim column As System.Int32 = 0
        Dim sourceRoot As System.String = If(form IsNot Nothing, form.Project.SourceRoot, "")
        Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(form)
            Dim answer As System.String = Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomInputBox(
                "Which selected column identifies the source file? Enter its position within the selected range (1, 2, ...), or 0 for automatic detection of the existing File/path header. All original columns will be retained.",
                "Red Ink — Tabular Review import", True, "0")
            If System.String.IsNullOrWhiteSpace(answer) Then Return Nothing
            If Not System.Int32.TryParse(answer.Trim(), column) OrElse column < 0 OrElse column > selected.Columns.Count Then Throw New System.InvalidOperationException("Enter a valid source-column position or 0.")
        End Using
        Return store.ImportSelection(selected, sourceRoot, column)
    End Function

    Private Function SaveTabularReviewWorkbook(store As TabularReviewWorkbookStore,
                                              Optional saveAs As System.Boolean = False,
                                              Optional automatic As System.Boolean = False) As Global.SharedLibrary.TabularReview.ReviewWorkbookSaveResult
        store.EnsureAvailable()
        Dim workbook As Microsoft.Office.Interop.Excel.Workbook = store.Workbook
        Dim state As Global.SharedLibrary.TabularReview.ReviewWorkbookSaveState = TabularReviewWorkbookSaveState(store)
        If automatic AndAlso Not state.CanAutoSave Then
            Return New Global.SharedLibrary.TabularReview.ReviewWorkbookSaveResult With {.Status = "skipped", .Message = state.Message}
        End If
        Dim form As Global.SharedLibrary.TabularReview.TabularReviewForm = Nothing
        If System.Object.ReferenceEquals(_tabularReviewStore, store) Then form = _tabularReview
        If Not saveAs AndAlso workbook.ReadOnly Then
            Return New Global.SharedLibrary.TabularReview.ReviewWorkbookSaveResult With {
                .Status = "failed", .Message = "This workbook is read-only. Use Save As to keep a writable copy."}
        End If
        If System.String.IsNullOrWhiteSpace(workbook.Path) Then saveAs = True
        If Not saveAs Then
            ' Prepare portable-folder metadata before the disk write, including a
            ' native Excel Save As that changed the captured workbook's location.
            PrepareTabularReviewSavePath(store, workbook.FullName)
            state = TabularReviewWorkbookSaveState(store)
            If automatic AndAlso Not state.CanAutoSave Then
                Return New Global.SharedLibrary.TabularReview.ReviewWorkbookSaveResult With {.Status = "skipped", .Message = state.Message}
            End If
            If state.Saved Then
                Return New Global.SharedLibrary.TabularReview.ReviewWorkbookSaveResult With {.Status = If(automatic, "skipped", "saved"), .Message = "Workbook is already saved."}
            End If
            Return WriteTabularReviewWorkbookFile(store, "", Nothing, automatic)
        End If
        Dim currentFormat As System.Int32 = CInt(workbook.FileFormat)
        Dim extension As System.String = If(currentFormat = 50, "xlsb", If(currentFormat = 52 OrElse workbook.HasVBProject, "xlsm", "xlsx"))
        Dim defaultName As System.String = System.IO.Path.ChangeExtension(workbook.Name, extension)
        If System.String.IsNullOrWhiteSpace(workbook.Path) AndAlso form IsNot Nothing AndAlso Not form.IsDisposed Then
            defaultName = TabularReviewFileBaseName(form.Project.Title, System.IO.Path.GetFileNameWithoutExtension(workbook.Name)) & "." & extension
        End If
        Using dialog As New System.Windows.Forms.SaveFileDialog With {
            .Title = "Save Tabular Review workbook as", .Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm|Excel Binary Workbook (*.xlsb)|*.xlsb",
            .AddExtension = True, .OverwritePrompt = True, .FileName = defaultName, .DefaultExt = extension,
            .FilterIndex = If(extension = "xlsb", 3, If(extension = "xlsm", 2, 1)),
            .CheckPathExists = True, .RestoreDirectory = True
        }
            If form IsNot Nothing AndAlso Not form.IsDisposed AndAlso System.IO.Directory.Exists(form.Project.SourceRoot) Then
                dialog.InitialDirectory = form.Project.SourceRoot
            ElseIf Not System.String.IsNullOrWhiteSpace(workbook.Path) AndAlso System.IO.Directory.Exists(workbook.Path) Then
                dialog.InitialDirectory = workbook.Path
            End If
            Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(form)
                Dim owner As System.Windows.Forms.IWin32Window = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If dialog.ShowDialog(owner) <> System.Windows.Forms.DialogResult.OK Then
                    Return New Global.SharedLibrary.TabularReview.ReviewWorkbookSaveResult With {
                        .Status = "cancelled", .Message = "Save As cancelled. Your review changes remain in the open Excel workbook."}
                End If
            End Using
            store.EnsureAvailable()
            If Me.HostWindowLifetimeToken.IsCancellationRequested Then Throw New System.OperationCanceledException(Me.HostWindowLifetimeToken)
            Dim format As Microsoft.Office.Interop.Excel.XlFileFormat
            Select Case System.IO.Path.GetExtension(dialog.FileName).ToLowerInvariant()
                Case ".xlsx"
                    If workbook.HasVBProject Then Throw New System.InvalidOperationException("Choose .xlsm or .xlsb to retain this workbook's VBA project.")
                    format = Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook
                Case ".xlsm"
                    format = Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled
                Case ".xlsb"
                    format = Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12
                Case Else
                    Throw New System.InvalidOperationException("Use .xlsx, .xlsm or .xlsb so that the complete review stays in the workbook.")
            End Select
            Dim targetPath As System.String = System.IO.Path.GetFullPath(dialog.FileName)
            Dim preparationCompleted As System.Boolean = False
            Dim result As Global.SharedLibrary.TabularReview.ReviewWorkbookSaveResult
            Try
                PrepareTabularReviewSavePath(store, targetPath)
                preparationCompleted = True
                result = WriteTabularReviewWorkbookFile(store, targetPath, format, False)
            Catch saveFailure As System.Exception
                Dim repairFailure As System.Exception = ReconcileTabularReviewSavePath(store, Not preparationCompleted)
                If repairFailure IsNot Nothing Then
                    Throw New System.IO.IOException(TabularReviewSavePathFailureMessage(saveFailure.Message, repairFailure),
                        New System.AggregateException(saveFailure, repairFailure))
                End If
                Throw
            End Try
            If result IsNot Nothing AndAlso result.Status = "saved" Then Return result
            Dim reconciliationFailure As System.Exception = ReconcileTabularReviewSavePath(store, False)
            If reconciliationFailure IsNot Nothing Then
                Return New Global.SharedLibrary.TabularReview.ReviewWorkbookSaveResult With {
                    .Status = "failed", .Message = TabularReviewSavePathFailureMessage(If(result Is Nothing, "", result.Message), reconciliationFailure)}
            End If
            Return If(result, New Global.SharedLibrary.TabularReview.ReviewWorkbookSaveResult With {
                .Status = "failed", .Message = "Excel did not confirm Save As. Your review remains open."})
        End Using
    End Function

    Private Sub PrepareTabularReviewSavePath(store As TabularReviewWorkbookStore, targetPath As System.String)
        Dim form As Global.SharedLibrary.TabularReview.TabularReviewForm = _tabularReview
        If System.Object.ReferenceEquals(_tabularReviewStore, store) AndAlso form IsNot Nothing AndAlso Not form.IsDisposed Then
            form.PrepareWorkbookSavePath(targetPath)
        End If
    End Sub

    Private Function ReconcileTabularReviewSavePath(store As TabularReviewWorkbookStore,
                                                   forceCheckpoint As System.Boolean) As System.Exception
        Try
            If Me.HostWindowLifetimeToken.IsCancellationRequested Then Throw New System.OperationCanceledException(Me.HostWindowLifetimeToken)
            store.EnsureAvailable()
            Dim form As Global.SharedLibrary.TabularReview.TabularReviewForm = _tabularReview
            If Not System.Object.ReferenceEquals(_tabularReviewStore, store) OrElse form Is Nothing OrElse form.IsDisposed Then
                Throw New System.InvalidOperationException("The workbook's review window is no longer available.")
            End If
            Dim workbook As Microsoft.Office.Interop.Excel.Workbook = store.Workbook
            Dim actualPath As System.String = If(System.String.IsNullOrWhiteSpace(workbook.Path), "", workbook.FullName)
            ' Excel can keep either location after rejecting Save As. Reconcile
            ' only the portable metadata with that actual location; never retry
            ' the native save. A failed preparation needs a guarded checkpoint
            ' even if the form still holds its earlier in-memory metadata.
            form.PrepareWorkbookSavePath(actualPath, forceCheckpoint:=forceCheckpoint)
            System.Diagnostics.Trace.WriteLine("[TabularReview:save_path_reconciled] Used the bound workbook's actual location.")
            Return Nothing
        Catch repairFailure As System.Exception
            System.Diagnostics.Trace.WriteLine("[TabularReview:save_path_reconcile_failed] " & repairFailure.ToString())
            Return repairFailure
        End Try
    End Function

    Private Shared Function TabularReviewSavePathFailureMessage(primaryMessage As System.String, repairFailure As System.Exception) As System.String
        Return If(System.String.IsNullOrWhiteSpace(primaryMessage), "Excel did not confirm Save As.", primaryMessage) &
            System.Environment.NewLine & "Red Ink could not reconcile the source-folder reference: " & repairFailure.Message &
            System.Environment.NewLine & "Keep this workbook open. In Tabular Review, choose Reload from Excel, then Save again before closing."
    End Function

    Private Shared Function TabularReviewFileBaseName(title As System.String, fallback As System.String) As System.String
        Dim value As System.String = If(title, "").Trim()
        Dim invalid As System.Char() = System.IO.Path.GetInvalidFileNameChars()
        Dim result As New System.Text.StringBuilder()
        For Each character As System.Char In value
            result.Append(If(System.Array.IndexOf(invalid, character) >= 0 OrElse System.Char.IsControl(character), "_"c, character))
        Next
        value = result.ToString().Trim().TrimEnd("."c, " "c)
        If value.Length = 0 Then value = If(fallback, "Tabular Review")
        If value.Length > 120 Then
            Dim length As System.Int32 = If(System.Char.IsHighSurrogate(value(119)), 119, 120)
            value = value.Substring(0, length).TrimEnd("."c, " "c)
        End If
        Dim deviceName As System.String = value.Split("."c)(0).ToUpperInvariant()
        If deviceName = "CON" OrElse deviceName = "PRN" OrElse deviceName = "AUX" OrElse deviceName = "NUL" OrElse
            System.Text.RegularExpressions.Regex.IsMatch(deviceName, "^(COM|LPT)[1-9]$") Then value = "Review - " & value
        Return value
    End Function

    Private Function WriteTabularReviewWorkbookFile(store As TabularReviewWorkbookStore,
                                                    targetPath As System.String,
                                                    format As System.Nullable(Of Microsoft.Office.Interop.Excel.XlFileFormat),
                                                    automatic As System.Boolean) As Global.SharedLibrary.TabularReview.ReviewWorkbookSaveResult
        store.EnsureAvailable()
        If Me.HostWindowLifetimeToken.IsCancellationRequested Then Throw New System.OperationCanceledException(Me.HostWindowLifetimeToken)
        Dim workbook As Microsoft.Office.Interop.Excel.Workbook = store.Workbook
        Dim application As Microsoft.Office.Interop.Excel.Application = workbook.Application
        Dim completed As System.Nullable(Of System.Boolean) = Nothing
        Dim afterSave As Microsoft.Office.Interop.Excel.AppEvents_WorkbookAfterSaveEventHandler =
            Sub(savedWorkbook As Microsoft.Office.Interop.Excel.Workbook, success As System.Boolean)
                If TabularReviewSameWorkbook(savedWorkbook, workbook) Then completed = success
            End Sub
        Dim previousAlerts As System.Boolean = application.DisplayAlerts
        AddHandler application.WorkbookAfterSave, afterSave
        Try
            ' Autosave never opens our picker or native Excel alerts. Retain Excel
            ' events so other add-ins and workbook validation can reject the save.
            If automatic Then application.DisplayAlerts = False
            If format.HasValue Then
                workbook.SaveAs(Filename:=targetPath, FileFormat:=format.Value)
            Else
                workbook.Save()
            End If
            store.EnsureAvailable()
            Dim hasPath As System.Boolean = Not System.String.IsNullOrWhiteSpace(workbook.Path)
            Dim correctPath As System.Boolean = Not format.HasValue OrElse System.String.Equals(workbook.FullName, targetPath, System.StringComparison.OrdinalIgnoreCase)
            Dim confirmed As System.Boolean = If(completed.HasValue, completed.Value, workbook.Saved)
            System.Diagnostics.Trace.WriteLine("[TabularReview:workbook_save] mode=" & If(automatic, "auto", If(format.HasValue, "save_as", "save")) &
                " | afterSave=" & If(completed.HasValue, completed.Value.ToString(), "not_observed") &
                " | savedFlag=" & workbook.Saved.ToString() & " | hasPath=" & hasPath.ToString() & " | targetMatches=" & correctPath.ToString())
            If Not confirmed OrElse Not hasPath OrElse Not correctPath Then
                Return New Global.SharedLibrary.TabularReview.ReviewWorkbookSaveResult With {
                    .Status = "failed", .Message = "Excel did not confirm the workbook save. Your review remains open. Check Excel for a save error or choose Save As."}
            End If
            If Not workbook.Saved Then
                Return New Global.SharedLibrary.TabularReview.ReviewWorkbookSaveResult With {
                    .Status = "failed", .Message = "Excel wrote the workbook file, then another Excel change made it unsaved again. Save again to keep those newer changes."}
            End If
            Return New Global.SharedLibrary.TabularReview.ReviewWorkbookSaveResult With {
                .Status = "saved", .Message = If(automatic, "Workbook autosaved.", "Workbook saved.")}
        Catch ex As System.Exception
            System.Diagnostics.Trace.WriteLine("[TabularReview:workbook_save_failed] " & ex.ToString())
            Throw
        Finally
            Try
                RemoveHandler application.WorkbookAfterSave, afterSave
            Finally
                If automatic Then application.DisplayAlerts = previousAlerts
            End Try
        End Try
    End Function

    Private Shared Function TabularReviewWorkbookSaveState(store As TabularReviewWorkbookStore) As Global.SharedLibrary.TabularReview.ReviewWorkbookSaveState
        store.EnsureAvailable()
        Dim workbook As Microsoft.Office.Interop.Excel.Workbook = store.Workbook
        Dim hasPath As System.Boolean = Not System.String.IsNullOrWhiteSpace(workbook.Path)
        Dim state As New Global.SharedLibrary.TabularReview.ReviewWorkbookSaveState With {.Saved = hasPath AndAlso workbook.Saved}
        If Not hasPath Then
            state.CanAutoSave = Not workbook.ReadOnly AndAlso workbook.Application.Ready
            state.Message = "Choose a filename on the first save."
        ElseIf workbook.ReadOnly Then
            state.Message = "This workbook is read-only. Use Save As to keep a writable copy."
        Else
            Dim format As System.Int32 = CInt(workbook.FileFormat)
            If format <> 51 AndAlso format <> 52 AndAlso format <> 50 Then
                state.Message = "Use Save As with .xlsx, .xlsm or .xlsb to retain the complete review."
            ElseIf Not workbook.Application.Ready Then
                state.Message = "AutoSave will resume when Excel finishes its current operation."
            Else
                state.CanAutoSave = True
                state.Message = If(state.Saved, "Workbook saved.", "Changes are in Excel; Save or AutoSave will write this workbook's file.")
            End If
        End If
        Return state
    End Function

    Private Shared Function TabularReviewWorkbookLabel(store As TabularReviewWorkbookStore) As System.String
        store.EnsureAvailable()
        Dim workbook As Microsoft.Office.Interop.Excel.Workbook = store.Workbook
        Return workbook.Name & If(System.String.IsNullOrWhiteSpace(workbook.Path), " · open in Excel; use Save to choose a file", If(workbook.Saved, " · saved", " · unsaved changes"))
    End Function

    Private Sub TabularReviewWorkbookBeforeClose(workbook As Microsoft.Office.Interop.Excel.Workbook, ByRef cancel As System.Boolean)
        If _tabularReview Is Nothing OrElse _tabularReview.IsDisposed OrElse _tabularReviewStore Is Nothing OrElse Me.HostWindowLifetimeToken.IsCancellationRequested Then Return
        If Not TabularReviewSameWorkbook(workbook, _tabularReviewStore.Workbook) Then Return
        cancel = True
        If _tabularCloseInProgress Then Return
        _tabularCloseInProgress = True
        Dim form As Global.SharedLibrary.TabularReview.TabularReviewForm = _tabularReview
        ' Do not capture Excel's ByRef cancel parameter in an async continuation.
        form.BeginInvoke(New System.Action(Sub() CloseTabularReviewWorkbookAsync(form, workbook)))
    End Sub

    Private Async Sub CloseTabularReviewWorkbookAsync(form As Global.SharedLibrary.TabularReview.TabularReviewForm,
                                                      workbook As Microsoft.Office.Interop.Excel.Workbook)
        Try
            If Await form.FlushAndCloseAsync() Then
                Await EnsureUIThread()
                If Not Me.HostWindowLifetimeToken.IsCancellationRequested Then workbook.Close()
            End If
        Catch ex As System.Exception
            If Not Me.HostWindowLifetimeToken.IsCancellationRequested Then
                Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Excel has kept the workbook open: " & ex.Message, "Red Ink — Tabular Review")
            End If
        Finally
            _tabularCloseInProgress = False
        End Try
    End Sub

    Private Shared Function TabularReviewSameWorkbook(first As Microsoft.Office.Interop.Excel.Workbook,
                                                      second As Microsoft.Office.Interop.Excel.Workbook) As System.Boolean
        If System.Object.ReferenceEquals(first, second) Then Return True
        Dim firstPointer As System.IntPtr = System.IntPtr.Zero
        Dim secondPointer As System.IntPtr = System.IntPtr.Zero
        Try
            firstPointer = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(first)
            secondPointer = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(second)
            Return firstPointer = secondPointer
        Finally
            If firstPointer <> System.IntPtr.Zero Then System.Runtime.InteropServices.Marshal.Release(firstPointer)
            If secondPointer <> System.IntPtr.Zero Then System.Runtime.InteropServices.Marshal.Release(secondPointer)
        End Try
    End Function
End Class
