' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. For license to use see https://redink.ai.
' Excel owns both the visible projection and the complete, versioned review state.
' All Office calls stay on the thread which created this adapter. No method saves
' the workbook to disk; the owning UI supplies that explicit operation.
Option Explicit On
Option Strict On

Friend NotInheritable Class TabularReviewWorkbookStore
    Private Const StoreMarker As System.String = "RedInk.TabularReview.Store.v1"
    Private Const ProjectionMarker As System.String = "RedInk.TabularReview.Project:"
    Private Const PayloadFirstRow As System.Int32 = 10
    Private Const PayloadChunkLength As System.Int32 = 30000
    Private Const MaximumCellLength As System.Int32 = 32767
    Private Const MaximumRows As System.Int32 = 1048576
    Private Const MaximumColumns As System.Int32 = 16384
    Private Const MaximumSnapshotBytes As System.Int64 = 536870912
    Private Const BlockRows As System.Int32 = 64
    Private Const ManagedWorksheetNotice As System.String = "Red Ink manages this worksheet automatically. Open Analyze → Tabular Review to work with this review."
    Private Const OverflowNotice As System.String = "[Full text retained in Tabular Review]"

    Private ReadOnly _workbook As Microsoft.Office.Interop.Excel.Workbook
    Private ReadOnly _uiThreadId As System.Int32
    Private _storeSheetName As System.String = ""
    Private _visibleSheetName As System.String = ""
    Private _snapshot As StoredSnapshot
    Private _loaded As Global.SharedLibrary.TabularReview.ReviewProject
    Private _activeSlot As System.Int32
    Private _committedSlotObserved As System.Int32
    Private _expectedFingerprint As System.String = ""
    Private _expectedRows As System.Int32
    Private _expectedColumns As System.Int32
    Private _lastNotice As System.String = ""

    ' A typed recovery condition; callers must never infer data-loss safety from
    ' the text of an exception or from a general InvalidDataException.
    Friend NotInheritable Class InvalidInitialSnapshotException
        Inherits System.IO.IOException
        Public Sub New(message As System.String, cause As System.Exception)
            MyBase.New(message, cause)
        End Sub
    End Class

    Private NotInheritable Class StoredSnapshot
        Public Property FormatVersion As System.Int32 = 1
        Public Property Project As Global.SharedLibrary.TabularReview.ReviewProject
        Public Property WorksheetName As System.String = ""
        Public Property WorksheetCodeName As System.String = ""
        Public Property ProjectionRows As System.Int32
        Public Property ProjectionColumns As System.Int32
        Public Property ProjectionFingerprint As System.String = ""
    End Class

    Private NotInheritable Class ProjectionColumn
        Public Property Category As System.String = ""
        Public Property Id As System.String = ""
        Public Property Title As System.String = ""
        Public Property Instruction As System.String = ""
        Public Property CriterionId As System.String = ""
        Public Property Field As System.String = ""
    End Class

    Private NotInheritable Class ProjectionBackupBlock
        Public Property FirstRow As System.Int32
        Public Property Values As System.Object(,)
    End Class

    Public Sub New(workbook As Microsoft.Office.Interop.Excel.Workbook)
        If workbook Is Nothing Then Throw New System.ArgumentNullException(NameOf(workbook))
        _workbook = workbook
        _uiThreadId = System.Threading.Thread.CurrentThread.ManagedThreadId
        EnsureAvailable()
    End Sub

    Public ReadOnly Property Workbook As Microsoft.Office.Interop.Excel.Workbook
        Get
            EnsureAvailable()
            Return _workbook
        End Get
    End Property

    Public ReadOnly Property VisibleSheetName As System.String
        Get
            Return _visibleSheetName
        End Get
    End Property

    Public ReadOnly Property LastNotice As System.String
        Get
            Return _lastNotice
        End Get
    End Property

    Public Sub EnsureAvailable()
        If System.Threading.Thread.CurrentThread.ManagedThreadId <> _uiThreadId Then
            Throw New System.InvalidOperationException("Tabular Review must access Excel on its owning Office UI thread.")
        End If
        Try
            Dim application As Microsoft.Office.Interop.Excel.Application = _workbook.Application
            Dim books As Microsoft.Office.Interop.Excel.Workbooks = application.Workbooks
            Try
                For index As System.Int32 = 1 To books.Count
                    ' Workbook/Application are borrowed host objects. Do not release their RCWs.
                    Dim candidate As Microsoft.Office.Interop.Excel.Workbook = books.Item(index)
                    If SameComIdentity(candidate, _workbook) Then Return
                Next
            Finally
                ReleaseTemporary(books)
            End Try
        Catch ex As System.Exception
            Throw New System.InvalidOperationException("This review's workbook is no longer available. Reopen it and choose Tabular Review again.", ex)
        End Try
        Throw New System.InvalidOperationException("This review's workbook has closed. Reopen it and choose Tabular Review again.")
    End Sub

    Public Function Load() As Global.SharedLibrary.TabularReview.ReviewProject
        EnsureAvailable()
        _lastNotice = ""
        Dim store As Microsoft.Office.Interop.Excel.Worksheet = FindStoreSheet()
        If store Is Nothing Then
            _snapshot = Nothing
            _loaded = Nothing
            _activeSlot = 0
            _committedSlotObserved = 0
            _expectedFingerprint = ""
            Return Nothing
        End If
        Try
            _storeSheetName = store.Name
            Dim committedSlot As System.Int32
            If Not System.Int32.TryParse(CellText(store, 2, 2), committedSlot) OrElse committedSlot < 0 OrElse committedSlot > 2 Then
                Throw New System.IO.InvalidDataException("The review commit marker is invalid. Keep a copy of this workbook before recovering its review data.")
            End If
            _committedSlotObserved = committedSlot
            Dim initialRecovery As System.Boolean = False
            If committedSlot = 0 Then
                ' Recover only one fully written first snapshot. Its checksum and
                ' projection still pass the ordinary validator below; do not guess
                ' between two uncommitted candidates or overwrite either candidate.
                Dim firstComplete As System.Boolean = CellText(store, 5, 4) = "complete"
                Dim secondComplete As System.Boolean = CellText(store, 5, 5) = "complete"
                If firstComplete AndAlso secondComplete Then
                    Throw New System.IO.InvalidDataException("This workbook contains two complete review snapshots but no commit marker. Keep this workbook and its visible results; restore a known-good copy before continuing.")
                End If
                If Not firstComplete AndAlso Not secondComplete Then
                    Dim leftState As System.String = CellText(store, 5, 4)
                    Dim rightState As System.String = CellText(store, 5, 5)
                    If leftState = "writing" OrElse rightState = "writing" Then
                        ' An interrupted write may have finished all its bytes and
                        ' projection. Verify BOTH before promoting an uncommitted slot.
                        ' Never guess if two candidate slots are present.
                        Dim writingSlot As System.Int32 = If(leftState = "writing" AndAlso rightState <> "writing", 1,
                            If(rightState = "writing" AndAlso leftState <> "writing", 2, 0))
                        If writingSlot = 0 Then Throw New InvalidInitialSnapshotException("An earlier review could not be reopened.", Nothing)
                        Dim pending As StoredSnapshot = Nothing
                        Dim pendingProjection As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                        Try
                            pending = ReadSnapshot(store, writingSlot, True)
                            pendingProjection = FindProjectionSheet(pending)
                            If pendingProjection Is Nothing OrElse Not ProjectionMatches(pendingProjection, pending.ProjectionRows, pending.ProjectionColumns, pending.ProjectionFingerprint) Then
                                Throw New System.IO.InvalidDataException("The interrupted review worksheet no longer matches its snapshot.")
                            End If
                            PutText(store, 5, writingSlot + 3, "complete")
                            PutText(store, 2, 2, writingSlot.ToString(System.Globalization.CultureInfo.InvariantCulture))
                            committedSlot = writingSlot
                            _committedSlotObserved = writingSlot
                            _lastNotice = "An interrupted review was recovered. Save the workbook to keep it."
                            System.Diagnostics.Trace.WriteLine("[TabularReview:initial_snapshot_recovered] slot=" & writingSlot.ToString())
                        Catch ex As System.Exception
                            System.Diagnostics.Trace.WriteLine("[TabularReview:initial_snapshot_unrecoverable] " & ex.ToString())
                            Throw New InvalidInitialSnapshotException("An earlier review could not be reopened.", ex)
                        Finally
                            ReleaseTemporary(pendingProjection)
                        End Try
                    Else
                        Return Nothing
                    End If
                End If
                If committedSlot = 0 Then
                    committedSlot = If(firstComplete, 1, 2)
                    initialRecovery = True
                End If
            End If
            Dim restored As StoredSnapshot = Nothing
            Dim loadFailure As System.Exception = Nothing
            Try
                restored = ReadSnapshot(store, committedSlot)
            Catch ex As System.Exception
                loadFailure = ex
            End Try
            If initialRecovery AndAlso restored Is Nothing Then
                System.Diagnostics.Trace.WriteLine("[TabularReview:first_snapshot_validation_failed] " & loadFailure.ToString())
                Throw New InvalidInitialSnapshotException("The interrupted first review snapshot failed validation (" & loadFailure.Message & "). Keep this workbook and its visible results; Tabular Review will not replace the stored data.", loadFailure)
            End If
            If restored Is Nothing Then
                Dim alternate As System.Int32 = 3 - committedSlot
                Try
                    restored = ReadSnapshot(store, alternate)
                    committedSlot = alternate
                Catch ex As System.Exception
                    Throw New System.IO.InvalidDataException("Neither saved review snapshot passed validation. Keep this workbook; its visible results remain available. Restore a known-good workbook copy before continuing.", New System.AggregateException(loadFailure, ex))
                End Try
                _lastNotice = "The latest review snapshot failed validation. Tabular Review recovered the preceding complete snapshot. Compare the visible Excel results, then use Reload Excel before saving further changes."
                System.Diagnostics.Debug.WriteLine("Tabular Review snapshot recovery: " & loadFailure.ToString())
            End If
            If initialRecovery Then
                _lastNotice = "The first review save stopped before activation. Tabular Review recovered its complete, validated snapshot. Review it and save the workbook to activate and retain that recovered state."
                System.Diagnostics.Debug.WriteLine("Tabular Review recovered one fully validated initial snapshot; commit marker still requires a successful save.")
            End If
            _snapshot = restored
            _activeSlot = committedSlot
            _loaded = Global.SharedLibrary.TabularReview.ReviewContract.Clone(restored.Project)
            _expectedFingerprint = restored.ProjectionFingerprint
            _expectedRows = restored.ProjectionRows
            _expectedColumns = restored.ProjectionColumns
            Dim projection As Microsoft.Office.Interop.Excel.Worksheet = FindProjectionSheet(restored)
            Try
                If projection Is Nothing Then
                    Throw New System.IO.InvalidDataException("The workbook contains review data but its review worksheet is missing. Restore that worksheet from a workbook copy; Tabular Review will not overwrite another sheet.")
                End If
                _visibleSheetName = projection.Name
                If Not ProjectionMatches(projection, _expectedRows, _expectedColumns, _expectedFingerprint) Then
                    If _lastNotice.Length > 0 Then _lastNotice &= System.Environment.NewLine
                    _lastNotice &= "Excel contains edits that the review has not imported. Use Reload Excel to reconcile them before editing the review."
                End If
            Finally
                ReleaseTemporary(projection)
            End Try
            If _lastNotice.Length > 0 Then
                Global.SharedLibrary.TabularReview.ReviewContract.AddAudit(_loaded, "workbook_load_notice", _lastNotice)
            End If
            Return Global.SharedLibrary.TabularReview.ReviewContract.Clone(_loaded)
        Finally
            ReleaseTemporary(store)
        End Try
    End Function

    ''' <summary>
    ''' Called only after explicit user approval. Archives the invalid/uncommitted
    ''' first store in-place, preserving all its cells and the visible projection.
    ''' A later save creates a separate, fresh store and review worksheet. The
    ''' recovery marker is deterministic and can be restored from an archived copy.
    ''' </summary>
    Public Sub ArchiveInvalidInitialSnapshot()
        EnsureAvailable()
        EnsureWritable()
        Dim store As Microsoft.Office.Interop.Excel.Worksheet = FindStoreSheet()
        If store Is Nothing Then Throw New System.InvalidOperationException("The interrupted review store is no longer present.")
        Try
            If CellText(store, 2, 2) <> "0" Then Throw New System.InvalidOperationException("A committed review database cannot be archived as an interrupted first save.")
            Dim leftState As System.String = CellText(store, 5, 4)
            Dim rightState As System.String = CellText(store, 5, 5)
            If leftState = "complete" AndAlso rightState = "complete" Then Throw New System.InvalidOperationException("Two complete uncommitted snapshots exist. Recover them manually; no data has been modified.")
            Dim firstSlot As System.Int32 = If(leftState = "complete", 1, If(rightState = "complete", 2, 0))
            If firstSlot <> 0 Then
                Dim recovered As StoredSnapshot = Nothing
                Try
                    recovered = ReadSnapshot(store, firstSlot)
                Catch ex As System.Exception
                    System.Diagnostics.Trace.WriteLine("[TabularReview:archive_uncommitted] invalid initial snapshot: " & ex.ToString())
                End Try
                If recovered IsNot Nothing Then Throw New System.InvalidOperationException("The first snapshot now validates. Reopen the existing review instead of archiving it.")
            ElseIf leftState <> "writing" AndAlso rightState <> "writing" Then
                Throw New System.InvalidOperationException("There is no interrupted first snapshot to archive.")
            End If
            If store.ProtectContents Then Throw New System.InvalidOperationException("Unprotect the review database sheet before archiving an interrupted first save.")
            Dim marker As System.String = "RedInk.TabularReview.ArchivedUncommitted.v1:" & System.Guid.NewGuid().ToString("N")
            PutText(store, 1, 1, marker)
            System.Diagnostics.Trace.WriteLine("[TabularReview:archive_uncommitted] " & store.Name & " preserved with marker " & marker)
            ' This adapter remains bound to the same workbook, not a different
            ' ActiveWorkbook. FindStoreSheet will now create a NEW store on save.
            _storeSheetName = ""
            _visibleSheetName = ""
            _snapshot = Nothing
            _loaded = Nothing
            _activeSlot = 0
            _committedSlotObserved = 0
            _expectedFingerprint = ""
            _expectedRows = 0
            _expectedColumns = 0
            _lastNotice = "The interrupted first review was preserved in an archived hidden worksheet. This new review starts empty. Save the workbook to retain both."
        Finally
            ReleaseTemporary(store)
        End Try
    End Sub

    Public Sub Save(project As Global.SharedLibrary.TabularReview.ReviewProject)
        EnsureAvailable()
        EnsureWritable()
        Global.SharedLibrary.TabularReview.ReviewContract.Validate(project)
        If _snapshot Is Nothing Then
            Dim existing As Global.SharedLibrary.TabularReview.ReviewProject = Load()
            If existing IsNot Nothing AndAlso existing.Id <> project.Id Then
                Throw New System.InvalidOperationException("This workbook already contains a different Tabular Review. Open its review or start the new review in another workbook.")
            End If
        End If
        If _snapshot IsNot Nothing AndAlso (project.Id <> _snapshot.Project.Id OrElse project.Revision <> _snapshot.Project.Revision) Then
            Throw New System.InvalidOperationException("The review has changed since this view loaded. Use Reload Excel before saving.")
        End If
        Dim candidate As Global.SharedLibrary.TabularReview.ReviewProject = Global.SharedLibrary.TabularReview.ReviewContract.Clone(project)
        candidate.SchemaVersion = 6
        candidate.Revision = If(_snapshot Is Nothing, 1L, _snapshot.Project.Revision + 1L)
        candidate.UpdatedUtc = System.DateTime.UtcNow.ToString("o")
        Dim columns As System.Collections.Generic.List(Of ProjectionColumn) = BuildColumns(candidate)
        Dim documents As System.Collections.Generic.List(Of Global.SharedLibrary.TabularReview.ReviewDocument) = ActiveDocuments(candidate)
        ValidateProjectionSize(documents.Count, columns.Count)

        Dim application As Microsoft.Office.Interop.Excel.Application = _workbook.Application
        Dim previousEvents As System.Boolean = application.EnableEvents
        Dim previousScreenUpdating As System.Boolean = application.ScreenUpdating
        Dim store As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim projection As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim backup As System.Collections.Generic.List(Of ProjectionBackupBlock) = Nothing
        Dim previousRows As System.Int32
        Dim previousColumns As System.Int32
        Dim nextRows As System.Int32 = documents.Count + 4
        Dim nextColumns As System.Int32 = columns.Count
        Dim committed As System.Boolean = False
        Dim saveStage As System.String = "checking existing review"
        Try
            store = FindStoreSheet()
            If _snapshot IsNot Nothing Then
                If store Is Nothing OrElse CellText(store, 2, 2) <> _committedSlotObserved.ToString(System.Globalization.CultureInfo.InvariantCulture) Then
                    Throw New System.InvalidOperationException("The workbook review database changed since this view loaded. Close and reopen Tabular Review before saving.")
                End If
                Dim currentSnapshot As StoredSnapshot = ReadSnapshot(store, _activeSlot)
                If Global.SharedLibrary.TabularReview.ReviewContract.Serialize(currentSnapshot) <> Global.SharedLibrary.TabularReview.ReviewContract.Serialize(_snapshot) Then
                    Throw New System.InvalidOperationException("Another change updated the workbook review database. Close and reopen Tabular Review before saving.")
                End If
                projection = FindProjectionSheet(_snapshot)
                If projection Is Nothing Then Throw New System.InvalidOperationException("The review worksheet is missing. Restore it before continuing.")
                If projection.ProtectContents Then Throw New System.InvalidOperationException("Unprotect the review worksheet before saving review changes.")
                If Not ProjectionMatches(projection, _expectedRows, _expectedColumns, _expectedFingerprint) Then
                    Throw New System.InvalidOperationException("Excel changed while the review was open. Use Reload Excel to import those edits; Tabular Review has not overwritten them.")
                End If
                previousRows = _expectedRows
                previousColumns = _expectedColumns
                backup = CaptureProjection(projection, previousRows, previousColumns)
            End If
            saveStage = "preparing Excel worksheets"
            application.EnableEvents = False
            application.ScreenUpdating = False
            Dim newProjection As System.Boolean = (projection Is Nothing)
            If newProjection Then projection = AddWorksheet("Tabular Review")
            _visibleSheetName = projection.Name
            If store Is Nothing Then
                store = AddWorksheet("_RI_TabularReview")
                PutText(store, 1, 1, StoreMarker)
                PutText(store, 1, 2, "1")
                PutText(store, 2, 1, "Committed snapshot")
                PutText(store, 2, 2, "0")
                PutText(store, 3, 1, "Keep this sheet: it contains complete review text, sources, history and recovery snapshots.")
            End If
            _storeSheetName = store.Name
            If store.ProtectContents Then Throw New System.InvalidOperationException("Unprotect the review data worksheet before saving review changes.")
            store.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVeryHidden

            saveStage = "writing visible review cells"
            WriteProjection(projection, candidate, columns, documents, previousRows, previousColumns)
            saveStage = "formatting the review worksheet"
            FormatProjection(projection, nextRows, nextColumns, If(newProjection, 0, previousColumns))
            ApplyQuestionFlagColors(projection, candidate, columns, documents, System.Math.Max(nextRows, previousRows), System.Math.Max(nextColumns, previousColumns))
            Dim prepared As New StoredSnapshot With {
                .Project = candidate,
                .WorksheetName = projection.Name,
                .WorksheetCodeName = SafeWorksheetCodeName(projection),
                .ProjectionRows = nextRows,
                .ProjectionColumns = nextColumns,
                .ProjectionFingerprint = ComputeFingerprint(projection, nextRows, nextColumns)
            }
            Dim inactiveSlot As System.Int32 = If(_activeSlot = 1, 2, 1)
            saveStage = "writing the review database"
            WriteSnapshot(store, inactiveSlot, prepared)
            saveStage = "checking the saved review data"
            Dim verified As StoredSnapshot = ReadSnapshot(store, inactiveSlot, True)
            If Global.SharedLibrary.TabularReview.ReviewContract.Serialize(verified) <> Global.SharedLibrary.TabularReview.ReviewContract.Serialize(prepared) Then
                Throw New System.IO.InvalidDataException("Excel did not preserve the complete review snapshot. Tabular Review kept the preceding snapshot.")
            End If
            ' The JSON/checksum alone are insufficient. Verify the identity
            ' against the very worksheet that was written, without obtaining an
            ' extra COM RCW that could invalidate the existing projection proxy.
            saveStage = "checking the review worksheet identity and cells"
            If Not System.String.Equals(verified.WorksheetName, projection.Name, System.StringComparison.OrdinalIgnoreCase) OrElse
               Not System.String.Equals(CellText(projection, 1, 1), ProjectionMarker & verified.Project.Id, System.StringComparison.Ordinal) Then
                Throw New System.IO.InvalidDataException("The review database does not identify its result worksheet.")
            End If
            If Not ProjectionMatches(projection, verified.ProjectionRows, verified.ProjectionColumns, verified.ProjectionFingerprint) Then
                Throw New System.IO.InvalidDataException("The review worksheet changed before its snapshot could be committed.")
            End If
            saveStage = "activating the verified review snapshot"
            PutText(store, 5, inactiveSlot + 3, "complete")
            Dim activated As StoredSnapshot = ReadSnapshot(store, inactiveSlot)
            If Global.SharedLibrary.TabularReview.ReviewContract.Serialize(activated) <> Global.SharedLibrary.TabularReview.ReviewContract.Serialize(prepared) Then
                Throw New System.IO.InvalidDataException("The review snapshot could not be verified after activation. The preceding snapshot is unchanged.")
            End If
            saveStage = "committing the verified review snapshot"
            PutText(store, 2, 2, inactiveSlot.ToString(System.Globalization.CultureInfo.InvariantCulture))
            committed = True
            If CellText(store, 2, 2) <> inactiveSlot.ToString(System.Globalization.CultureInfo.InvariantCulture) Then
                Throw New System.IO.InvalidDataException("Excel could not verify the review commit marker. Close the review and reopen it before continuing.")
            End If
            _snapshot = prepared
            _activeSlot = inactiveSlot
            _committedSlotObserved = inactiveSlot
            _loaded = candidate
            _expectedRows = nextRows
            _expectedColumns = nextColumns
            _expectedFingerprint = prepared.ProjectionFingerprint
            _lastNotice = ""
            project.SchemaVersion = candidate.SchemaVersion
            project.Revision = candidate.Revision
            project.UpdatedUtc = candidate.UpdatedUtc
        Catch ex As System.Exception
            System.Diagnostics.Trace.WriteLine("[TabularReview:save_failure] stage=" & saveStage & " | " & ex.ToString())
            If Not committed AndAlso _snapshot Is Nothing AndAlso store IsNot Nothing Then
                ' Never let a failed FIRST write permanently poison subsequent folder
                ' imports. Preserve the entire unfinished hidden store by renaming only
                ' its identifying marker. Do not delete its payload or visible results.
                Try
                    Dim archivedMarker As System.String = "RedInk.TabularReview.ArchivedUncommitted.v1:" & System.Guid.NewGuid().ToString("N")
                    PutText(store, 1, 1, archivedMarker)
                    _storeSheetName = ""
                    _visibleSheetName = ""
                    System.Diagnostics.Trace.WriteLine("[TabularReview:first_save_archived] " & archivedMarker)
                Catch archiveFailure As System.Exception
                    System.Diagnostics.Trace.WriteLine("[TabularReview:first_save_archive_failed] " & archiveFailure.ToString())
                    ' Leave the intact, identifiable store for explicit recovery on
                    ' the next launch; never silently replace its content.
                End Try
            End If
            If Not committed AndAlso projection IsNot Nothing AndAlso backup IsNot Nothing Then
                Try
                    RestoreProjection(projection, backup, previousRows, previousColumns, nextRows, nextColumns)
                    ApplyQuestionFlagColors(projection, _loaded, BuildColumns(_loaded), ActiveDocuments(_loaded), System.Math.Max(nextRows, previousRows), System.Math.Max(nextColumns, previousColumns))
                Catch rollbackFailure As System.Exception
                    System.Diagnostics.Debug.WriteLine("Tabular Review projection recovery failed: " & rollbackFailure.ToString())
                    Throw New System.InvalidOperationException("Excel could not finish the review save or restore its visible projection. The preceding complete review snapshot remains in the workbook. Keep this workbook and reopen Tabular Review before continuing.", New System.AggregateException(ex, rollbackFailure))
                End Try
            End If
            Throw New System.InvalidOperationException("Could not save the review (" & saveStage & "): " & ex.Message & ". Details were recorded in the diagnostic trace; existing review data was preserved.", ex)
        Finally
            ReleaseTemporary(store)
            ReleaseTemporary(projection)
            Try
                application.EnableEvents = previousEvents
            Finally
                application.ScreenUpdating = previousScreenUpdating
            End Try
        End Try
    End Sub

    Public Function RefreshFromWorksheet() As Global.SharedLibrary.TabularReview.ReviewProject
        EnsureAvailable()
        If _snapshot Is Nothing Then
            Dim restored As Global.SharedLibrary.TabularReview.ReviewProject = Load()
            If restored Is Nothing Then Return Nothing
        End If
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = FindProjectionSheet(_snapshot)
        If sheet Is Nothing Then Throw New System.InvalidOperationException("Restore the missing review worksheet before using Reload Excel.")
        Try
            Dim merged As Global.SharedLibrary.TabularReview.ReviewProject = Global.SharedLibrary.TabularReview.ReviewContract.Clone(_loaded)
            Dim columns As System.Collections.Generic.List(Of ProjectionColumn) = BuildColumns(_loaded)
            Dim originalDocuments As System.Collections.Generic.List(Of Global.SharedLibrary.TabularReview.ReviewDocument) = ActiveDocuments(_loaded)
            Dim extent As System.Tuple(Of System.Int32, System.Int32) = ContentExtent(sheet)
            If extent.Item1 <> originalDocuments.Count + 4 OrElse extent.Item2 <> columns.Count Then
                Throw New System.InvalidOperationException("Excel's review layout changed. Restore added/deleted rows or columns, then use the review's document and criteria controls to change its structure. Reload Excel will not discard worksheet cells.")
            End If
            Dim readFingerprint As System.String = ComputeFingerprint(sheet, extent.Item1, extent.Item2)
            Dim actualColumns As New System.Collections.Generic.Dictionary(Of System.String, System.Int32)(System.StringComparer.Ordinal)
            For column As System.Int32 = 1 To extent.Item2
                Dim id As System.String = ReadReviewCell(sheet, 4, column)
                If actualColumns.ContainsKey(id) Then Throw New System.InvalidOperationException("Excel contains duplicate review column IDs. Restore hidden row 4 before reloading.")
                actualColumns.Add(id, column)
            Next
            For Each definition As ProjectionColumn In columns
                If Not actualColumns.ContainsKey(definition.Id) Then Throw New System.InvalidOperationException("A review column ID is missing. Restore hidden row 4 before reloading; use the criteria editor to add or remove columns.")
            Next
            If actualColumns("d:id") <> 1 Then Throw New System.InvalidOperationException("Keep the review's hidden document ID in column A. Restore that column before reloading.")
            Dim oldTitle As System.String = merged.Title
            merged.Title = EditableText(merged.Title, ReadReviewCell(sheet, 1, 2), "review title")
            If System.String.IsNullOrWhiteSpace(merged.Title) Then Throw New System.InvalidOperationException("Enter a review title before reloading Excel.")
            If merged.Title <> oldTitle Then Global.SharedLibrary.TabularReview.ReviewContract.AddAudit(merged, "excel_title_edit", Global.SharedLibrary.TabularReview.ReviewContract.Serialize(New With {.previous = oldTitle, .current = merged.Title}))

            Dim changedCriteria As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each definition As ProjectionColumn In columns
                Dim actualColumn As System.Int32 = actualColumns(definition.Id)
                If definition.CriterionId.Length = 0 OrElse definition.Field <> "answer" Then
                    RequireUnchanged(ReadReviewCell(sheet, 2, actualColumn), definition.Title, "column heading")
                    RequireUnchanged(ReadReviewCell(sheet, 3, actualColumn), definition.Instruction, "column description")
                    Continue For
                End If
                Dim criterion As Global.SharedLibrary.TabularReview.ReviewCriterion = merged.Criteria.Find(Function(item As Global.SharedLibrary.TabularReview.ReviewCriterion) item.Id = definition.CriterionId)
                Dim title As System.String = EditableText(criterion.Title, ReadReviewCell(sheet, 2, actualColumn), "criterion title")
                Dim instruction As System.String = EditableText(criterion.Instruction, ReadReviewCell(sheet, 3, actualColumn), "criterion question")
                If System.String.IsNullOrWhiteSpace(title) OrElse (criterion.Kind = "analysis" AndAlso System.String.IsNullOrWhiteSpace(instruction)) Then
                    Throw New System.InvalidOperationException("Every analysis column needs its short title in row 2 and complete question in row 3.")
                End If
                If title <> criterion.Title OrElse instruction <> criterion.Instruction Then
                    Global.SharedLibrary.TabularReview.ReviewContract.AddAudit(merged, "excel_criterion_edit", Global.SharedLibrary.TabularReview.ReviewContract.Serialize(New With {.previousTitle = criterion.Title, .previousInstruction = criterion.Instruction, .title = title, .instruction = instruction}), criterionId:=criterion.Id)
                    criterion.Title = title
                    criterion.Instruction = instruction
                    criterion.Revision += 1
                    changedCriteria.Add(criterion.Id)
                    For Each document As Global.SharedLibrary.TabularReview.ReviewDocument In merged.Documents
                        Dim answer As Global.SharedLibrary.TabularReview.ReviewAnswer = Nothing
                        If Not criterion.AppliesTo(document) Then
                            If document.Answers.TryGetValue(criterion.Id, answer) Then answer.Status = "stale"
                            Continue For
                        End If
                        If document.Answers.TryGetValue(criterion.Id, answer) Then Global.SharedLibrary.TabularReview.ReviewContract.ArchiveAnswer(answer, "Excel criterion edit")
                        Global.SharedLibrary.TabularReview.ReviewContract.Invalidate(document, criterion.Id)
                    Next
                End If
            Next

            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim orderedDocuments As New System.Collections.Generic.List(Of Global.SharedLibrary.TabularReview.ReviewDocument)()
            Dim changedAnswers As System.Int32
            For firstRow As System.Int32 = 5 To extent.Item1 Step BlockRows
                Dim count As System.Int32 = System.Math.Min(BlockRows, extent.Item1 - firstRow + 1)
                Dim range As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, firstRow, 1, firstRow + count - 1, extent.Item2)
                Try
                    Dim values As System.Object = range.Value2
                    Dim formulas As System.Object = range.Formula
                    For offset As System.Int32 = 0 To count - 1
                        Dim documentId As System.String = ScalarText(ArrayItem(values, offset, 0))
                        Dim original As Global.SharedLibrary.TabularReview.ReviewDocument = originalDocuments.Find(Function(item As Global.SharedLibrary.TabularReview.ReviewDocument) item.Id = documentId)
                        If original Is Nothing OrElse Not seen.Add(documentId) Then Throw New System.InvalidOperationException("Excel contains a missing, duplicate or unknown document ID. Restore hidden column A; when sorting, include the entire review row.")
                        Dim document As Global.SharedLibrary.TabularReview.ReviewDocument = merged.Documents.Find(Function(item As Global.SharedLibrary.TabularReview.ReviewDocument) item.Id = documentId)
                        If document IsNot Nothing Then orderedDocuments.Add(document)
                        Dim rowValues As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
                        For Each definition As ProjectionColumn In columns
                            Dim column As System.Int32 = actualColumns(definition.Id)
                            Dim formula As System.String = ScalarText(ArrayItem(formulas, offset, column - 1))
                            If formula.StartsWith("=", System.StringComparison.Ordinal) Then AssertNoFormula(sheet, firstRow + offset, column)
                            Dim text As System.String = ScalarText(ArrayItem(values, offset, column - 1))
                            rowValues.Add(definition.Id, text)
                            If document Is Nothing OrElse definition.Field = "id" OrElse definition.Field = "file" OrElse definition.Field = "path" OrElse definition.Field = "extraction" OrElse definition.Field = "import" OrElse definition.Field = "sources" Then
                                RequireUnchanged(text, DocumentValue(definition, original), If(definition.Field = "sources", "source reference/status", "document identity/source field"))
                            End If
                        Next

                        If document Is Nothing Then Continue For
                        Dim previousMetadata As System.String = DocumentMetadata(original)
                        If rowValues.ContainsKey("d:category") Then document.Category = EditableText(original.Category, rowValues("d:category"), "category").Trim()
                        If rowValues.ContainsKey("d:group") Then document.GroupName = EditableText(original.GroupName, rowValues("d:group"), "group").Trim()
                        document.Number = EditableText(original.Number, rowValues("d:number"), "document number")
                        document.Priority = EditableText(original.Priority, rowValues("d:priority"), "priority")
                        document.Reviewer = EditableText(original.Reviewer, rowValues("d:reviewer"), "reviewer")
                        document.Notes = EditableText(original.Notes, rowValues("d:notes"), "review notes")
                        Dim originalFlags As System.String = System.String.Join(System.Environment.NewLine, original.Flags)
                        Dim newFlags As System.String = EditableText(originalFlags, rowValues("d:flags"), "flags")
                        If newFlags <> originalFlags Then document.Flags = ParseFlags(newFlags)
                        Dim answerEdited As System.Boolean = False
                        For Each criterion As Global.SharedLibrary.TabularReview.ReviewCriterion In merged.Criteria
                            If criterion.Archived Then Continue For
                            If Not criterion.AppliesTo(original) OrElse Not criterion.AppliesTo(document) Then
                                For Each field As ProjectionColumn In columns
                                    If field.CriterionId = criterion.Id Then RequireUnchanged(rowValues(field.Id), DocumentValue(field, original), "non-applicable answer (reload category changes before editing answers)")
                                Next
                                Continue For
                            End If
                            Dim previous As Global.SharedLibrary.TabularReview.ReviewAnswer = Nothing
                            original.Answers.TryGetValue(criterion.Id, previous)
                            Dim text As System.String = EditableText(If(previous Is Nothing, "", previous.Text), rowValues("c:" & criterion.Id & ":answer"), "answer")
                            Dim assessment As System.String = If(previous Is Nothing, "", previous.Assessment)
                            If criterion.Kind = "analysis" Then assessment = EditableText(assessment, rowValues("c:" & criterion.Id & ":assessment"), "assessment")
                            Dim oldFlag As System.String = If(previous Is Nothing, "", previous.RiskFlag)
                            Dim newFlag As System.String = oldFlag
                            If criterion.FlagValues().Count > 0 Then
                                newFlag = EditableText(oldFlag, rowValues("c:" & criterion.Id & ":riskFlag"), "question flag")
                                If newFlag <> oldFlag AndAlso newFlag.Length > 0 AndAlso Not criterion.FlagValues().Contains(newFlag) Then Throw New System.InvalidOperationException("Choose a configured flag for " & criterion.Title)
                            End If
                            If newFlag = oldFlag AndAlso text = If(previous Is Nothing, "", previous.Text) AndAlso assessment = If(previous Is Nothing, "", previous.Assessment) Then Continue For
                            If (criterion.Kind = "priority" OrElse criterion.Kind = "flag") AndAlso criterion.Options.Count > 0 AndAlso text.Length > 0 AndAlso Not criterion.Options.Contains(text) Then
                                Throw New System.InvalidOperationException("Choose one of the exact values configured for the '" & criterion.Title & "' column before reloading Excel.")
                            End If
                            Dim answer As Global.SharedLibrary.TabularReview.ReviewAnswer = Nothing
                            If Not document.Answers.TryGetValue(criterion.Id, answer) Then
                                answer = New Global.SharedLibrary.TabularReview.ReviewAnswer With {.CriterionRevision = criterion.Revision, .SourceHash = document.SourceHash}
                                document.Answers.Add(criterion.Id, answer)
                            Else
                                Global.SharedLibrary.TabularReview.ReviewContract.ArchiveAnswer(answer, "Excel answer edit")
                            End If
                            answer.RiskFlag = newFlag
                            answer.Text = text
                            answer.Assessment = assessment
                            answer.UserEdited = True
                            answer.UpdatedUtc = System.DateTime.UtcNow.ToString("o")
                            answer.Status = If(changedCriteria.Contains(criterion.Id), "stale", "draft")
                            answer.Error = ""
                            changedAnswers += 1
                            answerEdited = True
                            Global.SharedLibrary.TabularReview.ReviewContract.AddAudit(merged, "excel_answer_edit", "The reviewer edited the worksheet answer/assessment. The answer history retains the preceding text and references.", document.Id, criterion.Id)
                        Next
                        Dim requestedStatus As System.String = EditableText(original.ReviewStatus, rowValues("d:status"), "review status").Trim().ToLowerInvariant().Replace(" ", "_")
                        If Not Global.SharedLibrary.TabularReview.ReviewContract.Contains(New System.String() {"unreviewed", "in_review", "reviewed"}, requestedStatus) Then
                            Throw New System.InvalidOperationException("Use unreviewed, in_review or reviewed in the Review status column before reloading Excel.")
                        End If
                        If requestedStatus <> original.ReviewStatus Then
                            If requestedStatus = "reviewed" Then EnsureDocumentCanBeReviewed(merged, document)
                            document.ReviewStatus = requestedStatus
                            document.LastReviewedUtc = If(requestedStatus = "reviewed", System.DateTime.UtcNow.ToString("o"), "")
                        ElseIf answerEdited OrElse document.Category <> original.Category Then
                            document.ReviewStatus = "in_review"
                            document.LastReviewedUtc = ""
                        End If
                        Dim currentMetadata As System.String = DocumentMetadata(document)
                        If previousMetadata <> currentMetadata Then Global.SharedLibrary.TabularReview.ReviewContract.AddAudit(merged, "excel_document_edit", Global.SharedLibrary.TabularReview.ReviewContract.Serialize(New With {.previous = previousMetadata, .current = currentMetadata}), document.Id)
                    Next
                Finally
                    ReleaseTemporary(range)
                End Try
            Next
            If seen.Count <> originalDocuments.Count Then Throw New System.InvalidOperationException("A review document row is missing. Restore it before reloading Excel.")
            For Each document As Global.SharedLibrary.TabularReview.ReviewDocument In merged.Documents
                If document.Archived Then orderedDocuments.Add(document)
            Next
            merged.Documents = orderedDocuments
            Dim orderedCriteria As New System.Collections.Generic.List(Of Global.SharedLibrary.TabularReview.ReviewCriterion)()
            For Each criterion As Global.SharedLibrary.TabularReview.ReviewCriterion In merged.Criteria
                If Not criterion.Archived Then orderedCriteria.Add(criterion)
            Next
            orderedCriteria.Sort(Function(left As Global.SharedLibrary.TabularReview.ReviewCriterion, right As Global.SharedLibrary.TabularReview.ReviewCriterion) actualColumns("c:" & left.Id & ":answer").CompareTo(actualColumns("c:" & right.Id & ":answer")))
            For Each criterion As Global.SharedLibrary.TabularReview.ReviewCriterion In merged.Criteria
                If criterion.Archived Then orderedCriteria.Add(criterion)
            Next
            merged.Criteria = orderedCriteria
            Global.SharedLibrary.TabularReview.ReviewContract.Validate(merged)
            Global.SharedLibrary.TabularReview.ReviewContract.AddAudit(merged, "excel_reload", "Reconciled worksheet edits by stable document/criterion IDs; changed criteria=" & changedCriteria.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & "; edited answers=" & changedAnswers.ToString(System.Globalization.CultureInfo.InvariantCulture) & ".")
            Dim previousFingerprint As System.String = _expectedFingerprint
            Dim previousRows As System.Int32 = _expectedRows
            Dim previousColumns As System.Int32 = _expectedColumns
            If Not ProjectionMatches(sheet, extent.Item1, extent.Item2, readFingerprint) Then Throw New System.InvalidOperationException("Worksheet cells changed during reload. Let Excel finish synchronizing, then reload again.")
            _expectedFingerprint = readFingerprint
            _expectedRows = extent.Item1
            _expectedColumns = extent.Item2
            Try
                Save(merged)
            Catch
                _expectedFingerprint = previousFingerprint
                _expectedRows = previousRows
                _expectedColumns = previousColumns
                Throw
            End Try
            Return merged
        Finally
            ReleaseTemporary(sheet)
        End Try
    End Function

    Public Function ImportSelection(range As Microsoft.Office.Interop.Excel.Range,
                                    Optional sourceRoot As System.String = "",
                                    Optional fileColumnIndex As System.Int32 = 0) As Global.SharedLibrary.TabularReview.ReviewProject
        EnsureAvailable()
        If range Is Nothing Then Throw New System.ArgumentNullException(NameOf(range))
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = range.Worksheet
        Dim selected As Microsoft.Office.Interop.Excel.Range = range
        Dim ownsSelected As System.Boolean = False
        Dim areas As Microsoft.Office.Interop.Excel.Areas = Nothing
        Try
            If Not SameComIdentity(sheet.Parent, _workbook) Then Throw New System.InvalidOperationException("Select the source table in this review's workbook, then import it again.")
            areas = range.Areas
            If areas.Count <> 1 Then Throw New System.InvalidOperationException("Select one rectangular source table, including its column headings.")
            Dim rowCollection As Microsoft.Office.Interop.Excel.Range = range.Rows
            Dim columnCollection As Microsoft.Office.Interop.Excel.Range = range.Columns
            Try
                If rowCollection.Count = 1 AndAlso columnCollection.Count = 1 Then
                    selected = range.CurrentRegion
                    ownsSelected = True
                End If
            Finally
                ReleaseTemporary(rowCollection)
                ReleaseTemporary(columnCollection)
            End Try
            Dim bounds As System.Tuple(Of System.Int32, System.Int32) = SelectionExtent(selected)
            Dim firstRow As System.Int32 = selected.Row
            Dim firstColumn As System.Int32 = selected.Column
            Dim lastRow As System.Int32 = bounds.Item1
            Dim lastColumn As System.Int32 = bounds.Item2
            If lastRow <= firstRow OrElse lastColumn < firstColumn Then Throw New System.InvalidOperationException("Select a table with headings in its first row and at least one result row below them.")
            Dim countColumns As System.Int32 = lastColumn - firstColumn + 1
            If fileColumnIndex < 0 OrElse fileColumnIndex > countColumns Then Throw New System.ArgumentOutOfRangeException(NameOf(fileColumnIndex), "The source-file column must lie inside the selected table.")
            Dim headers As New System.Collections.Generic.List(Of System.String)()
            For column As System.Int32 = firstColumn To lastColumn
                headers.Add(ImportCellText(sheet, firstRow, column))
            Next
            If fileColumnIndex = 0 Then fileColumnIndex = DetectFileColumn(selected, headers, firstColumn, lastColumn)
            Dim detectedRoot As System.String = ""
            Dim summaryStart As System.Int32 = FindExtractorSummary(sheet, firstRow + 1, System.Math.Min(MaximumRows - 2, lastRow + 12), firstColumn)
            If summaryStart > 0 Then
                detectedRoot = CellText(sheet, summaryStart, firstColumn + 1)
                If summaryStart <= lastRow Then lastRow = summaryStart - 1
            End If
            Dim normalizedRoot As System.String = NormalizeImportRoot(If(System.String.IsNullOrWhiteSpace(sourceRoot), detectedRoot, sourceRoot))
            If System.String.IsNullOrWhiteSpace(sourceRoot) AndAlso detectedRoot.Length > 0 AndAlso fileColumnIndex > 0 AndAlso normalizedRoot.Length > 0 Then
                Dim firstSource As System.String = ImportCellText(sheet, firstRow + 1, firstColumn + fileColumnIndex - 1)
                ' Data Extractor's Directory field contains the complete source path
                ' for a single-file export. Match the actual File cell, not a list of
                ' extensions; an explicitly supplied folder always remains a folder.
                If firstSource.Length > 0 AndAlso System.String.Equals(System.IO.Path.GetFileName(normalizedRoot), System.IO.Path.GetFileName(firstSource), System.StringComparison.OrdinalIgnoreCase) Then
                    normalizedRoot = If(System.IO.Path.GetDirectoryName(normalizedRoot), "")
                End If
            End If
            Dim project As New Global.SharedLibrary.TabularReview.ReviewProject With {
                .Title = sheet.Name & " — Tabular Review",
                .SourceRoot = normalizedRoot
            }
            For index As System.Int32 = 0 To headers.Count - 1
                Dim title As System.String = If(System.String.IsNullOrWhiteSpace(headers(index)), "Column " & (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), headers(index))
                project.Criteria.Add(New Global.SharedLibrary.TabularReview.ReviewCriterion With {
                    .Title = title,
                    .Instruction = If(index + 1 = fileColumnIndex, "", title),
                    .Kind = If(index + 1 = fileColumnIndex, "text", "analysis")
                })
            Next
            ValidateProjectionSize(lastRow - firstRow, BuildColumns(project).Count)
            Dim missingSources As System.Int32
            Dim duplicateSources As System.Int32
            Dim sourceCounts As New System.Collections.Generic.Dictionary(Of System.String, System.Int32)(System.StringComparer.OrdinalIgnoreCase)
            For firstDataRow As System.Int32 = firstRow + 1 To lastRow Step BlockRows
                Dim count As System.Int32 = System.Math.Min(BlockRows, lastRow - firstDataRow + 1)
                Dim dataRange As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, firstDataRow, firstColumn, firstDataRow + count - 1, lastColumn)
                Try
                    Dim rawData As System.Object = dataRange.Value2
                    For rowOffset As System.Int32 = 0 To count - 1
                        Dim row As System.Int32 = firstDataRow + rowOffset
                        Dim values As New System.Collections.Generic.List(Of System.String)()
                        Dim hasValue As System.Boolean = False
                        For column As System.Int32 = firstColumn To lastColumn
                            Dim raw As System.Object = ArrayItem(rawData, rowOffset, column - firstColumn)
                            Dim value As System.String = If(raw Is Nothing OrElse TypeOf raw Is System.String, ScalarText(raw), ImportCellText(sheet, row, column))
                            values.Add(value)
                            If value.Length > 0 Then hasValue = True
                        Next
                        If Not hasValue Then Continue For
                        Dim sourceText As System.String = If(fileColumnIndex = 0, "", values(fileColumnIndex - 1))
                        Dim hyperlink As System.String = If(fileColumnIndex = 0, "", ImportHyperlink(sheet, row, firstColumn + fileColumnIndex - 1))
                        Dim sourcePath As System.String = ""
                        Dim sourceNotice As System.String = ""
                        ResolveImportedSource(sourceText, hyperlink, normalizedRoot, _workbook.Path, sourcePath, sourceNotice)
                        If sourcePath.Length = 0 Then missingSources += 1
                        If sourcePath.Length > 0 Then
                            If sourceCounts.ContainsKey(sourcePath) Then
                                sourceCounts(sourcePath) += 1
                                duplicateSources += 1
                            Else
                                sourceCounts.Add(sourcePath, 1)
                            End If
                        End If
                        Dim relative As System.String = ImportedRelativePath(sourcePath, normalizedRoot)
                        Dim document As New Global.SharedLibrary.TabularReview.ReviewDocument With {
                            .FileName = If(sourcePath.Length = 0, If(sourceText.Length = 0, "Imported row " & row.ToString(System.Globalization.CultureInfo.InvariantCulture), sourceText), System.IO.Path.GetFileName(sourcePath)),
                            .Number = If(relative.Length > 0, relative, If(sourceText.Length > 0, sourceText, row.ToString(System.Globalization.CultureInfo.InvariantCulture))),
                            .SourcePath = sourcePath,
                            .RelativePath = relative,
                            .ExtractionStatus = "imported",
                            .ExtractionMessage = sourceNotice,
                            .ImportedFrom = _workbook.Name & " / " & sheet.Name & " / row " & row.ToString(System.Globalization.CultureInfo.InvariantCulture)
                        }
                        For index As System.Int32 = 0 To project.Criteria.Count - 1
                            document.Answers.Add(project.Criteria(index).Id, New Global.SharedLibrary.TabularReview.ReviewAnswer With {
                                .Text = values(index), .Status = "imported", .CriterionRevision = project.Criteria(index).Revision,
                                .UpdatedUtc = System.DateTime.UtcNow.ToString("o"), .ModelKey = "imported"
                            })
                        Next
                        project.Documents.Add(document)
                    Next
                Finally
                    ReleaseTemporary(dataRange)
                End Try
            Next
            If project.Documents.Count = 0 Then Throw New System.InvalidOperationException("The selected table contains no non-empty result rows.")
            Global.SharedLibrary.TabularReview.ReviewContract.AddAudit(project, "import_excel", Global.SharedLibrary.TabularReview.ReviewContract.Serialize(New With {
                .workbook = _workbook.Name, .worksheet = sheet.Name, .headerRow = firstRow, .firstColumn = firstColumn, .lastRow = lastRow, .lastColumn = lastColumn,
                .headers = headers, .sourceFileColumn = fileColumnIndex, .detectedSourceRoot = detectedRoot,
                .unresolvedSources = missingSources, .additionalRowsForSameSource = duplicateSources,
                .notice = "The import retained every result cell and kept repeated source rows separate with row provenance. The original worksheet remains unchanged. Analysis questions initially equal the original column headings; edit them before generating new analysis. Imported answers have no fabricated evidence. Load each source document before reviewing citations or running AI."
            }))
            Global.SharedLibrary.TabularReview.ReviewContract.Validate(project)
            Return project
        Finally
            ReleaseTemporary(areas)
            If ownsSelected Then ReleaseTemporary(selected)
            ReleaseTemporary(sheet)
        End Try
    End Function

    Private Shared Sub RequireUnchanged(actual As System.String, previous As System.String, caption As System.String)
        If actual <> VisibleText(previous) Then
            Throw New System.InvalidOperationException("Excel changed a managed " & caption & ". Restore that field before reloading. Include hidden column A when sorting rows; use Tabular Review to change sources, evidence or column structure.")
        End If
    End Sub

    Private Shared Function DocumentMetadata(document As Global.SharedLibrary.TabularReview.ReviewDocument) As System.String
        Return Global.SharedLibrary.TabularReview.ReviewContract.Serialize(New With {
            .category = document.Category, .groupName = document.GroupName, .number = document.Number, .reviewStatus = document.ReviewStatus, .priority = document.Priority,
            .flags = document.Flags, .reviewer = document.Reviewer, .notes = document.Notes, .lastReviewedUtc = document.LastReviewedUtc
        })
    End Function

    Private Shared Sub EnsureDocumentCanBeReviewed(project As Global.SharedLibrary.TabularReview.ReviewProject, document As Global.SharedLibrary.TabularReview.ReviewDocument)
        For Each criterion As Global.SharedLibrary.TabularReview.ReviewCriterion In project.Criteria
            If criterion.Archived OrElse criterion.Kind = "text" OrElse Not criterion.AppliesTo(document) Then Continue For
            Dim answer As Global.SharedLibrary.TabularReview.ReviewAnswer = Nothing
            If Not document.Answers.TryGetValue(criterion.Id, answer) OrElse System.String.IsNullOrWhiteSpace(answer.Text) OrElse answer.Status = "stale" OrElse answer.Status = "error" OrElse answer.Status = "pending" OrElse (criterion.FlagValues().Count > 0 AndAlso Not criterion.FlagValues().Contains(answer.RiskFlag)) Then
                Throw New System.InvalidOperationException("Resolve empty, stale or failed answers before marking this document reviewed.")
            End If
        Next
    End Sub

    Private Shared Function ParseFlags(text As System.String) As System.Collections.Generic.List(Of System.String)
        Dim flags As New System.Collections.Generic.List(Of System.String)()
        For Each flag As System.String In text.Replace(Microsoft.VisualBasic.vbCrLf, Microsoft.VisualBasic.vbLf).Replace(Microsoft.VisualBasic.vbCr, Microsoft.VisualBasic.vbLf).Split(New System.Char() {Microsoft.VisualBasic.ChrW(10)})
            If Not System.String.IsNullOrWhiteSpace(flag) Then flags.Add(flag)
        Next
        Return flags
    End Function

    Private Shared Function ReadReviewCell(sheet As Microsoft.Office.Interop.Excel.Worksheet, row As System.Int32, column As System.Int32) As System.String
        AssertNoFormula(sheet, row, column)
        Return CellText(sheet, row, column)
    End Function

    Private Shared Sub AssertNoFormula(sheet As Microsoft.Office.Interop.Excel.Worksheet, row As System.Int32, column As System.Int32)
        Dim cell As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, row, column, row, column)
        Try
            If CBool(cell.HasFormula) Then Throw New System.InvalidOperationException("Review fields store literal text. Replace worksheet formulas in the review with their intended text before using Reload Excel.")
        Finally
            ReleaseTemporary(cell)
        End Try
    End Sub

    Private Shared Function SelectionExtent(range As Microsoft.Office.Interop.Excel.Range) As System.Tuple(Of System.Int32, System.Int32)
        Dim rowEnd As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim columnEnd As Microsoft.Office.Interop.Excel.Range = Nothing
        Try
            rowEnd = range.Find(What:="*", LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, MatchCase:=False, SearchFormat:=False)
            columnEnd = range.Find(What:="*", LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, MatchCase:=False, SearchFormat:=False)
            Return System.Tuple.Create(If(rowEnd Is Nothing, 0, rowEnd.Row), If(columnEnd Is Nothing, 0, columnEnd.Column))
        Finally
            ReleaseTemporary(rowEnd)
            ReleaseTemporary(columnEnd)
        End Try
    End Function

    Private Shared Function ImportCellText(sheet As Microsoft.Office.Interop.Excel.Worksheet, row As System.Int32, column As System.Int32) As System.String
        Dim cell As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, row, column, row, column)
        Try
            Dim value As System.Object = cell.Value2
            If value Is Nothing OrElse TypeOf value Is System.String Then Return ScalarText(value)
            Dim displayed As System.String = ScalarText(cell.Text)
            If displayed.Length > 0 AndAlso displayed.Trim("#"c).Length > 0 Then Return displayed
            Return ScalarText(value)
        Finally
            ReleaseTemporary(cell)
        End Try
    End Function

    Private Shared Function DetectFileColumn(range As Microsoft.Office.Interop.Excel.Range, headers As System.Collections.Generic.List(Of System.String), firstColumn As System.Int32, lastColumn As System.Int32) As System.Int32
        Dim candidates As New System.Collections.Generic.HashSet(Of System.Int32)()
        For index As System.Int32 = 0 To headers.Count - 1
            Dim normalized As System.String = headers(index).Trim().Replace(" ", "").Replace("_", "").Replace("-", "").ToLowerInvariant()
            If normalized = "file" OrElse normalized = "filename" OrElse normalized = "filepath" OrElse normalized = "sourcefile" OrElse normalized = "sourcepath" OrElse normalized = "documentpath" Then candidates.Add(index + 1)
        Next
        If candidates.Count = 1 Then
            For Each candidate As System.Int32 In candidates
                Return candidate
            Next
        End If
        If candidates.Count > 1 Then Return 0
        Dim links As Microsoft.Office.Interop.Excel.Hyperlinks = range.Hyperlinks
        Try
            For index As System.Int32 = 1 To links.Count
                Dim link As Microsoft.Office.Interop.Excel.Hyperlink = links.Item(index)
                Dim linkedRange As Microsoft.Office.Interop.Excel.Range = Nothing
                Try
                    linkedRange = link.Range
                    If linkedRange IsNot Nothing AndAlso linkedRange.Column >= firstColumn AndAlso linkedRange.Column <= lastColumn AndAlso Not System.String.IsNullOrEmpty(link.Address) Then
                        If ResolveLocalPath(link.Address, "").Length > 0 OrElse Not HasNonFileScheme(link.Address) Then candidates.Add(linkedRange.Column - firstColumn + 1)
                    End If
                Finally
                    ReleaseTemporary(linkedRange)
                    ReleaseTemporary(link)
                End Try
            Next
        Finally
            ReleaseTemporary(links)
        End Try
        If candidates.Count = 1 Then
            For Each candidate As System.Int32 In candidates
                Return candidate
            Next
        End If
        Return 0
    End Function

    Private Shared Function ImportHyperlink(sheet As Microsoft.Office.Interop.Excel.Worksheet, row As System.Int32, column As System.Int32) As System.String
        Dim cell As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, row, column, row, column)
        Dim links As Microsoft.Office.Interop.Excel.Hyperlinks = Nothing
        Dim link As Microsoft.Office.Interop.Excel.Hyperlink = Nothing
        Try
            links = cell.Hyperlinks
            If links.Count <> 1 Then Return ""
            link = links.Item(1)
            Return If(link.Address, "")
        Finally
            ReleaseTemporary(link)
            ReleaseTemporary(links)
            ReleaseTemporary(cell)
        End Try
    End Function

    Private Shared Function FindExtractorSummary(sheet As Microsoft.Office.Interop.Excel.Worksheet, firstRow As System.Int32, lastRow As System.Int32, firstColumn As System.Int32) As System.Int32
        If lastRow < firstRow Then Return 0
        Dim search As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, firstRow, firstColumn, lastRow, firstColumn)
        Dim found As Microsoft.Office.Interop.Excel.Range = Nothing
        Try
            found = search.Find(What:="Directory:", LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False, SearchFormat:=False)
            If found Is Nothing Then Return 0
            If CellText(sheet, found.Row + 1, firstColumn) = "Files processed:" AndAlso CellText(sheet, found.Row + 2, firstColumn) = "Files failed:" Then Return found.Row
            Return 0
        Finally
            ReleaseTemporary(found)
            ReleaseTemporary(search)
        End Try
    End Function

    Private Shared Function NormalizeImportRoot(value As System.String) As System.String
        If System.String.IsNullOrWhiteSpace(value) Then Return ""
        Dim expanded As System.String = Global.SharedLibrary.SharedLibrary.SharedMethods.ExpandEnvironmentVariables(value.Trim())
        Dim absolute As System.String = ResolveLocalPath(expanded, "")
        If absolute.Length = 0 Then Throw New System.InvalidOperationException("Enter an absolute local folder or UNC path for the imported source documents.")
        Return absolute
    End Function

    Private Shared Sub ResolveImportedSource(sourceText As System.String, hyperlink As System.String, sourceRoot As System.String, workbookDirectory As System.String,
                                             ByRef sourcePath As System.String, ByRef notice As System.String)
        Dim fromText As System.String = ResolveLocalPath(sourceText, sourceRoot)
        Dim fromLink As System.String = ResolveLocalPath(hyperlink, If(System.String.IsNullOrEmpty(workbookDirectory), sourceRoot, workbookDirectory))
        If fromText.Length > 0 AndAlso fromLink.Length > 0 AndAlso Not System.String.Equals(fromText, fromLink, System.StringComparison.OrdinalIgnoreCase) Then
            sourcePath = ""
            notice = "The cell text and hyperlink identify different files. Relink this row to its exact source before analysis."
            Return
        End If
        sourcePath = If(fromLink.Length > 0, fromLink, fromText)
        notice = If(sourcePath.Length = 0,
                    "No unambiguous local or UNC source path was found. Choose the source folder or relink this row before analysis; imported cell text remains available.",
                    "Imported result; source availability and text have not yet been checked. Load the source before generating or verifying answers.")
    End Sub

    Private Shared Function ResolveLocalPath(value As System.String, baseDirectory As System.String) As System.String
        If System.String.IsNullOrWhiteSpace(value) Then Return ""
        Try
            Dim candidate As System.String = value.Trim()
            If candidate.StartsWith("file:", System.StringComparison.OrdinalIgnoreCase) Then
                Dim uri As System.Uri = Nothing
                If Not System.Uri.TryCreate(candidate, System.UriKind.Absolute, uri) OrElse Not uri.IsFile Then Return ""
                candidate = uri.LocalPath
            ElseIf HasNonFileScheme(candidate) Then
                Return ""
            End If
            If candidate.StartsWith("\\?\", System.StringComparison.Ordinal) OrElse candidate.StartsWith("\\.\", System.StringComparison.Ordinal) Then Return ""
            Dim isDrive As System.Boolean = candidate.Length >= 3 AndAlso System.Char.IsLetter(candidate(0)) AndAlso candidate(1) = ":"c AndAlso (candidate(2) = "\"c OrElse candidate(2) = "/"c)
            Dim isUnc As System.Boolean = candidate.StartsWith("\\", System.StringComparison.Ordinal)
            If Not isDrive AndAlso Not isUnc Then
                If System.String.IsNullOrEmpty(baseDirectory) OrElse System.IO.Path.IsPathRooted(candidate) Then Return ""
                candidate = System.IO.Path.Combine(baseDirectory, candidate)
            End If
            Return System.IO.Path.GetFullPath(candidate)
        Catch ex As System.Exception When TypeOf ex Is System.ArgumentException OrElse TypeOf ex Is System.NotSupportedException OrElse TypeOf ex Is System.IO.PathTooLongException OrElse TypeOf ex Is System.Security.SecurityException
            Return ""
        End Try
    End Function

    Private Shared Function HasNonFileScheme(value As System.String) As System.Boolean
        Dim colon As System.Int32 = value.IndexOf(":"c)
        Return colon > 1 AndAlso Not value.StartsWith("file:", System.StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function ImportedRelativePath(sourcePath As System.String, sourceRoot As System.String) As System.String
        If System.String.IsNullOrEmpty(sourcePath) Then Return ""
        If Not System.String.IsNullOrEmpty(sourceRoot) Then
            Dim prefix As System.String = sourceRoot.TrimEnd("\"c, "/"c) & System.IO.Path.DirectorySeparatorChar
            If sourcePath.StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase) Then Return sourcePath.Substring(prefix.Length)
        End If
        Return System.IO.Path.GetFileName(sourcePath)
    End Function

    Private Sub EnsureWritable()
        If _workbook.ReadOnly Then Throw New System.InvalidOperationException("Save a writable copy of this workbook before editing its Tabular Review.")
        If _workbook.ProtectStructure Then Throw New System.InvalidOperationException("Unprotect the workbook structure before editing its Tabular Review.")
        Dim format As System.Int32 = CInt(_workbook.FileFormat)
        If format <> 51 AndAlso format <> 52 AndAlso format <> 50 Then
            Throw New System.InvalidOperationException("Save this workbook as .xlsx, .xlsm or .xlsb before creating a Tabular Review. CSV and older formats cannot retain the complete review database.")
        End If
    End Sub

    Private Shared Sub ValidateProjectionSize(documentCount As System.Int32, columnCount As System.Int32)
        If documentCount > MaximumRows - 4 OrElse columnCount > MaximumColumns Then
            Throw New System.InvalidOperationException("This review exceeds Excel's worksheet capacity. Split it into additional workbooks; Tabular Review will not discard documents or criteria.")
        End If
    End Sub

    Private Shared Function BuildColumns(project As Global.SharedLibrary.TabularReview.ReviewProject) As System.Collections.Generic.List(Of ProjectionColumn)
        Dim columns As New System.Collections.Generic.List(Of ProjectionColumn) From {
            New ProjectionColumn With {.Id = "d:id", .Title = "Review document ID", .Instruction = "Include this hidden ID column when sorting document rows.", .Field = "id"},
            New ProjectionColumn With {.Id = "d:number", .Title = "Document no.", .Instruction = "Document number derived from the source path; edit as required.", .Field = "number"},
            New ProjectionColumn With {.Id = "d:file", .Title = "Document", .Instruction = "Source filename. Use the review to load or relocate documents.", .Field = "file"},
            New ProjectionColumn With {.Id = "d:path", .Title = "Source file", .Instruction = "Original local or UNC path; the source document remains unchanged.", .Field = "path"},
            New ProjectionColumn With {.Id = "d:status", .Title = "Review status", .Instruction = "unreviewed, in_review or reviewed.", .Field = "status"},
            New ProjectionColumn With {.Id = "d:priority", .Title = "Priority", .Instruction = "User-defined review priority.", .Field = "priority"},
            New ProjectionColumn With {.Id = "d:flags", .Title = "Flags", .Instruction = "User-defined flags, one flag per line.", .Field = "flags"},
            New ProjectionColumn With {.Id = "d:reviewer", .Title = "Reviewer", .Instruction = "Person responsible for the review.", .Field = "reviewer"},
            New ProjectionColumn With {.Id = "d:notes", .Title = "Review notes", .Instruction = "Reviewer notes and follow-up actions.", .Field = "notes"},
            New ProjectionColumn With {.Id = "d:extraction", .Title = "Source status", .Instruction = "Extraction/availability status and any incomplete-source notice.", .Field = "extraction"},
            New ProjectionColumn With {.Id = "d:import", .Title = "Imported from", .Instruction = "Provenance of an existing Excel result row, when applicable.", .Field = "import"}
        }
        If project.SchemaVersion >= 4 Then
            columns.Add(New ProjectionColumn With {.Id = "d:category", .Title = "Category", .Instruction = "The category selects the questions for this document.", .Field = "category"})
            columns.Add(New ProjectionColumn With {.Id = "d:group", .Title = "Group", .Instruction = "Assign the same group name to documents to review together.", .Field = "group"})
        End If
        For Each criterion As Global.SharedLibrary.TabularReview.ReviewCriterion In project.Criteria
            If criterion.Archived Then Continue For
            columns.Add(New ProjectionColumn With {.Id = "c:" & criterion.Id & ":answer", .Title = criterion.Title, .Instruction = criterion.Instruction, .Category = criterion.Category, .CriterionId = criterion.Id, .Field = "answer"})
            If criterion.Kind = "analysis" Then
                If criterion.FlagValues().Count > 0 Then columns.Add(New ProjectionColumn With {.Id = "c:" & criterion.Id & ":riskFlag", .Title = criterion.Title & " — Flag", .Instruction = System.String.Join(System.Environment.NewLine, criterion.FlagValues()), .Category = criterion.Category, .CriterionId = criterion.Id, .Field = "riskFlag"})
                columns.Add(New ProjectionColumn With {.Id = "c:" & criterion.Id & ":assessment", .Title = criterion.Title & " — Assessment", .Instruction = "Assessment for: " & criterion.Title, .Category = criterion.Category, .CriterionId = criterion.Id, .Field = "assessment"})
                columns.Add(New ProjectionColumn With {.Id = "c:" & criterion.Id & ":sources", .Title = criterion.Title & " — Sources / status", .Instruction = "Answer status and host-validated source references. Edit answers or regenerate them in Tabular Review.", .Category = criterion.Category, .CriterionId = criterion.Id, .Field = "sources"})
            End If
        Next
        Return columns
    End Function

    Private Shared Function ActiveDocuments(project As Global.SharedLibrary.TabularReview.ReviewProject) As System.Collections.Generic.List(Of Global.SharedLibrary.TabularReview.ReviewDocument)
        Dim result As New System.Collections.Generic.List(Of Global.SharedLibrary.TabularReview.ReviewDocument)()
        For Each document As Global.SharedLibrary.TabularReview.ReviewDocument In project.Documents
            If Not document.Archived Then result.Add(document)
        Next
        Dim summaries = Global.SharedLibrary.TabularReview.ReviewContract.GroupSummaryRows(project)
        If project.SchemaVersion >= 5 Then
            Dim ordered As New System.Collections.Generic.List(Of Global.SharedLibrary.TabularReview.ReviewDocument)()
            Dim emitted As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each document As Global.SharedLibrary.TabularReview.ReviewDocument In result
                If System.String.IsNullOrEmpty(document.GroupName) Then
                    ordered.Add(document)
                ElseIf emitted.Add(document.GroupName) Then
                    ordered.Add(summaries.Find(Function(row As Global.SharedLibrary.TabularReview.ReviewDocument) row.GroupName = document.GroupName))
                    ordered.AddRange(result.FindAll(Function(item As Global.SharedLibrary.TabularReview.ReviewDocument) item.GroupName = document.GroupName))
                End If
            Next
            Return ordered
        End If
        result.AddRange(summaries)
        Return result
    End Function

    Private Shared Function ProjectionValue(project As Global.SharedLibrary.TabularReview.ReviewProject,
                                            columns As System.Collections.Generic.List(Of ProjectionColumn),
                                            documents As System.Collections.Generic.List(Of Global.SharedLibrary.TabularReview.ReviewDocument),
                                            row As System.Int32, column As System.Int32) As System.String
        If row = 1 Then
            If column = 1 Then Return ProjectionMarker & project.Id
            If column = 2 Then Return If(project.Title, "")
            If column = 3 Then Return ManagedWorksheetNotice
            Return ""
        End If
        Dim definition As ProjectionColumn = columns(column - 1)
        If row = 2 Then Return definition.Title
        If row = 3 Then Return If(definition.Instruction, "")
        If row = 4 Then Return definition.Id
        Dim document As Global.SharedLibrary.TabularReview.ReviewDocument = documents(row - 5)
        Return DocumentValue(definition, document)
    End Function

    Private Shared Function DocumentValue(definition As ProjectionColumn, document As Global.SharedLibrary.TabularReview.ReviewDocument) As System.String
        If definition.CriterionId.Length > 0 Then
            Dim groupRow As System.Boolean = document.Id.StartsWith("group:", System.StringComparison.Ordinal)
            If groupRow AndAlso Not document.Answers.ContainsKey(definition.CriterionId) Then Return ""
            If Not groupRow AndAlso Not System.String.IsNullOrEmpty(definition.Category) AndAlso Not System.String.Equals(definition.Category, document.Category, System.StringComparison.OrdinalIgnoreCase) Then Return "N/A"
            Dim answer As Global.SharedLibrary.TabularReview.ReviewAnswer = Nothing
            If Not document.Answers.TryGetValue(definition.CriterionId, answer) Then
                Return If(definition.Field = "sources", "Status: pending", "")
            End If
            Select Case definition.Field
                Case "answer" : Return answer.Text
                Case "riskFlag" : Return answer.RiskFlag
                Case "assessment" : Return If(answer.Assessment, "")
                Case "sources" : Return SourceDisplay(answer)
            End Select
        End If
        Select Case definition.Field
            Case "id" : Return document.Id
            Case "category" : Return document.Category
            Case "group" : Return document.GroupName
            Case "number" : Return If(document.Number, "")
            Case "file" : Return If(document.FileName, "")
            Case "path" : Return If(document.SourcePath, "")
            Case "status" : Return If(document.ReviewStatus, "")
            Case "priority" : Return If(document.Priority, "")
            Case "flags" : Return System.String.Join(System.Environment.NewLine, document.Flags)
            Case "reviewer" : Return If(document.Reviewer, "")
            Case "notes" : Return If(document.Notes, "")
            Case "extraction" : Return document.ExtractionStatus & If(System.String.IsNullOrEmpty(document.ExtractionMessage), "", System.Environment.NewLine & document.ExtractionMessage)
            Case "import" : Return If(document.ImportedFrom, "")
        End Select
        Throw New System.IO.InvalidDataException("Unknown review projection column.")
    End Function

    Private Shared Function SourceDisplay(answer As Global.SharedLibrary.TabularReview.ReviewAnswer) As System.String
        Dim builder As New System.Text.StringBuilder("Status: " & answer.Status)
        If Not System.String.IsNullOrWhiteSpace(answer.Error) Then builder.AppendLine().Append("Notice: ").Append(answer.Error)
        If answer.Evidence.Count = 0 Then builder.AppendLine().Append("No verified source reference.")
        For Each evidence As Global.SharedLibrary.TabularReview.ReviewEvidence In answer.Evidence
            builder.AppendLine().AppendLine().Append("[").Append(evidence.MatchStatus).Append("] ").Append(evidence.Locator)
            builder.AppendLine().Append(evidence.Quote)
        Next
        Return builder.ToString()
    End Function

    Private Shared Function VisibleText(value As System.String) As System.String
        Dim text As System.String = If(value, "")
        If text.Length <= MaximumCellLength Then Return text
        Dim suffix As System.String = System.Environment.NewLine & System.Environment.NewLine & OverflowNotice
        Dim take As System.Int32 = MaximumCellLength - suffix.Length
        If take > 0 AndAlso System.Char.IsHighSurrogate(text(take - 1)) Then take -= 1
        Return text.Substring(0, take) & suffix
    End Function

    Private Shared Function EditableText(previous As System.String, visible As System.String, caption As System.String) As System.String
        If System.String.Equals(VisibleText(previous), visible, System.StringComparison.Ordinal) Then Return If(previous, "")
        If visible.IndexOf(OverflowNotice, System.StringComparison.Ordinal) >= 0 Then
            Throw New System.InvalidOperationException("Edit the complete " & caption & " in Tabular Review. Excel displays only the beginning of this long field; its full text remains stored.")
        End If
        Return visible
    End Function

    Private Shared Sub WriteProjection(sheet As Microsoft.Office.Interop.Excel.Worksheet,
                                        project As Global.SharedLibrary.TabularReview.ReviewProject,
                                        columns As System.Collections.Generic.List(Of ProjectionColumn),
                                        documents As System.Collections.Generic.List(Of Global.SharedLibrary.TabularReview.ReviewDocument),
                                        previousRows As System.Int32, previousColumns As System.Int32)
        Dim rows As System.Int32 = documents.Count + 4
        ClearRectangle(sheet, System.Math.Max(rows, previousRows), System.Math.Max(columns.Count, previousColumns))
        For firstRow As System.Int32 = 1 To rows Step BlockRows
            Dim count As System.Int32 = System.Math.Min(BlockRows, rows - firstRow + 1)
            Dim data(count - 1, columns.Count - 1) As System.Object
            Dim expected(count - 1, columns.Count - 1) As System.String
            For row As System.Int32 = 0 To count - 1
                For column As System.Int32 = 0 To columns.Count - 1
                    expected(row, column) = VisibleText(ProjectionValue(project, columns, documents, firstRow + row, column + 1))
                    data(row, column) = LiteralValue(expected(row, column))
                Next
            Next
            Dim target As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, firstRow, 1, firstRow + count - 1, columns.Count)
            Try
                target.NumberFormat = "@"
                target.Value2 = data
                Dim actual As System.Object = target.Value2
                For row As System.Int32 = 0 To count - 1
                    For column As System.Int32 = 0 To columns.Count - 1
                        If ScalarText(ArrayItem(actual, row, column)) <> expected(row, column) Then
                            Throw New System.IO.InvalidDataException("Excel did not preserve a review cell's complete text. Tabular Review did not commit this save.")
                        End If
                    Next
                Next
            Finally
                ReleaseTemporary(target)
            End Try
        Next
    End Sub

    Private Shared Sub FormatProjection(sheet As Microsoft.Office.Interop.Excel.Worksheet, rows As System.Int32, columns As System.Int32, previousColumns As System.Int32)
        Dim title As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, 1, 2, 1, columns)
        Dim headings As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, 2, 2, 2, columns)
        Dim instructions As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, 3, 2, 3, columns)
        Dim identityRow As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, 4, 1, 4, columns)
        Dim identityColumn As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, 1, 1, rows, 1)
        Dim body As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, 1, 2, rows, columns)
        Try
            body.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            body.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop
            identityColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            identityColumn.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop
            If previousColumns = 0 Then
                Dim titleFont As Microsoft.Office.Interop.Excel.Font = title.Font
                Dim headingFont As Microsoft.Office.Interop.Excel.Font = headings.Font
                Dim headingFill As Microsoft.Office.Interop.Excel.Interior = headings.Interior
                Dim instructionFill As Microsoft.Office.Interop.Excel.Interior = instructions.Interior
                Try
                    titleFont.Bold = True
                    titleFont.Size = 16
                    headingFont.Bold = True
                    headingFont.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
                    headingFill.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(150, 30, 45))
                    instructionFill.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(247, 240, 241))
                    title.RowHeight = 27
                    headings.RowHeight = 31
                    instructions.RowHeight = 75
                    body.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop
                    body.WrapText = True
                Finally
                    ReleaseTemporary(titleFont)
                    ReleaseTemporary(headingFont)
                    ReleaseTemporary(headingFill)
                    ReleaseTemporary(instructionFill)
                End Try
            End If
            If columns > previousColumns Then
                Dim newColumns As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, 1, System.Math.Max(2, previousColumns + 1), rows, columns)
                Dim entire As Microsoft.Office.Interop.Excel.Range = newColumns.EntireColumn
                Try
                    entire.ColumnWidth = 38
                    newColumns.WrapText = True
                    newColumns.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop
                Finally
                    ReleaseTemporary(entire)
                    ReleaseTemporary(newColumns)
                End Try
            End If
            Dim notice As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, 1, 3, 1, 3)
            Dim noticeFont As Microsoft.Office.Interop.Excel.Font = notice.Font
            Try
                notice.WrapText = True
                noticeFont.Size = 10
                noticeFont.Bold = False
                noticeFont.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(150, 30, 45))
                Dim noticeRow As Microsoft.Office.Interop.Excel.Range = notice.EntireRow
                Try
                    noticeRow.AutoFit()
                Finally
                    ReleaseTemporary(noticeRow)
                End Try
            Finally
                ReleaseTemporary(noticeFont)
                ReleaseTemporary(notice)
            End Try
            Dim rowRange As Microsoft.Office.Interop.Excel.Range = identityRow.EntireRow
            Dim columnRange As Microsoft.Office.Interop.Excel.Range = identityColumn.EntireColumn
            Try
                rowRange.Hidden = True
                columnRange.Hidden = True
            Finally
                ReleaseTemporary(rowRange)
                ReleaseTemporary(columnRange)
            End Try
        Finally
            ReleaseTemporary(title)
            ReleaseTemporary(headings)
            ReleaseTemporary(instructions)
            ReleaseTemporary(identityRow)
            ReleaseTemporary(identityColumn)
            ReleaseTemporary(body)
        End Try
    End Sub

    Private Shared Sub ApplyQuestionFlagColors(sheet As Microsoft.Office.Interop.Excel.Worksheet,
            project As Global.SharedLibrary.TabularReview.ReviewProject, columns As System.Collections.Generic.List(Of ProjectionColumn),
            documents As System.Collections.Generic.List(Of Global.SharedLibrary.TabularReview.ReviewDocument), rows As System.Int32, width As System.Int32)
        If rows < 5 Then Return
        Const marker As System.String = "=N(""RedInk.TabularReview.Flag"")=0"
        Dim area As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, 5, 1, rows, width)
        Dim conditions As Microsoft.Office.Interop.Excel.FormatConditions = Nothing
        Try
            conditions = area.FormatConditions
            For index As System.Int32 = conditions.Count To 1 Step -1
                Dim existing As System.Object = conditions.Item(index)
                Try
                    Dim rule As Microsoft.Office.Interop.Excel.FormatCondition = TryCast(existing, Microsoft.Office.Interop.Excel.FormatCondition)
                    If rule IsNot Nothing AndAlso rule.Type = Microsoft.Office.Interop.Excel.XlFormatConditionType.xlExpression AndAlso
                       System.String.Equals(System.Convert.ToString(rule.Formula1), marker, System.StringComparison.OrdinalIgnoreCase) Then rule.Delete()
                Finally
                    ReleaseTemporary(existing)
                End Try
            Next
        Finally
            ReleaseTemporary(conditions)
            ReleaseTemporary(area)
        End Try
        Dim fileColumn As System.Int32 = columns.FindIndex(Function(item As ProjectionColumn) item.Field = "file") + 1
        If fileColumn > 0 AndAlso documents.Count > 0 AndAlso project.SchemaVersion >= 5 Then
            Dim startRow As System.Int32 = 0
            While startRow < documents.Count
                Dim child As System.Boolean = IsGroupMember(documents(startRow))
                Dim endRow As System.Int32 = startRow + 1
                While endRow < documents.Count AndAlso IsGroupMember(documents(endRow)) = child
                    endRow += 1
                End While
                Dim labels As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, startRow + 5, fileColumn, endRow + 4, fileColumn)
                Try
                    labels.IndentLevel = If(child, 1, 0)
                Finally
                    ReleaseTemporary(labels)
                End Try
                startRow = endRow
            End While
        End If
        For column As System.Int32 = 0 To columns.Count - 1
            Dim definition As ProjectionColumn = columns(column)
            If definition.Field <> "riskFlag" Then Continue For
            Dim start As System.Int32 = 0
            While start < documents.Count
                Dim label As System.String = VisibleFlag(definition, documents(start))
                Dim finish As System.Int32 = start + 1
                While finish < documents.Count AndAlso VisibleFlag(definition, documents(finish)) = label
                    finish += 1
                End While
                Dim color As System.String = Nothing
                If label.Length > 0 AndAlso project.FlagColors.TryGetValue(label, color) Then
                    Dim target As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, start + 5, column + 1, finish + 4, column + 1)
                    Dim rules As Microsoft.Office.Interop.Excel.FormatConditions = Nothing
                    Dim rule As Microsoft.Office.Interop.Excel.FormatCondition = Nothing
                    Dim fill As Microsoft.Office.Interop.Excel.Interior = Nothing
                    Try
                        rules = target.FormatConditions
                        rule = DirectCast(rules.Add(Microsoft.Office.Interop.Excel.XlFormatConditionType.xlExpression, Formula1:=marker), Microsoft.Office.Interop.Excel.FormatCondition)
                        fill = rule.Interior
                        fill.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml(color))
                        rule.StopIfTrue = False
                    Finally
                        ReleaseTemporary(fill)
                        ReleaseTemporary(rule)
                        ReleaseTemporary(rules)
                        ReleaseTemporary(target)
                    End Try
                End If
                start = finish
            End While
        Next
    End Sub

    Private Shared Function IsGroupMember(document As Global.SharedLibrary.TabularReview.ReviewDocument) As System.Boolean
        Return Not System.String.IsNullOrEmpty(document.GroupName) AndAlso Not document.Id.StartsWith("group:", System.StringComparison.Ordinal)
    End Function

    Private Shared Function VisibleFlag(definition As ProjectionColumn, document As Global.SharedLibrary.TabularReview.ReviewDocument) As System.String
        If Not document.Id.StartsWith("group:", System.StringComparison.Ordinal) AndAlso
           Not System.String.IsNullOrWhiteSpace(definition.Category) AndAlso Not System.String.Equals(definition.Category, document.Category, System.StringComparison.OrdinalIgnoreCase) Then Return ""
        Dim answer As Global.SharedLibrary.TabularReview.ReviewAnswer = Nothing
        Return If(document.Answers.TryGetValue(definition.CriterionId, answer), If(answer.RiskFlag, ""), "")
    End Function

    Private Shared Function LiteralValue(text As System.String) As System.Object
        If System.String.IsNullOrEmpty(text) Then Return Nothing
        ' Prefix every string, including a literal leading apostrophe. Excel consumes
        ' one quote prefix; readback above verifies that no text or formula escaped.
        Return "'" & text
    End Function

    Private Shared Sub ClearRectangle(sheet As Microsoft.Office.Interop.Excel.Worksheet, rows As System.Int32, columns As System.Int32)
        If rows <= 0 OrElse columns <= 0 Then Return
        Dim target As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, 1, 1, rows, columns)
        Try
            target.ClearContents()
        Finally
            ReleaseTemporary(target)
        End Try
    End Sub

    Private Shared Function CaptureProjection(sheet As Microsoft.Office.Interop.Excel.Worksheet, rows As System.Int32, columns As System.Int32) As System.Collections.Generic.List(Of ProjectionBackupBlock)
        Dim result As New System.Collections.Generic.List(Of ProjectionBackupBlock)()
        For firstRow As System.Int32 = 1 To rows Step BlockRows
            Dim count As System.Int32 = System.Math.Min(BlockRows, rows - firstRow + 1)
            Dim range As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, firstRow, 1, firstRow + count - 1, columns)
            Try
                Dim read As System.Object = range.Value2
                Dim values(count - 1, columns - 1) As System.Object
                For row As System.Int32 = 0 To count - 1
                    For column As System.Int32 = 0 To columns - 1
                        values(row, column) = ArrayItem(read, row, column)
                    Next
                Next
                result.Add(New ProjectionBackupBlock With {.FirstRow = firstRow, .Values = values})
            Finally
                ReleaseTemporary(range)
            End Try
        Next
        Return result
    End Function

    Private Shared Sub RestoreProjection(sheet As Microsoft.Office.Interop.Excel.Worksheet,
                                         backup As System.Collections.Generic.List(Of ProjectionBackupBlock),
                                         rows As System.Int32, columns As System.Int32, newRows As System.Int32, newColumns As System.Int32)
        ClearRectangle(sheet, System.Math.Max(rows, newRows), System.Math.Max(columns, newColumns))
        For Each block As ProjectionBackupBlock In backup
            Dim count As System.Int32 = block.Values.GetLength(0)
            Dim data(count - 1, columns - 1) As System.Object
            For row As System.Int32 = 0 To count - 1
                For column As System.Int32 = 0 To columns - 1
                    Dim value As System.Object = block.Values(row, column)
                    data(row, column) = If(TypeOf value Is System.String, LiteralValue(CStr(value)), value)
                Next
            Next
            Dim range As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, block.FirstRow, 1, block.FirstRow + count - 1, columns)
            Try
                range.NumberFormat = "@"
                range.Value2 = data
                Dim restored As System.Object = range.Value2
                For row As System.Int32 = 0 To count - 1
                    For column As System.Int32 = 0 To columns - 1
                        If ScalarText(ArrayItem(restored, row, column)) <> ScalarText(block.Values(row, column)) Then
                            Throw New System.IO.InvalidDataException("Excel could not restore a review cell after a failed save.")
                        End If
                    Next
                Next
            Finally
                ReleaseTemporary(range)
            End Try
        Next
    End Sub

    Private Sub WriteSnapshot(sheet As Microsoft.Office.Interop.Excel.Worksheet, slot As System.Int32, snapshot As StoredSnapshot)
        Dim json As System.String = Global.SharedLibrary.TabularReview.ReviewContract.Serialize(snapshot)
        Dim bytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(json)
        If bytes.LongLength > MaximumSnapshotBytes Then Throw New System.InvalidOperationException("The complete review state exceeds the supported snapshot size. Split the review into workbooks; no text has been discarded.")
        Dim payload As System.String
        Using output As New System.IO.MemoryStream()
            Using compressor As New System.IO.Compression.GZipStream(output, System.IO.Compression.CompressionLevel.Optimal, True)
                compressor.Write(bytes, 0, bytes.Length)
            End Using
            payload = System.Convert.ToBase64String(output.ToArray())
        End Using
        Dim chunks As System.Int32 = CInt((CLng(payload.Length) + PayloadChunkLength - 1L) \ PayloadChunkLength)
        If chunks <= 0 OrElse chunks > MaximumRows - PayloadFirstRow + 1 Then Throw New System.InvalidOperationException("The complete review snapshot exceeds Excel's capacity.")
        Dim column As System.Int32 = slot + 3
        PutText(sheet, 5, column, "writing")
        For first As System.Int32 = 0 To chunks - 1 Step BlockRows
            Dim count As System.Int32 = System.Math.Min(BlockRows, chunks - first)
            Dim data(count - 1, 0) As System.Object
            For index As System.Int32 = 0 To count - 1
                Dim offset As System.Int32 = (first + index) * PayloadChunkLength
                data(index, 0) = payload.Substring(offset, System.Math.Min(PayloadChunkLength, payload.Length - offset))
            Next
            Dim target As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, PayloadFirstRow + first, column, PayloadFirstRow + first + count - 1, column)
            Try
                target.NumberFormat = "@"
                ' Excel interop can return a scalar for a one-cell Range. Write a
                ' scalar in that case; multi-cell transfers retain the block path.
                If count = 1 Then
                    target.Value2 = data(0, 0)
                Else
                    target.Value2 = data
                End If
                Dim readback As System.Object = target.Value2
                For index As System.Int32 = 0 To count - 1
                    If ScalarText(ArrayItem(readback, index, 0)) <> CStr(data(index, 0)) Then
                        Throw New System.IO.InvalidDataException("Excel did not retain a review payload chunk. The uncommitted snapshot was not activated.")
                    End If
                Next
            Finally
                ReleaseTemporary(target)
            End Try
        Next
        PutText(sheet, 1, column, snapshot.Project.Revision.ToString(System.Globalization.CultureInfo.InvariantCulture))
        PutText(sheet, 2, column, Global.SharedLibrary.TabularReview.ReviewContract.Digest(json))
        PutText(sheet, 3, column, chunks.ToString(System.Globalization.CultureInfo.InvariantCulture))
        PutText(sheet, 4, column, bytes.LongLength.ToString(System.Globalization.CultureInfo.InvariantCulture))
        ' Status remains "writing" until Save has read back and verified every field.
    End Sub

    Private Shared Function ReadSnapshot(sheet As Microsoft.Office.Interop.Excel.Worksheet, slot As System.Int32,
                                         Optional allowPreparing As System.Boolean = False) As StoredSnapshot
        Dim column As System.Int32 = slot + 3
        Dim state As System.String = CellText(sheet, 5, column)
        If state <> "complete" AndAlso Not (allowPreparing AndAlso state = "writing") Then Throw New System.IO.InvalidDataException("The review snapshot is incomplete.")
        Dim chunks As System.Int32
        Dim byteCount As System.Int64
        Dim revision As System.Int64
        If Not System.Int32.TryParse(CellText(sheet, 3, column), chunks) OrElse chunks < 1 OrElse chunks > MaximumRows - PayloadFirstRow + 1 Then Throw New System.IO.InvalidDataException("Invalid review chunk count.")
        If Not System.Int64.TryParse(CellText(sheet, 4, column), byteCount) OrElse byteCount < 1 OrElse byteCount > MaximumSnapshotBytes Then Throw New System.IO.InvalidDataException("Invalid review snapshot size.")
        If Not System.Int64.TryParse(CellText(sheet, 1, column), revision) OrElse revision < 1 Then Throw New System.IO.InvalidDataException("Invalid review revision.")
        If CLng(chunks) * PayloadChunkLength > MaximumSnapshotBytes * 2L Then Throw New System.IO.InvalidDataException("The encoded review snapshot exceeds its safe read limit.")
        Dim payload As New System.Text.StringBuilder()
        For first As System.Int32 = 0 To chunks - 1 Step BlockRows
            Dim count As System.Int32 = System.Math.Min(BlockRows, chunks - first)
            Dim target As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, PayloadFirstRow + first, column, PayloadFirstRow + first + count - 1, column)
            Try
                Dim values As System.Object = target.Value2
                For index As System.Int32 = 0 To count - 1
                    Dim chunk As System.String = ScalarText(ArrayItem(values, index, 0))
                    ' Older builds prefixed base64 with an Excel apostrophe. If a
                    ' text-formatted cell retained it literally, allow that legacy
                    ' representation; the SHA-256 and data validation still apply.
                    If chunk.StartsWith("'", System.StringComparison.Ordinal) Then chunk = chunk.Substring(1)
                    If chunk.Length < 1 OrElse chunk.Length > PayloadChunkLength Then Throw New System.IO.InvalidDataException("A review snapshot chunk is missing or invalid.")
                    payload.Append(chunk)
                Next
            Finally
                ReleaseTemporary(target)
            End Try
        Next
        Dim json As System.String
        Using source As New System.IO.MemoryStream(System.Convert.FromBase64String(payload.ToString()))
            Using decompressor As New System.IO.Compression.GZipStream(source, System.IO.Compression.CompressionMode.Decompress)
                Using output As New System.IO.MemoryStream()
                    Dim buffer(8191) As System.Byte
                    Do
                        Dim read As System.Int32 = decompressor.Read(buffer, 0, buffer.Length)
                        If read = 0 Then Exit Do
                        If output.Length + read > byteCount Then Throw New System.IO.InvalidDataException("The review snapshot expanded beyond its declared size.")
                        output.Write(buffer, 0, read)
                    Loop
                    If output.Length <> byteCount Then Throw New System.IO.InvalidDataException("The review snapshot is truncated.")
                    json = New System.Text.UTF8Encoding(False, True).GetString(output.ToArray())
                End Using
            End Using
        End Using
        If Not System.String.Equals(Global.SharedLibrary.TabularReview.ReviewContract.Digest(json), CellText(sheet, 2, column), System.StringComparison.Ordinal) Then Throw New System.IO.InvalidDataException("The review snapshot checksum does not match.")
        Dim snapshot As StoredSnapshot = Newtonsoft.Json.JsonConvert.DeserializeObject(Of StoredSnapshot)(json, Global.SharedLibrary.TabularReview.ReviewContract.JsonSettings)
        If snapshot Is Nothing OrElse snapshot.FormatVersion <> 1 OrElse snapshot.Project Is Nothing Then Throw New System.IO.InvalidDataException("Unsupported review snapshot format.")
        Global.SharedLibrary.TabularReview.ReviewContract.Validate(snapshot.Project)
        If snapshot.Project.Revision <> revision OrElse snapshot.ProjectionRows < 4 OrElse snapshot.ProjectionRows > MaximumRows OrElse snapshot.ProjectionColumns < 11 OrElse snapshot.ProjectionColumns > MaximumColumns Then Throw New System.IO.InvalidDataException("Invalid review projection metadata.")
        If ActiveDocuments(snapshot.Project).Count + 4 <> snapshot.ProjectionRows OrElse BuildColumns(snapshot.Project).Count <> snapshot.ProjectionColumns Then Throw New System.IO.InvalidDataException("The review snapshot's rows and columns do not match its complete project.")
        ' CodeName is not a guaranteed Excel contract (notably for newly added
        ' sheets); older snapshots may legitimately have no CodeName at all.
        ' Accept either persisted name or CodeName, but never accept a snapshot
        ' without an actual worksheet reference or projection checksum.
        If System.String.IsNullOrWhiteSpace(snapshot.WorksheetName) AndAlso System.String.IsNullOrWhiteSpace(snapshot.WorksheetCodeName) Then
            Throw New System.IO.InvalidDataException("The review snapshot has no worksheet name or code name.")
        End If
        If System.String.IsNullOrWhiteSpace(snapshot.ProjectionFingerprint) Then
            Throw New System.IO.InvalidDataException("The review snapshot has no worksheet fingerprint.")
        End If
        Return snapshot
    End Function

    Private Shared Function ComputeFingerprint(sheet As Microsoft.Office.Interop.Excel.Worksheet, rows As System.Int32, columns As System.Int32) As System.String
        Using hash As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
            HashText(hash, rows.ToString(System.Globalization.CultureInfo.InvariantCulture) & ":" & columns.ToString(System.Globalization.CultureInfo.InvariantCulture) & ";")
            For firstRow As System.Int32 = 1 To rows Step BlockRows
                Dim count As System.Int32 = System.Math.Min(BlockRows, rows - firstRow + 1)
                Dim range As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, firstRow, 1, firstRow + count - 1, columns)
                Try
                    Dim values As System.Object = range.Value2
                    Dim formulas As System.Object = range.Formula
                    Dim builder As New System.Text.StringBuilder()
                    For row As System.Int32 = 0 To count - 1
                        For column As System.Int32 = 0 To columns - 1
                            AppendScalar(builder, ArrayItem(values, row, column))
                            AppendScalar(builder, ArrayItem(formulas, row, column))
                        Next
                    Next
                    HashText(hash, builder.ToString())
                Finally
                    ReleaseTemporary(range)
                End Try
            Next
            hash.TransformFinalBlock(New System.Byte() {}, 0, 0)
            Return System.BitConverter.ToString(hash.Hash).Replace("-", "").ToLowerInvariant()
        End Using
    End Function

    Private Shared Sub HashText(hash As System.Security.Cryptography.HashAlgorithm, value As System.String)
        Dim bytes As System.Byte() = System.Text.Encoding.UTF8.GetBytes(value)
        hash.TransformBlock(bytes, 0, bytes.Length, bytes, 0)
    End Sub

    Private Shared Sub AppendScalar(builder As System.Text.StringBuilder, value As System.Object)
        If value Is Nothing Then
            builder.Append("n;")
            Return
        End If
        Dim text As System.String = ScalarText(value)
        builder.Append(value.GetType().FullName).Append(":").Append(text.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append(":").Append(text).Append(";")
    End Sub

    Private Shared Function ProjectionMatches(sheet As Microsoft.Office.Interop.Excel.Worksheet, rows As System.Int32, columns As System.Int32, fingerprint As System.String) As System.Boolean
        Dim extent As System.Tuple(Of System.Int32, System.Int32) = ContentExtent(sheet)
        Return extent.Item1 = rows AndAlso extent.Item2 = columns AndAlso System.String.Equals(ComputeFingerprint(sheet, rows, columns), fingerprint, System.StringComparison.Ordinal)
    End Function

    Private Shared Function ContentExtent(sheet As Microsoft.Office.Interop.Excel.Worksheet) As System.Tuple(Of System.Int32, System.Int32)
        Dim cells As Microsoft.Office.Interop.Excel.Range = CType(sheet.Cells, Microsoft.Office.Interop.Excel.Range)
        Dim first As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, 1, 1, 1, 1)
        Dim lastRow As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim lastColumn As Microsoft.Office.Interop.Excel.Range = Nothing
        Try
            lastRow = cells.Find(What:="*", After:=first, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, MatchCase:=False, SearchFormat:=False)
            lastColumn = cells.Find(What:="*", After:=first, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, MatchCase:=False, SearchFormat:=False)
            Return System.Tuple.Create(If(lastRow Is Nothing, 0, lastRow.Row), If(lastColumn Is Nothing, 0, lastColumn.Column))
        Finally
            ReleaseTemporary(lastRow)
            ReleaseTemporary(lastColumn)
            ReleaseTemporary(first)
            ReleaseTemporary(cells)
        End Try
    End Function

    Private Function FindStoreSheet() As Microsoft.Office.Interop.Excel.Worksheet
        Dim sheets As Microsoft.Office.Interop.Excel.Sheets = _workbook.Worksheets
        Dim found As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Try
            For index As System.Int32 = 1 To sheets.Count
                Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = CType(sheets.Item(index), Microsoft.Office.Interop.Excel.Worksheet)
                If CellText(sheet, 1, 1) = StoreMarker Then
                    If found IsNot Nothing Then
                        ReleaseTemporary(sheet)
                        Throw New System.IO.InvalidDataException("This workbook contains multiple review databases. Restore one unambiguous workbook copy before opening Tabular Review.")
                    End If
                    found = sheet
                Else
                    ReleaseTemporary(sheet)
                End If
            Next
            Return found
        Catch
            ReleaseTemporary(found)
            Throw
        Finally
            ReleaseTemporary(sheets)
        End Try
    End Function

    Private Function FindProjectionSheet(snapshot As StoredSnapshot) As Microsoft.Office.Interop.Excel.Worksheet
        If snapshot Is Nothing OrElse snapshot.Project Is Nothing Then Throw New System.ArgumentNullException(NameOf(snapshot))
        Dim sheets As Microsoft.Office.Interop.Excel.Sheets = _workbook.Worksheets
        Dim matches As New System.Collections.Generic.List(Of Microsoft.Office.Interop.Excel.Worksheet)()
        Dim selected As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Try
            Dim expectedMarker As System.String = ProjectionMarker & snapshot.Project.Id
            For index As System.Int32 = 1 To sheets.Count
                Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = CType(sheets.Item(index), Microsoft.Office.Interop.Excel.Worksheet)
                Try
                    If System.String.Equals(CellText(sheet, 1, 1), expectedMarker, System.StringComparison.Ordinal) Then
                        matches.Add(sheet)
                        sheet = Nothing
                    End If
                Finally
                    ReleaseTemporary(sheet)
                End Try
            Next

            ' Prefer the immutable CodeName when both Excel and the snapshot
            ' expose it, then the persisted sheet name. Both are constrained by
            ' the exact project-ID marker in A1, not by an ActiveSheet fallback.
            If Not System.String.IsNullOrWhiteSpace(snapshot.WorksheetCodeName) Then
                For Each candidate As Microsoft.Office.Interop.Excel.Worksheet In matches
                    If System.String.Equals(SafeWorksheetCodeName(candidate), snapshot.WorksheetCodeName, System.StringComparison.Ordinal) Then
                        selected = candidate
                        Exit For
                    End If
                Next
            End If
            If selected Is Nothing AndAlso Not System.String.IsNullOrWhiteSpace(snapshot.WorksheetName) Then
                For Each candidate As Microsoft.Office.Interop.Excel.Worksheet In matches
                    If System.String.Equals(candidate.Name, snapshot.WorksheetName, System.StringComparison.OrdinalIgnoreCase) Then
                        selected = candidate
                        Exit For
                    End If
                Next
            End If
            ' A user may have renamed the sheet. Recovery remains safe only if
            ' exactly one worksheet carries this project's unique identity.
            If selected Is Nothing AndAlso matches.Count = 1 Then selected = matches(0)
            If selected Is Nothing AndAlso matches.Count > 1 Then
                Throw New System.IO.InvalidDataException("Multiple worksheets identify the same review. Restore an unambiguous review worksheet before continuing.")
            End If
            Return selected
        Finally
            For Each candidate As Microsoft.Office.Interop.Excel.Worksheet In matches
                If Not System.Object.ReferenceEquals(candidate, selected) Then ReleaseTemporary(candidate)
            Next
            ReleaseTemporary(sheets)
        End Try
    End Function

    Private Shared Function SafeWorksheetCodeName(sheet As Microsoft.Office.Interop.Excel.Worksheet) As System.String
        Try
            Return If(sheet.CodeName, "")
        Catch ex As System.Runtime.InteropServices.COMException
            ' CodeName is optional metadata. The saved name, unique A1 marker,
            ' and checked cell fingerprint remain mandatory.
            System.Diagnostics.Trace.WriteLine("[TabularReview:codename_unavailable] " & ex.Message)
            Return ""
        End Try
    End Function

    Private Function AddWorksheet(baseName As System.String) As Microsoft.Office.Interop.Excel.Worksheet
        Dim sheets As Microsoft.Office.Interop.Excel.Sheets = _workbook.Worksheets
        Dim lastSheet As System.Object = Nothing
        Try
            Dim names As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For index As System.Int32 = 1 To sheets.Count
                Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = CType(sheets.Item(index), Microsoft.Office.Interop.Excel.Worksheet)
                names.Add(sheet.Name)
                ReleaseTemporary(sheet)
            Next
            Dim name As System.String = baseName
            Dim suffix As System.Int32 = 2
            While names.Contains(name)
                name = baseName & " " & suffix.ToString(System.Globalization.CultureInfo.InvariantCulture)
                suffix += 1
            End While
            If sheets.Count > 0 Then lastSheet = sheets.Item(sheets.Count)
            Dim added As Microsoft.Office.Interop.Excel.Worksheet = CType(sheets.Add(After:=lastSheet), Microsoft.Office.Interop.Excel.Worksheet)
            added.Name = name
            Return added
        Finally
            ReleaseTemporary(lastSheet)
            ReleaseTemporary(sheets)
        End Try
    End Function

    Private Shared Function Rectangle(sheet As Microsoft.Office.Interop.Excel.Worksheet, firstRow As System.Int32, firstColumn As System.Int32, lastRow As System.Int32, lastColumn As System.Int32) As Microsoft.Office.Interop.Excel.Range
        Dim cells As Microsoft.Office.Interop.Excel.Range = CType(sheet.Cells, Microsoft.Office.Interop.Excel.Range)
        Dim first As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim last As Microsoft.Office.Interop.Excel.Range = Nothing
        Try
            first = CType(cells.Item(firstRow, firstColumn), Microsoft.Office.Interop.Excel.Range)
            last = CType(cells.Item(lastRow, lastColumn), Microsoft.Office.Interop.Excel.Range)
            Return sheet.Range(first, last)
        Finally
            ReleaseTemporary(first)
            ReleaseTemporary(last)
            ReleaseTemporary(cells)
        End Try
    End Function

    Private Shared Sub PutText(sheet As Microsoft.Office.Interop.Excel.Worksheet, row As System.Int32, column As System.Int32, value As System.String)
        If value IsNot Nothing AndAlso value.Length > MaximumCellLength Then Throw New System.ArgumentException("A review metadata cell exceeds Excel's text capacity.", NameOf(value))
        Dim cell As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, row, column, row, column)
        Try
            cell.NumberFormat = "@"
            Dim expected As System.String = If(value, "")
            cell.Value2 = LiteralValue(expected)
            If ScalarText(cell.Value2) <> expected Then
                ' Excel versions differ in how a leading quote is treated in
                ' text-formatted cells. Retry the literal string, then verify.
                cell.Value2 = expected
                System.Diagnostics.Trace.WriteLine("[TabularReview:metadata_literal_retry] " & row.ToString() & ":" & column.ToString())
            End If
            If CBool(cell.HasFormula) OrElse ScalarText(cell.Value2) <> expected Then Throw New System.IO.InvalidDataException("Excel did not retain a review metadata value.")
        Finally
            ReleaseTemporary(cell)
        End Try
    End Sub

    Private Shared Function CellText(sheet As Microsoft.Office.Interop.Excel.Worksheet, row As System.Int32, column As System.Int32) As System.String
        Dim cell As Microsoft.Office.Interop.Excel.Range = Rectangle(sheet, row, column, row, column)
        Try
            Return ScalarText(cell.Value2)
        Finally
            ReleaseTemporary(cell)
        End Try
    End Function

    Private Shared Function ArrayItem(value As System.Object, row As System.Int32, column As System.Int32) As System.Object
        Dim matrix As System.Array = TryCast(value, System.Array)
        If matrix IsNot Nothing AndAlso matrix.Rank = 2 Then Return matrix.GetValue(row + matrix.GetLowerBound(0), column + matrix.GetLowerBound(1))
        If row = 0 AndAlso column = 0 Then Return value
        Return Nothing
    End Function

    Private Shared Function ScalarText(value As System.Object) As System.String
        If value Is Nothing Then Return ""
        If TypeOf value Is System.Double Then Return CDbl(value).ToString("R", System.Globalization.CultureInfo.InvariantCulture)
        Return System.Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Shared Function SameComIdentity(left As System.Object, right As System.Object) As System.Boolean
        If System.Object.ReferenceEquals(left, right) Then Return True
        If left Is Nothing OrElse right Is Nothing Then Return False
        Dim leftPointer As System.IntPtr = System.IntPtr.Zero
        Dim rightPointer As System.IntPtr = System.IntPtr.Zero
        Try
            leftPointer = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(left)
            rightPointer = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(right)
            Return leftPointer = rightPointer
        Finally
            If leftPointer <> System.IntPtr.Zero Then System.Runtime.InteropServices.Marshal.Release(leftPointer)
            If rightPointer <> System.IntPtr.Zero Then System.Runtime.InteropServices.Marshal.Release(rightPointer)
        End Try
    End Function

    Private Shared Sub ReleaseTemporary(value As System.Object)
        If value Is Nothing OrElse Not System.Runtime.InteropServices.Marshal.IsComObject(value) Then Return
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(value)
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("Tabular Review temporary COM release: " & ex.Message)
        End Try
    End Sub
End Class
