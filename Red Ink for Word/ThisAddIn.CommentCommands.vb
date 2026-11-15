' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.CommentCommands.vb
' PURPOSE
'   Implements Word modern-comment resolution commands and guarded user-facing failure handling.
'
' KEY ARCHITECTURE
'   • Host: Word comment commands resolve selected comment evidence before execution.
'   • Threading: Word COM and custom dialogs remain on the Office STA.
' =============================================================================

'
' Command-only modern-comment resolution and user-facing failure boundary.
' Does not change TryResolveActiveComment or the general comment-editing policy.
' All Word COM and dialogs stay on Office's STA. Shared dialog helpers own windows.

Option Explicit On
Option Strict On

Partial Public Class ThisAddIn
    Private Enum WordEventSubscriptionStatus
        Inactive
        Active
        Unknown
    End Enum

    ' These states describe this add-in's own COM subscriptions, never other add-ins.
    ' Unknown is retained after a failed COM operation; it must not be reported as inactive.
    Private ReadOnly _wordEventSubscriptions As New System.Collections.Generic.Dictionary(Of System.String, WordEventSubscriptionStatus)(System.StringComparer.Ordinal)
    Private ReadOnly _wordEventSubscriptionErrors As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
    Private ReadOnly _wordEventChangesInProgress As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
    Private ReadOnly _deferredWordEventRemovals As New System.Collections.Generic.Dictionary(Of System.String, System.Action)(System.StringComparer.Ordinal)
    Private _wordSelectionModeInitialized As System.Boolean
    Private _wordSelectionShutdown As System.Boolean
    Private _wordSelectionShutdownInProgress As System.Boolean
    Private _commentSelectionApplication As Microsoft.Office.Interop.Word.Application
    Private _commentSelectionTracker As WordCommentSelectionTracker
    Private _retiredCommentSelectionTracker As WordCommentSelectionTracker
    ' Keep this as the explicit startup opt-out: tracker failure must retain the conservative fallback.
    Private _commentSelectionTrackingDisabled As System.Boolean
    Private _commentProbeTimer As System.Windows.Forms.Timer
    Private _commentProbePending As System.Boolean
    Private _commentProbeRunning As System.Boolean
    Private _commentContextDocument As Microsoft.Office.Interop.Word.Document
    Private _commentContextWindow As System.Int32
    Private _commentContextRevision As System.Int32
    Private _lastCommentBodyWindow As System.Int32
    Private _lastCommentBodyStart As System.Int32 = -1
    Private _lastCommentBodyEnd As System.Int32 = -1
    Private _lastCommentBodyStory As System.Int32 = -1

    Private ReadOnly Property WordSelectionTrackingEnabled As System.Boolean
        Get
            Return _wordSelectionModeInitialized AndAlso Not _commentSelectionTrackingDisabled AndAlso Not _wordSelectionShutdown
        End Get
    End Property

    Private Function GetWordEventSubscriptionStatus(eventName As System.String) As WordEventSubscriptionStatus
        Dim status As WordEventSubscriptionStatus
        If _wordEventSubscriptions.TryGetValue(eventName, status) Then Return status
        Return WordEventSubscriptionStatus.Inactive
    End Function

    ' All callers run on the Word UI thread. Each COM operation is isolated so one
    ' failed removal cannot prevent another handler from being detached.
    Private Function SetWordEventSubscription(eventName As System.String,
                                              enabled As System.Boolean,
                                              changeSubscription As System.Action) As System.Boolean
        Dim desired As WordEventSubscriptionStatus = If(enabled, WordEventSubscriptionStatus.Active, WordEventSubscriptionStatus.Inactive)
        Dim current As WordEventSubscriptionStatus = GetWordEventSubscriptionStatus(eventName)
        If current = desired Then Return True
        ' COM may reenter Word while AddHandler/RemoveHandler is in progress. Keep
        ' the source reference alive and finish a requested removal after that call.
        If _wordEventChangesInProgress.Contains(eventName) Then
            If Not enabled Then _deferredWordEventRemovals(eventName) = changeSubscription
            Return False
        End If
        ' An uncertain AddHandler may already have succeeded. Never add it twice.
        If enabled AndAlso current = WordEventSubscriptionStatus.Unknown Then Return False

        _wordEventChangesInProgress.Add(eventName)
        _wordEventSubscriptions(eventName) = WordEventSubscriptionStatus.Unknown
        Try
            changeSubscription.Invoke()
            _wordEventSubscriptions(eventName) = desired
            _wordEventSubscriptionErrors.Remove(eventName)
        Catch ex As System.Exception
            _wordEventSubscriptions(eventName) = WordEventSubscriptionStatus.Unknown
            _wordEventSubscriptionErrors(eventName) = ex.ToString()
            LogWordCommandDiagnostic("SelectionTracking.Subscription", eventName & "; requested=" & desired.ToString() & "; actual=Unknown; " & ex.ToString())
        Finally
            _wordEventChangesInProgress.Remove(eventName)
            Dim deferredRemoval As System.Action = Nothing
            If _deferredWordEventRemovals.TryGetValue(eventName, deferredRemoval) Then
                _deferredWordEventRemovals.Remove(eventName)
                SetWordEventSubscription(eventName, False, deferredRemoval)
            End If
        End Try
        Return GetWordEventSubscriptionStatus(eventName) = desired
    End Function

    ' Called only after the startup configuration has returned. The documented
    ' restart-required mode governs both selection handlers and is never re-latched.
    Private Sub StartWordSelectionTracking()
        If _wordSelectionShutdown Then
            LogWordSelectionTrackingStatus("StartIgnoredAfterShutdown")
            Return
        End If
        If Not _wordSelectionModeInitialized Then
            _commentSelectionTrackingDisabled = _context Is Nothing OrElse Not _context.INI_WordCommentSelectionTracking
            _wordSelectionModeInitialized = True
        End If
        If Not WordSelectionTrackingEnabled Then
            LogWordCommandDiagnostic("CommentSelection.Startup", "tracking_off")
            LogWordSelectionTrackingStatus("StartupDisabled")
            Return
        End If

        ' Preserve the original order and independence: DocCheck was bound before
        ' the comment tracker, and remains available if only tracker startup fails.
        StartDocCheckSelectionTracking()
        StartCommentSelectionTracking()
        LogWordSelectionTrackingStatus("Startup")
    End Sub

    Private Sub StartCommentSelectionTracking()
        If Not WordSelectionTrackingEnabled OrElse _commentSelectionTracker IsNot Nothing OrElse _retiredCommentSelectionTracker IsNot Nothing Then Return
        If GetWordEventSubscriptionStatus("Comment.DocumentChange") <> WordEventSubscriptionStatus.Inactive OrElse
           GetWordEventSubscriptionStatus("Comment.WindowSelectionChange") <> WordEventSubscriptionStatus.Inactive Then Return
        Try
            _commentProbeTimer = New System.Windows.Forms.Timer() With {.Interval = 75}
            AddHandler _commentProbeTimer.Tick, AddressOf ProcessPendingCommentProbe
            _commentSelectionTracker = New WordCommentSelectionTracker()
            _commentSelectionApplication = Application
            If Not WordSelectionTrackingEnabled Then
                StopCommentSelectionTracking()
                Return
            End If
            If _commentSelectionApplication Is Nothing Then Throw New System.InvalidOperationException("The Word event source is unavailable.")
            If Not SetWordEventSubscription("Comment.DocumentChange", True,
                Sub()
                    AddHandler _commentSelectionApplication.DocumentChange, AddressOf CommentSelectionDocumentChanged
                End Sub) Then
                StopCommentSelectionTracking()
                Return
            End If
            If Not SetWordEventSubscription("Comment.WindowSelectionChange", True,
                Sub()
                    AddHandler _commentSelectionApplication.WindowSelectionChange, AddressOf CommentSelectionBodyChanged
                End Sub) Then
                StopCommentSelectionTracking()
                Return
            End If
            WatchActiveCommentWindow()
        Catch ex As System.Exception
            StopCommentSelectionTracking()
            LogWordCommandDiagnostic("CommentSelection.Startup", ex.ToString())
        End Try
    End Sub

    Private Sub StopCommentSelectionTracking()
        If _commentProbeTimer IsNot Nothing Then
            _commentProbeTimer.Stop()
            RemoveHandler _commentProbeTimer.Tick, AddressOf ProcessPendingCommentProbe
            _commentProbeTimer.Dispose()
            _commentProbeTimer = Nothing
        End If
        _commentProbePending = False
        _commentContextDocument = Nothing
        _commentContextWindow = 0
        If _commentSelectionTracker IsNot Nothing Then
            _retiredCommentSelectionTracker = _commentSelectionTracker
            _commentSelectionTracker = Nothing
        End If
        If _commentSelectionApplication IsNot Nothing Then
            SetWordEventSubscription("Comment.WindowSelectionChange", False,
                Sub()
                    RemoveHandler _commentSelectionApplication.WindowSelectionChange, AddressOf CommentSelectionBodyChanged
                End Sub)
            SetWordEventSubscription("Comment.DocumentChange", False,
                Sub()
                    RemoveHandler _commentSelectionApplication.DocumentChange, AddressOf CommentSelectionDocumentChanged
                End Sub)
            If GetWordEventSubscriptionStatus("Comment.WindowSelectionChange") = WordEventSubscriptionStatus.Inactive AndAlso
               GetWordEventSubscriptionStatus("Comment.DocumentChange") = WordEventSubscriptionStatus.Inactive Then
                _commentSelectionApplication = Nothing
            End If
        End If
        If _retiredCommentSelectionTracker IsNot Nothing Then
            Try
                _retiredCommentSelectionTracker.Dispose()
                _retiredCommentSelectionTracker = Nothing
            Catch ex As System.Exception
                LogWordCommandDiagnostic("CommentSelection.Dispose", ex.ToString())
            End Try
        End If
    End Sub

    Private Sub LogWordSelectionTrackingStatus(reason As System.String)
        If _context Is Nothing OrElse Not _context.INI_APIDebug Then Return
        Try
            Dim configured As System.Boolean = _context.INI_WordCommentSelectionTracking
            Dim effective As System.Boolean = _wordSelectionModeInitialized AndAlso Not _commentSelectionTrackingDisabled
            Dim details As New System.Text.StringBuilder()
            details.Append("reason=").Append(reason)
            Using process As System.Diagnostics.Process = System.Diagnostics.Process.GetCurrentProcess()
                details.Append("; processId=").Append(process.Id.ToString(System.Globalization.CultureInfo.InvariantCulture))
            End Using
            details.Append("; version=").Append(Version)
            details.Append("; WordCommentSelectionTracking=").Append(configured.ToString())
            details.Append("; modeInitialized=").Append(_wordSelectionModeInitialized.ToString())
            details.Append("; effectiveStartupEnabled=").Append(effective.ToString())
            details.Append("; restartRequired=").Append((_wordSelectionModeInitialized AndAlso configured <> effective).ToString())
            details.Append("; shutdown=").Append(_wordSelectionShutdown.ToString())
            details.Append("; commentCallbackMode=DeferredSTA")
            details.Append("; commentProbePending=").Append(_commentProbePending.ToString())
            details.Append("; commentProbeRunning=").Append(_commentProbeRunning.ToString())
            details.Append("; trackerCreated=").Append((_commentSelectionTracker IsNot Nothing).ToString())
            details.Append("; trackerCleanupPending=").Append((_retiredCommentSelectionTracker IsNot Nothing).ToString())
            For Each eventName As System.String In New System.String() {
                "Comment.WindowSelectionChange", "Comment.DocumentChange", "DocCheck.WindowSelectionChange",
                "Startup.WindowActivate", "Startup.DocumentOpen", "Startup.NewDocument",
                "Startup.ProtectedViewWindowOpen", "Startup.ProtectedViewWindowBeforeEdit",
                "Startup.ProtectedViewWindowActivate", "Startup.DocumentChange"}
                details.Append("; ").Append(eventName).Append("=").Append(GetWordEventSubscriptionStatus(eventName).ToString())
            Next
            For Each failure As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In _wordEventSubscriptionErrors
                details.Append("; subscriptionError[").Append(failure.Key).Append("]=").Append(failure.Value)
            Next
            ' No Word selection, UI Automation probe or window activation is needed
            ' to report these states. The performance log also works without LogPath.
            LogWordCommandDiagnostic("SelectionTracking.Status", details.ToString())
            Global.SharedLibrary.SharedLibrary.PerformanceLogger.LogEvent(
                _context, "WordSelectionTracking", "Status", details.ToString(), hostName:="Word")
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("Word selection tracking status logging: " & ex.ToString())
        End Try
    End Sub

    ' Only call from Office event handlers or after EnsureUIThread. The main application
    ' HWND or the foreground window is not proof of the active document's identity.
    Private Shared Function GetActiveWordDocumentWindow(app As Microsoft.Office.Interop.Word.Application) As Microsoft.Office.Interop.Word.Window
        If app Is Nothing Then Return Nothing
        Try
            Dim docs As Microsoft.Office.Interop.Word.Documents = app.Documents
            If docs Is Nothing OrElse docs.Count = 0 Then Return Nothing
            Dim windows As Microsoft.Office.Interop.Word.Windows = app.Windows
            If windows Is Nothing OrElse windows.Count = 0 Then Return Nothing
            Return app.ActiveWindow
        Catch closed As System.Runtime.InteropServices.COMException When closed.ErrorCode = &H800A1098I
            ' The last document closed between the collection and ActiveWindow calls.
            Return Nothing
        End Try
    End Function

    Private Sub WatchActiveCommentWindow()
        Dim tracker As WordCommentSelectionTracker = _commentSelectionTracker
        If tracker Is Nothing Then Return
        Try
            Dim window As Microsoft.Office.Interop.Word.Window = GetActiveWordDocumentWindow(Application)
            tracker.WatchWindow(If(window Is Nothing, System.IntPtr.Zero, New System.IntPtr(window.Hwnd)))
        Catch closed As System.Runtime.InteropServices.COMException When closed.ErrorCode = &H800A1098I
            tracker.WatchWindow(System.IntPtr.Zero)
        Catch ex As System.Exception
            tracker.WatchWindow(System.IntPtr.Zero)
            LogWordCommandDiagnostic("CommentSelection.ActiveWindow", ex.ToString())
        End Try
    End Sub

    Private Sub CommentSelectionShutdown() Handles Me.Shutdown, Application.Quit
        _wordSelectionShutdown = True
        If _wordSelectionShutdownInProgress Then Return
        _wordSelectionShutdownInProgress = True
        Try
            StopCommentSelectionTracking()
            StopDocCheckSelectionTracking()
            RemoveStartupHandlers()
            LogWordSelectionTrackingStatus("Shutdown")
        Finally
            _wordSelectionShutdownInProgress = False
        End Try
    End Sub

    ' Event callbacks only schedule work. Never query Word COM from selection events.
    Private Sub CommentSelectionDocumentChanged()
        If Not WordSelectionTrackingEnabled Then Return
        _commentContextRevision += 1
        QueueCommentProbe()
    End Sub

    Private Sub CommentSelectionBodyChanged(selection As Microsoft.Office.Interop.Word.Selection)
        If Not WordSelectionTrackingEnabled Then Return
        QueueCommentProbe()
    End Sub

    Private Sub QueueCommentProbe()
        If _commentSelectionTracker Is Nothing OrElse _commentProbeTimer Is Nothing Then Return
        _commentProbePending = True
        _commentProbeTimer.Start()
    End Sub

    Private Sub ProcessPendingCommentProbe(sender As System.Object, args As System.EventArgs)
        If _commentProbeRunning OrElse Not WordSelectionTrackingEnabled Then Return
        If _commentProbeTimer IsNot Nothing Then _commentProbeTimer.Stop()
        If Not _commentProbePending OrElse _commentSelectionTracker Is Nothing Then Return
        _commentProbePending = False
        _commentProbeRunning = True
        Dim revision As System.Int32 = _commentContextRevision
        Try
            Dim window As Microsoft.Office.Interop.Word.Window = GetActiveWordDocumentWindow(Application)
            If window Is Nothing Then
                _commentSelectionTracker.Invalidate()
                _commentSelectionTracker.WatchWindow(System.IntPtr.Zero)
                _commentContextDocument = Nothing
                _commentContextWindow = 0
                Return
            End If
            Dim doc As Microsoft.Office.Interop.Word.Document = window.Document
            Dim hwnd As System.Int32 = window.Hwnd
            If revision <> _commentContextRevision Then Return
            If hwnd <> _commentContextWindow OrElse Not IsSameCommentDocument(doc, _commentContextDocument) Then
                ' Switching windows must preserve evidence already observed in the NEW
                ' window. Only a reused HWND with a different document needs a full reset.
                If hwnd = _commentContextWindow Then _commentSelectionTracker.Invalidate()
                LogCommentCommandResolution("context_switch; oldWindow=" & _commentContextWindow.ToString() &
                    "; activeWindow=" & hwnd.ToString() & "; document=" & CommentDocumentIdentity(doc))
                _commentContextDocument = doc
                _commentContextWindow = hwnd
                _lastCommentBodyWindow = 0
                _lastCommentBodyStart = -1
                _lastCommentBodyEnd = -1
                _lastCommentBodyStory = -1
            End If
            _commentSelectionTracker.WatchWindow(New System.IntPtr(hwnd))
            Dim selection As Microsoft.Office.Interop.Word.Selection = window.Selection
            If selection Is Nothing OrElse selection.StoryType = Microsoft.Office.Interop.Word.WdStoryType.wdCommentsStory OrElse
               CBool(selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdInCommentPane)) Then Return
            Dim currentRange As Microsoft.Office.Interop.Word.Range = selection.Range
            If revision <> _commentContextRevision OrElse currentRange Is Nothing OrElse
               Not IsSameCommentDocument(doc, currentRange.Document) Then Return
            If hwnd <> _lastCommentBodyWindow OrElse currentRange.Start <> _lastCommentBodyStart OrElse
               currentRange.End <> _lastCommentBodyEnd OrElse CInt(currentRange.StoryType) <> _lastCommentBodyStory Then
                _lastCommentBodyWindow = hwnd
                _lastCommentBodyStart = currentRange.Start
                _lastCommentBodyEnd = currentRange.End
                _lastCommentBodyStory = CInt(currentRange.StoryType)
                _commentSelectionTracker.ObserveBodySelectionChange(New System.IntPtr(hwnd))
            End If
        Catch ex As System.Exception
            LogWordCommandDiagnostic("CommentSelection.DeferredProbe", ex.ToString())
        Finally
            _commentProbeRunning = False
            If _commentProbePending AndAlso _commentProbeTimer IsNot Nothing Then _commentProbeTimer.Start()
        End Try
    End Sub

    Private Shared Function CommentDocumentIdentity(doc As Microsoft.Office.Interop.Word.Document) As System.String
        If doc Is Nothing Then Return "none"
        Dim identity As System.IntPtr = System.IntPtr.Zero
        Try
            identity = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(doc)
            Return identity.ToInt64().ToString("X", System.Globalization.CultureInfo.InvariantCulture)
        Finally
            If identity <> System.IntPtr.Zero Then System.Runtime.InteropServices.Marshal.Release(identity)
        End Try
    End Function

    Private Shared Function IsSameCommentDocument(first As Microsoft.Office.Interop.Word.Document,
                                                   second As Microsoft.Office.Interop.Word.Document) As System.Boolean
        If first Is Nothing OrElse second Is Nothing Then Return False
        If first Is second Then Return True
        Return System.String.Equals(CommentDocumentIdentity(first), CommentDocumentIdentity(second), System.StringComparison.Ordinal)
    End Function

    Private Async Function CaptureCommentSelectionEvidenceAsync() As System.Threading.Tasks.Task(Of WordCommentSelectionTracker.Evidence)
        Await EnsureUIThread()
        QueueCommentProbe()
        ProcessPendingCommentProbe(Nothing, System.EventArgs.Empty)
        Dim before As Microsoft.Office.Interop.Word.Window = GetActiveWordDocumentWindow(Globals.ThisAddIn.Application)
        If before Is Nothing Then Throw New WordCommandNoticeException("Open a document and select its comment first.")
        Dim doc As Microsoft.Office.Interop.Word.Document = before.Document
        Dim hwnd As New System.IntPtr(before.Hwnd)
        Dim tracker As WordCommentSelectionTracker = _commentSelectionTracker
        If tracker Is Nothing Then
            LogCommentCommandResolution("no_background_tracker; using_word_selection")
            Return WordCommentSelectionTracker.Evidence.Unknown("tracker_unavailable")
        End If
        Dim evidence As WordCommentSelectionTracker.Evidence = Await tracker.CaptureAsync(hwnd)
        Await EnsureUIThread()
        Dim after As Microsoft.Office.Interop.Word.Window = GetActiveWordDocumentWindow(Globals.ThisAddIn.Application)
        If after Is Nothing OrElse after.Hwnd <> hwnd.ToInt32() OrElse Not IsSameCommentDocument(doc, after.Document) Then
            Throw New WordCommandNoticeException("The active document changed. Please select the comment again and retry.")
        End If
        LogCommentCommandResolution("uia; activeWindow=" & hwnd.ToString() & "; evidenceWindow=" & evidence.WindowHandle.ToString() &
            "; document=" & CommentDocumentIdentity(doc) & "; " & evidence.Diagnostic)
        Return evidence
    End Function

    Private Function TryResolveCommentEvidence(doc As Microsoft.Office.Interop.Word.Document,
                                              evidence As WordCommentSelectionTracker.Evidence,
                                              ByRef selected As Microsoft.Office.Interop.Word.Comment) As System.Boolean
        selected = Nothing
        If evidence Is Nothing OrElse Not evidence.KnownFocus OrElse Not evidence.IsComment Then Return False
        Try
            Dim matches As New System.Collections.Generic.List(Of Microsoft.Office.Interop.Word.Comment)()
            For Each entry As CommentChoice In ReadCommandCommentCatalog(doc)
                AddMatchingCommentEvidence(doc, evidence, entry.Comment, matches)
            Next
            If matches.Count <> 1 Then
                LogCommentCommandResolution("uia_match_count=" & matches.Count.ToString(System.Globalization.CultureInfo.InvariantCulture))
                Return False
            End If
            selected = matches(0)
            LogCommentCommandResolution("uia_unique_exact_match")
            Return True
        Catch ex As System.Exception
            ' An unreadable competitor prevents proving uniqueness. Do not ignore it
            ' and silently choose a possibly different comment.
            LogWordCommandDiagnostic("CommentSelection.Match", ex.ToString())
            selected = Nothing
            Return False
        End Try
    End Function

    Private Shared Sub AddMatchingCommentEvidence(doc As Microsoft.Office.Interop.Word.Document,
                                                  evidence As WordCommentSelectionTracker.Evidence,
                                                  candidate As Microsoft.Office.Interop.Word.Comment,
                                                  matches As System.Collections.Generic.List(Of Microsoft.Office.Interop.Word.Comment))
        Dim range As Microsoft.Office.Interop.Word.Range = candidate.Range
        If Not IsSameCommentDocument(doc, range.Document) Then Return
        If Not WordCommentSelectionTracker.Matches(evidence, range.Text, candidate.Author) Then Return
        ' Some Word builds include a reply in both enumerations. Count its identity
        ' once, but never merge different comments merely because their text matches.
        For Each existing As Microsoft.Office.Interop.Word.Comment In matches
            If SameCommandComment(existing, candidate) Then Return
        Next
        matches.Add(candidate)
    End Sub

    Private NotInheritable Class CommentChoice
        Public Comment As Microsoft.Office.Interop.Word.Comment
        Public Text As System.String
        Public Anchor As Microsoft.Office.Interop.Word.Range
        Public Author As System.String
        Public IsReply As System.Boolean
    End Class

    ' Use Count/Item, NEVER the COM enumerator of Comment.Replies. A collection's
    ' enumerator is not proof that an item is a child (the reported regression
    ' produced all document comments even for empty reply collections).
    ' All command paths share this catalog: picker, UIA identity and classic story.
    Private Function ReadCommandCommentCatalog(doc As Microsoft.Office.Interop.Word.Document,
                                               Optional allowedRoots As System.Collections.Generic.IReadOnlyList(Of Microsoft.Office.Interop.Word.Comment) = Nothing) As System.Collections.Generic.List(Of CommentChoice)
        Dim entries As New System.Collections.Generic.List(Of CommentChoice)()
        Dim identityBuckets As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of CommentChoice))(System.StringComparer.Ordinal)
        If allowedRoots IsNot Nothing Then
            For Each root As Microsoft.Office.Interop.Word.Comment In allowedRoots
                AddCatalogComment(doc, root, entries, identityBuckets, False)
            Next
            Return entries
        End If
        Dim comments As Microsoft.Office.Interop.Word.Comments = doc.Comments
        Dim count As System.Int32 = comments.Count
        For index As System.Int32 = 1 To count
            Dim comment As Microsoft.Office.Interop.Word.Comment = comments(index)
            Dim ancestor As Microsoft.Office.Interop.Word.Comment = Nothing
            Try
                ancestor = comment.Ancestor
            Catch unsupported As System.Runtime.InteropServices.COMException
                LogWordCommandDiagnostic("CommentSelection.AncestorUnavailable", unsupported.ToString())
            End Try
            Dim isReply As System.Boolean = ancestor IsNot Nothing AndAlso Not SameCommandComment(ancestor, comment)
            AddCatalogComment(doc, comment, entries, identityBuckets, isReply)
            ' A reply cannot own a new thread. Do not walk it recursively.
            If isReply Then Continue For
            Dim replies As Microsoft.Office.Interop.Word.Comments = comment.Replies
            Dim replyCount As System.Int32 = replies.Count
            For replyIndex As System.Int32 = 1 To replyCount
                Dim reply As Microsoft.Office.Interop.Word.Comment = replies(replyIndex)
                If SameCommandComment(reply, comment) Then
                    LogCommentCommandResolution("catalog_rejected_self_reply")
                    Continue For
                End If
                Dim parent As Microsoft.Office.Interop.Word.Comment = reply.Ancestor
                If parent Is Nothing OrElse Not SameCommandComment(parent, comment) Then
                    LogCommentCommandResolution("catalog_rejected_unrelated_reply")
                    Continue For
                End If
                AddCatalogComment(doc, reply, entries, identityBuckets, True)
            Next
        Next
        If comments.Count <> count Then
            Throw New WordCommandNoticeException("The comments changed while Inky was reading them. Please select the comment again.")
        End If
        Return entries
    End Function

    Private Shared Sub AddCatalogComment(doc As Microsoft.Office.Interop.Word.Document,
                                         comment As Microsoft.Office.Interop.Word.Comment,
                                         entries As System.Collections.Generic.List(Of CommentChoice),
                                         identityBuckets As System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of CommentChoice)),
                                         isReply As System.Boolean)
        Dim body As Microsoft.Office.Interop.Word.Range = comment.Range
        Dim anchor As Microsoft.Office.Interop.Word.Range = comment.Scope
        If Not IsSameCommentDocument(doc, body.Document) OrElse
           Not IsSameCommentDocument(doc, anchor.Document) OrElse
           body.StoryType <> Microsoft.Office.Interop.Word.WdStoryType.wdCommentsStory Then
            Throw New WordCommandNoticeException("A comment is no longer available. Please select it again.")
        End If
        ' Use range coordinates only as a bucket, not as final identity. This
        ' avoids comparing every comment with every other comment in large files.
        ' A bucket collision must still pass document/body/scope/reference equality.
        Dim key As System.String = CInt(body.StoryType).ToString(System.Globalization.CultureInfo.InvariantCulture) & ":" &
            body.Start.ToString(System.Globalization.CultureInfo.InvariantCulture) & ":" &
            body.End.ToString(System.Globalization.CultureInfo.InvariantCulture)
        Dim bucket As System.Collections.Generic.List(Of CommentChoice) = Nothing
        If identityBuckets.TryGetValue(key, bucket) Then
            For Each existing As CommentChoice In bucket
                If SameCommandComment(existing.Comment, comment) Then
                    existing.IsReply = existing.IsReply OrElse isReply
                    Return
                End If
            Next
        Else
            bucket = New System.Collections.Generic.List(Of CommentChoice)()
            identityBuckets.Add(key, bucket)
        End If
        Dim entry As New CommentChoice With {
            .Comment = comment, .Text = If(body.Text, System.String.Empty),
            .Anchor = anchor.Duplicate, .Author = comment.Author, .IsReply = isReply
        }
        entries.Add(entry)
        bucket.Add(entry)
    End Sub

    Private Shared Function SameCommandComment(left As Microsoft.Office.Interop.Word.Comment,
                                                right As Microsoft.Office.Interop.Word.Comment) As System.Boolean
        If left Is Nothing OrElse right Is Nothing Then Return False
        If left Is right Then Return True
        Dim leftBody As Microsoft.Office.Interop.Word.Range = left.Range
        Dim rightBody As Microsoft.Office.Interop.Word.Range = right.Range
        ' RCW reference equality and comment text are NOT durable identity. Keep
        ' equal-text comments at different anchors (or different comment ranges).
        Return IsSameCommentDocument(leftBody.Document, rightBody.Document) AndAlso
            leftBody.StoryType = rightBody.StoryType AndAlso leftBody.IsEqual(rightBody) AndAlso
            left.Scope.IsEqual(right.Scope) AndAlso left.Reference.IsEqual(right.Reference)
    End Function

    ' <summary>
    ' Finds root comments whose actual document anchors meet a main-text selection.
    ' This is separate from the selected-comment-card resolver: selecting text in
    ' the document must not require UI Automation focus evidence or a picker.
    ' Reply uses ALL matches; Apply Comment may use a single unambiguous match.
    ' No selection, document position or comment is changed here.
    ' </summary>
    Private Function FindCommandCommentsInBodyRange(
        doc As Microsoft.Office.Interop.Word.Document,
        bodyRange As Microsoft.Office.Interop.Word.Range) As System.Collections.Generic.List(Of Microsoft.Office.Interop.Word.Comment)

        Dim matches As New System.Collections.Generic.List(Of Microsoft.Office.Interop.Word.Comment)()
        If doc Is Nothing OrElse bodyRange Is Nothing OrElse
           bodyRange.StoryType <> Microsoft.Office.Interop.Word.WdStoryType.wdMainTextStory OrElse
           Not IsSameCommentDocument(doc, bodyRange.Document) Then Return matches

        Dim startPos As System.Int32 = bodyRange.Start
        Dim endPos As System.Int32 = bodyRange.End
        Dim comments As Microsoft.Office.Interop.Word.Comments = doc.Comments
        Dim expectedCount As System.Int32 = comments.Count
        For index As System.Int32 = 1 To expectedCount
            Dim candidate As Microsoft.Office.Interop.Word.Comment = comments(index)
            Dim ancestor As Microsoft.Office.Interop.Word.Comment = Nothing
            Try
                ancestor = candidate.Ancestor
            Catch unsupported As System.Runtime.InteropServices.COMException
                ' Classic Word versions may not expose reply ancestry.
            End Try
            If ancestor IsNot Nothing AndAlso Not SameCommandComment(candidate, ancestor) Then Continue For

            Dim anchor As Microsoft.Office.Interop.Word.Range = Nothing
            Try
                anchor = candidate.Scope
            Catch noScope As System.Runtime.InteropServices.COMException
                anchor = candidate.Reference
            End Try
            If anchor Is Nothing OrElse
               anchor.StoryType <> Microsoft.Office.Interop.Word.WdStoryType.wdMainTextStory OrElse
               Not IsSameCommentDocument(doc, anchor.Document) Then Continue For

            Dim intersects As System.Boolean
            If startPos = endPos Then
                ' Word's insertion point at either end of the annotated text is
                ' still a useful target. Include zero-width comment anchors.
                intersects = startPos >= anchor.Start AndAlso startPos <= anchor.End
            ElseIf anchor.Start = anchor.End Then
                intersects = anchor.Start >= startPos AndAlso anchor.Start < endPos
            Else
                intersects = startPos < anchor.End AndAlso endPos > anchor.Start
            End If
            If Not intersects Then Continue For
            Dim alreadyIncluded As System.Boolean = False
            For Each existing As Microsoft.Office.Interop.Word.Comment In matches
                If SameCommandComment(existing, candidate) Then
                    alreadyIncluded = True
                    Exit For
                End If
            Next
            If Not alreadyIncluded Then matches.Add(candidate)
        Next
        If comments.Count <> expectedCount Then
            Throw New WordCommandNoticeException("The comments changed. Please try your selection again.")
        End If
        Return matches
    End Function

    Private Function ChooseCommentForCommand(doc As Microsoft.Office.Interop.Word.Document,
                                              actionLabel As System.String,
                                              Optional allowedComments As System.Collections.Generic.IReadOnlyList(Of Microsoft.Office.Interop.Word.Comment) = Nothing,
                                              Optional allowDocumentSelection As System.Boolean = False) As Microsoft.Office.Interop.Word.Comment
        Dim choices As New System.Collections.Generic.List(Of CommentChoice)()
        Dim items As New System.Collections.Generic.List(Of Global.SharedLibrary.SharedLibrary.SharedMethods.SelectionItem)()
        For Each entry As CommentChoice In ReadCommandCommentCatalog(doc, allowedComments)
            If System.String.IsNullOrWhiteSpace(entry.Text) Then Continue For
            choices.Add(entry)
            Dim label As System.String = If(entry.IsReply, "Reply", "Comment") & " " &
                choices.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " - " & entry.Author &
                ": " & CommentChoicePreview(entry.Text, 100) & " | Text: " & CommentChoicePreview(entry.Anchor.Text, 60)
            items.Add(New Global.SharedLibrary.SharedLibrary.SharedMethods.SelectionItem(label, choices.Count))
        Next
        If allowDocumentSelection Then
            items.Add(New Global.SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Use the comments in the selected document text", choices.Count + 1))
        End If
        If choices.Count = 0 Then
            Throw New WordCommandNoticeException("There is no readable comment in this document. Please select a comment and try again.")
        End If
        LogCommentCommandResolution("explicit_picker; action=" & actionLabel &
                                    "; choices=" & choices.Count.ToString(System.Globalization.CultureInfo.InvariantCulture))
        ' Reuse the existing helper: it validates the owner on the current Office
        ' thread according to SharedMethods.DialogOwner.vb. No new custom window.
        Dim picked As System.Int32 = Global.SharedLibrary.SharedLibrary.SharedMethods.SelectValue(
            items, 0,
            "Please choose the comment you want Inky to use. Nothing will change until you confirm.",
            "Inky - " & actionLabel)
        ' Only an explicit choice may retain the caller's captured document selection.
        If allowDocumentSelection AndAlso picked = choices.Count + 1 Then
            LogCommentCommandResolution("explicit_picker; confirmed_document_selection")
            Return Nothing
        End If
        If picked < 1 OrElse picked > choices.Count Then
            Throw New System.OperationCanceledException("Comment selection cancelled by user.")
        End If
        Dim chosen As CommentChoice = choices(picked - 1)
        Try
            If Not IsSameCommentDocument(doc, Globals.ThisAddIn.Application.ActiveDocument) OrElse
               Not IsSameCommentDocument(doc, chosen.Comment.Range.Document) OrElse
               Not chosen.Comment.Scope.IsEqual(chosen.Anchor) OrElse
               Not System.String.Equals(chosen.Comment.Range.Text, chosen.Text, System.StringComparison.Ordinal) Then
                Throw New WordCommandNoticeException("The comment or document changed while the list was open. Please select the comment again.")
            End If
        Catch ex As WordCommandNoticeException
            Throw
        Catch ex As System.Exception
            LogWordCommandDiagnostic("CommentSelection.RevalidateChoice", ex.ToString())
            Throw New WordCommandNoticeException("This comment is no longer available. Please select another comment and try again.")
        End Try
        Return chosen.Comment
    End Function

    Private Shared Function CommentChoicePreview(value As System.String, maximum As System.Int32) As System.String
        Dim preview As System.String = WordCommentSelectionTracker.NormalizeText(value)
        If preview.Length > maximum Then Return preview.Substring(0, maximum) & "..."
        Return preview
    End Function

    ' Only this exception may carry a deliberately authored user-facing notice.
    ' COM, networking and model exceptions are never copied into a dialog.
    Private NotInheritable Class WordCommandNoticeException
        Inherits System.Exception
        Public Sub New(userMessage As System.String, Optional inner As System.Exception = Nothing)
            MyBase.New(userMessage, inner)
        End Sub
    End Class

    Private Shared Function IsWordCommandCancellation(ex As System.Exception) As System.Boolean
        If TypeOf ex Is System.OperationCanceledException Then Return True
        Dim aggregate As System.AggregateException = TryCast(ex, System.AggregateException)
        If aggregate IsNot Nothing Then
            Dim errors As System.Collections.ObjectModel.ReadOnlyCollection(Of System.Exception) = aggregate.Flatten().InnerExceptions
            If errors.Count = 0 Then Return False
            For Each inner As System.Exception In errors
                If Not IsWordCommandCancellation(inner) Then Return False
            Next
            Return True
        End If
        Dim invocation As System.Reflection.TargetInvocationException = TryCast(ex, System.Reflection.TargetInvocationException)
        Return invocation IsNot Nothing AndAlso invocation.InnerException IsNot Nothing AndAlso
            IsWordCommandCancellation(invocation.InnerException)
    End Function

    ' Wrap only exceptions raised by an actual model or post-correction call.
    ' This preserves the distinction between an AI-service failure and an
    ' unrelated Word/COM/formatting problem, without leaking procedure names.
    Private Async Function AwaitWordModelRequestAsync(
        request As System.Func(Of System.Threading.Tasks.Task(Of System.String))) _
        As System.Threading.Tasks.Task(Of System.String)
        Try
            Return Await request().ConfigureAwait(False)
        Catch ex As System.OperationCanceledException
            Throw
        Catch ex As System.Exception When IsWordCommandCancellation(ex)
            Throw
        Catch ex As System.Exception
            Throw New WordCommandNoticeException(
                Global.SharedLibrary.SharedLibrary.SharedMethods.FormatModelRequestFailureForUser(ex), ex)
        End Try
    End Function

    Private Sub ReportWordCommandFailure(operation As System.String, ex As System.Exception,
                                         Optional userMessage As System.String = Nothing)
        If IsWordCommandCancellation(ex) Then
            LogWordCommandDiagnostic(operation, "cancelled_by_user")
            Return
        End If
        LogWordCommandDiagnostic(operation, ex.ToString())
        Dim notice As WordCommandNoticeException = TryCast(ex, WordCommandNoticeException)
        If notice IsNot Nothing Then
            userMessage = notice.Message
        ElseIf TypeOf ex Is System.TimeoutException Then
            userMessage = "Inky did not receive an answer in time. Please try again."
        ElseIf System.String.IsNullOrWhiteSpace(userMessage) Then
            userMessage = "Inky could not complete this request. Please try again. If the problem persists, contact your support team."
        End If
        Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox(userMessage, "Inky")
    End Sub

    Private Sub LogWordCommandDiagnostic(operation As System.String, detail As System.String)
        Dim message As System.String = "Word command [" & operation & "]: " & detail
        System.Diagnostics.Debug.WriteLine(message)
        Try
            If _context IsNot Nothing AndAlso _context.INI_APIDebug Then
                Global.SharedLibrary.SharedLogger.Log(_context, _context.RDV, message)
            End If
        Catch logError As System.Exception
            System.Diagnostics.Debug.WriteLine("Word command diagnostic logging: " & logError.ToString())
        End Try
    End Sub
End Class
