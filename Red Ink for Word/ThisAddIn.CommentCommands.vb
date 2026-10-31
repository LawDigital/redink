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
    Private _commentSelectionTracker As WordCommentSelectionTracker
    Private _lastCommentBodyWindow As System.Int32
    Private _lastCommentBodyStart As System.Int32 = -1
    Private _lastCommentBodyEnd As System.Int32 = -1
    Private _lastCommentBodyStory As System.Int32 = -1

    Private Sub CommentSelectionStartup() Handles Me.Startup
        Try
            _commentSelectionTracker = New WordCommentSelectionTracker()
            WatchActiveCommentWindow()
        Catch ex As System.Exception
            LogWordCommandDiagnostic("CommentSelection.Startup", ex.ToString())
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
        Dim tracker As WordCommentSelectionTracker = _commentSelectionTracker
        _commentSelectionTracker = Nothing
        If tracker IsNot Nothing Then tracker.Dispose()
    End Sub

    Private Sub CommentSelectionDocumentChanged() Handles Application.DocumentChange
        _lastCommentBodyWindow = 0
        _lastCommentBodyStart = -1
        _lastCommentBodyEnd = -1
        _lastCommentBodyStory = -1
        If _commentSelectionTracker IsNot Nothing Then
            _commentSelectionTracker.Invalidate()
            WatchActiveCommentWindow()
        End If
    End Sub

    Private Sub CommentSelectionBodyChanged(selection As Microsoft.Office.Interop.Word.Selection) Handles Application.WindowSelectionChange
        Try
            If _commentSelectionTracker Is Nothing OrElse selection Is Nothing OrElse
               selection.StoryType = Microsoft.Office.Interop.Word.WdStoryType.wdCommentsStory OrElse
               CBool(selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdInCommentPane)) Then Return
            Dim currentRange As Microsoft.Office.Interop.Word.Range = selection.Range
            If currentRange Is Nothing Then Return
            Dim window As Microsoft.Office.Interop.Word.Window = GetActiveWordDocumentWindow(selection.Application)
            If window Is Nothing OrElse Not IsSameProcessingDocument(window.Document, currentRange.Document) Then Return
            Dim hwnd As System.Int32 = window.Hwnd
            If hwnd <> _lastCommentBodyWindow OrElse currentRange.Start <> _lastCommentBodyStart OrElse
               currentRange.End <> _lastCommentBodyEnd OrElse CInt(currentRange.StoryType) <> _lastCommentBodyStory Then
                _lastCommentBodyWindow = hwnd
                _lastCommentBodyStart = currentRange.Start
                _lastCommentBodyEnd = currentRange.End
                _lastCommentBodyStory = CInt(currentRange.StoryType)
                ' Word may move the body anchor when a card is selected, or restore
                ' the old body selection on a ribbon click. That is NOT a document
                ' change. Let the input/UIA observer distinguish those transitions.
                _commentSelectionTracker.ObserveBodySelectionChange(New System.IntPtr(hwnd))
            End If
        Catch closed As System.Runtime.InteropServices.COMException When closed.ErrorCode = &H800A1098I
            ' A final selection event can race the last document closing.
            If _commentSelectionTracker IsNot Nothing Then _commentSelectionTracker.WatchWindow(System.IntPtr.Zero)
        Catch ex As System.Exception
            LogWordCommandDiagnostic("CommentSelection.SelectionEvent", ex.ToString())
        End Try
    End Sub

    Private Async Function CaptureCommentSelectionEvidenceAsync() As System.Threading.Tasks.Task(Of WordCommentSelectionTracker.Evidence)
        Await EnsureUIThread()
        Dim before As Microsoft.Office.Interop.Word.Window = GetActiveWordDocumentWindow(Globals.ThisAddIn.Application)
        If before Is Nothing Then Throw New WordCommandNoticeException("Open a document and select its comment first.")
        Dim doc As Microsoft.Office.Interop.Word.Document = before.Document
        Dim hwnd As New System.IntPtr(before.Hwnd)
        Dim tracker As WordCommentSelectionTracker = _commentSelectionTracker
        If tracker Is Nothing Then Return WordCommentSelectionTracker.Evidence.Unknown("tracker_unavailable")
        Dim evidence As WordCommentSelectionTracker.Evidence = Await tracker.CaptureAsync(hwnd)
        Await EnsureUIThread()
        Dim after As Microsoft.Office.Interop.Word.Window = GetActiveWordDocumentWindow(Globals.ThisAddIn.Application)
        If after Is Nothing OrElse after.Hwnd <> hwnd.ToInt32() OrElse Not IsSameProcessingDocument(doc, after.Document) Then
            Throw New WordCommandNoticeException("The active document changed. Please select the comment again and retry.")
        End If
        LogCommentCommandResolution("uia; " & evidence.Diagnostic)
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
        If Not IsSameProcessingDocument(doc, range.Document) Then Return
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
    Private Function ReadCommandCommentCatalog(doc As Microsoft.Office.Interop.Word.Document) As System.Collections.Generic.List(Of CommentChoice)
        Dim entries As New System.Collections.Generic.List(Of CommentChoice)()
        Dim identityBuckets As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of CommentChoice))(System.StringComparer.Ordinal)
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
        If Not IsSameProcessingDocument(doc, body.Document) OrElse
           Not IsSameProcessingDocument(doc, anchor.Document) OrElse
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
        Return IsSameProcessingDocument(leftBody.Document, rightBody.Document) AndAlso
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
           Not IsSameProcessingDocument(doc, bodyRange.Document) Then Return matches

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
               Not IsSameProcessingDocument(doc, anchor.Document) Then Continue For

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
                                              Optional allowedComments As System.Collections.Generic.IReadOnlyList(Of Microsoft.Office.Interop.Word.Comment) = Nothing) As Microsoft.Office.Interop.Word.Comment
        Dim choices As New System.Collections.Generic.List(Of CommentChoice)()
        Dim items As New System.Collections.Generic.List(Of Global.SharedLibrary.SharedLibrary.SharedMethods.SelectionItem) From {
            New Global.SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Choose a comment...", 0)
        }
        For Each entry As CommentChoice In ReadCommandCommentCatalog(doc)
            If System.String.IsNullOrWhiteSpace(entry.Text) Then Continue For
            If allowedComments IsNot Nothing Then
                Dim included As System.Boolean = False
                For Each allowed As Microsoft.Office.Interop.Word.Comment In allowedComments
                    If SameCommandComment(allowed, entry.Comment) Then
                        included = True
                        Exit For
                    End If
                Next
                If Not included Then Continue For
            End If
            choices.Add(entry)
            Dim label As System.String = If(entry.IsReply, "Reply", "Comment") & " " &
                choices.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " - " & entry.Author &
                ": " & CommentChoicePreview(entry.Text, 100) & " | Text: " & CommentChoicePreview(entry.Anchor.Text, 60)
            items.Add(New Global.SharedLibrary.SharedLibrary.SharedMethods.SelectionItem(label, choices.Count))
        Next
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
        If picked < 1 OrElse picked > choices.Count Then
            Throw New System.OperationCanceledException("Comment selection cancelled by user.")
        End If
        Dim chosen As CommentChoice = choices(picked - 1)
        Try
            If Not IsSameProcessingDocument(doc, Globals.ThisAddIn.Application.ActiveDocument) OrElse
               Not IsSameProcessingDocument(doc, chosen.Comment.Range.Document) OrElse
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
