' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: WordCommentSelectionTracker.vb
' PURPOSE
'   Captures read-only evidence of the active modern Word comment using UI Automation.
'
' KEY ARCHITECTURE
'   • Observation: Tracks the focused modern comment card through UI Automation.
'   • Boundary: Produces read-only evidence without sending Word COM/UIA objects outside the
'     host.
' =============================================================================

'
' Read-only bridge for the modern Word comments UI. Word.Selection can still refer
' to an unrelated body selection after a comment card receives focus. Capture the
' selected/clicked card BEFORE the ribbon takes focus, not the comments intersecting that
' stale body range. No Word COM object, UIA object, or document text leaves this host.
'
' UIA runs exclusively on a dedicated MTA thread (including subscribe/unsubscribe).
' A read-only, Word-UI-thread message hook also observes actual clicks: selecting
' a collapsed card is NOT the same event as entering its editor. It never consumes
' an input message, sets focus, or opens a comment. No desktop/global input hook.
' Only immutable, bounded evidence crosses back to the Office UI thread. This code
' never selects, clicks, types, searches the desktop, or reads another process.
' Unknown/duplicate evidence requires an explicit user choice in the command layer.
'
' Native interface prefixes and IDs follow the Windows SDK UIAutomationClient.h.
' Unused vtable slots are declarations ONLY and must never be invoked.
' https://learn.microsoft.com/en-us/dotnet/framework/ui-automation/ui-automation-threading-issues
' https://learn.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationelement

Option Explicit On
Option Strict On

Public NotInheritable Class WordCommentSelectionTracker
    Implements System.IDisposable

    Public NotInheritable Class Evidence
        Public ReadOnly WindowHandle As System.IntPtr
        Public ReadOnly KnownFocus As System.Boolean
        Public ReadOnly IsComment As System.Boolean
        Public ReadOnly Texts As System.Collections.Generic.IReadOnlyList(Of System.String)
        Public ReadOnly Author As System.String
        Public ReadOnly Diagnostic As System.String
        Public ReadOnly DocumentGeneration As System.Int32
        Public ReadOnly SelectedText As System.String

        Friend Sub New(windowHandle As System.IntPtr, knownFocus As System.Boolean,
                       isComment As System.Boolean, texts As System.Collections.Generic.IEnumerable(Of System.String),
                       author As System.String, diagnostic As System.String,
                       Optional documentGeneration As System.Int32 = -1,
                       Optional selectedText As System.String = Nothing)
            Me.WindowHandle = windowHandle
            Me.KnownFocus = knownFocus
            Me.IsComment = isComment
            Me.Texts = New System.Collections.ObjectModel.ReadOnlyCollection(Of System.String)(
                New System.Collections.Generic.List(Of System.String)(If(texts, New System.String() {})))
            Me.Author = If(author, System.String.Empty)
            Me.Diagnostic = diagnostic
            Me.DocumentGeneration = documentGeneration
            Me.SelectedText = If(selectedText, System.String.Empty)
        End Sub

        Friend Function ForGeneration(generation As System.Int32) As Evidence
            Return New Evidence(WindowHandle, KnownFocus, IsComment, Texts, Author, Diagnostic, generation, SelectedText)
        End Function

        Friend Shared Function Unknown(reason As System.String) As Evidence
            Return New Evidence(System.IntPtr.Zero, False, False, Nothing, Nothing, reason)
        End Function
    End Class

    Private Enum Observation
        Focus
        Pointer
        Keyboard
        Selection
        TextSelection
        Capture
        Refresh
        BodyProjection
    End Enum

    Private Const MaxTextLength As System.Int32 = 65536
    Private Const MaxNodes As System.Int32 = 96
    Private Const MaxAncestors As System.Int32 = 48
    Private Const UIA_ProcessId As System.Int32 = 30002
    Private Const UIA_ControlType As System.Int32 = 30003
    Private Const UIA_Name As System.Int32 = 30005
    Private Const UIA_AutomationId As System.Int32 = 30011
    Private Const UIA_ClassName As System.Int32 = 30012
    Private Const UIA_NativeWindowHandle As System.Int32 = 30020
    Private Const UIA_ValueValue As System.Int32 = 30045
    Private Const UIA_AnnotationType As System.Int32 = 30113
    Private Const UIA_AnnotationAuthor As System.Int32 = 30115
    Private Const UIA_IsAnnotationPatternAvailable As System.Int32 = 30118
    Private Const UIA_TextPattern As System.Int32 = 10014
    Private Const UIA_CommentAnnotation As System.Int32 = 60003
    Private Const UIA_TreeItem As System.Int32 = 50024
    Private Const UIA_Document As System.Int32 = 50030
    Private Const UIA_Text As System.Int32 = 50020
    Private Const UIA_Edit As System.Int32 = 50004
    Private Const UIA_Custom As System.Int32 = 50025
    Private Const UIA_ToolBar As System.Int32 = 50021
    Private Const UIA_Menu As System.Int32 = 50009
    Private Const UIA_MenuBar As System.Int32 = 50010

    Private ReadOnly _queue As New System.Collections.Concurrent.ConcurrentQueue(Of System.Action)()
    Private ReadOnly _wake As New System.Threading.AutoResetEvent(False)
    Private ReadOnly _thread As System.Threading.Thread
    Private ReadOnly _processId As System.Int32
    Private _stopping As System.Int32
    Private _generation As System.Int32
    Private _automation As NativeAutomation
    Private _walker As NativeTreeWalker
    Private _handler As FocusHandler
    Private _selectionHandler As SelectionHandler
    Private _watchedRoot As NativeElement
    Private _watchedWindow As System.IntPtr
    Private ReadOnly _selectionEvents As System.Int32() = {20010, 20011, 20012, 20014}
    Private ReadOnly _registeredEvents As New System.Collections.Generic.List(Of System.Int32)()
    ' A native hook is not a GC root. Retain its managed delegate until Windows
    ' confirms removal, including the rare unhook-failure case during shutdown.
    Private Shared ReadOnly LiveMessageHandlers As New System.Collections.Concurrent.ConcurrentDictionary(Of System.IntPtr, MessageHook)()
    Private ReadOnly _messageHandler As MessageHook
    Private _inputHook As System.IntPtr
    Private _pendingKeyboardProbe As System.Int32
    Private _pendingBodyProbe As System.Int32
    Private _pendingSelectionEvents As System.Int32
    Private _bodyNavigationPending As System.Boolean
    Private _lastCard As NativeElement
    ' Only accessed on the MTA worker. No cross-thread COM access or UI-thread waits.
    Private _last As Evidence = Evidence.Unknown("not_initialized")
    Private _commandSurfaceWindow As System.IntPtr = System.IntPtr.Zero
    ' A failed direct-input probe must not later turn the old body focus into
    ' positive proof. Only a new verified interaction clears this uncertainty.
    Private _uncertainInputWindow As System.IntPtr = System.IntPtr.Zero
    Private _pendingFocusEvents As System.Int32

    Public Sub New()
        Using process As System.Diagnostics.Process = System.Diagnostics.Process.GetCurrentProcess()
            _processId = process.Id
        End Using
        _thread = New System.Threading.Thread(AddressOf WorkerMain) With {
            .IsBackground = True,
            .Name = "Inky Word comment focus"
        }
        _thread.SetApartmentState(System.Threading.ApartmentState.MTA)
        _thread.Start()
        ' Constructor is called by the Word startup event on its UI thread.
        ' Only that thread is observed. hMod=NULL is required for a same-process hook.
        _messageHandler = AddressOf ObserveWindowMessage
        Try
            _inputHook = SetWindowsHookEx(3, _messageHandler, System.IntPtr.Zero, GetCurrentThreadId())
            If _inputHook = System.IntPtr.Zero Then
                System.Diagnostics.Debug.WriteLine("Word comment input observer unavailable: " &
                    System.Runtime.InteropServices.Marshal.GetLastWin32Error().ToString(System.Globalization.CultureInfo.InvariantCulture))
            Else
                LiveMessageHandlers.TryAdd(_inputHook, _messageHandler)
            End If
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("Word comment input observer initialization: " & ex.ToString())
        End Try
    End Sub

    Public Async Function CaptureAsync(windowHandle As System.IntPtr) As System.Threading.Tasks.Task(Of Evidence)
        If windowHandle = System.IntPtr.Zero OrElse System.Threading.Volatile.Read(_stopping) <> 0 Then
            Return Evidence.Unknown("no_active_window")
        End If
        Dim completion As New System.Threading.Tasks.TaskCompletionSource(Of Evidence)(
            System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
        Dim generation As System.Int32 = System.Threading.Volatile.Read(_generation)
        Enqueue(Sub()
                    Try
                        ' Drop expired queued probes before making any provider calls.
                        If completion.Task.IsCompleted Then Return
                        If generation <> System.Threading.Volatile.Read(_generation) Then
                            completion.TrySetResult(Evidence.Unknown("document_changed"))
                            Return
                        End If
                        If _automation Is Nothing OrElse System.Threading.Volatile.Read(_stopping) <> 0 Then
                            completion.TrySetResult(Evidence.Unknown("uia_unavailable"))
                            Return
                        End If
                        WatchWindowCore(windowHandle)
                        If _last.IsComment AndAlso _last.WindowHandle = windowHandle AndAlso _lastCard IsNot Nothing Then
                            ReadFocus(_lastCard, Observation.Refresh, windowHandle)
                        End If
                        Dim focused As NativeElement = Nothing
                        Check(_automation.GetFocusedElement(focused))
                        ReadFocus(focused, Observation.Capture)
                        If generation <> System.Threading.Volatile.Read(_generation) Then
                            completion.TrySetResult(Evidence.Unknown("document_changed"))
                        ElseIf _last.WindowHandle = windowHandle Then
                            If completion.TrySetResult(_last.ForGeneration(generation)) Then _commandSurfaceWindow = System.IntPtr.Zero
                        Else
                            completion.TrySetResult(Evidence.Unknown("focus_window_mismatch"))
                        End If
                    Catch ex As System.Exception
                        RejectObservedInput(windowHandle, "capture_failed")
                        System.Diagnostics.Debug.WriteLine("Word comment focus capture: " & ex.ToString())
                        completion.TrySetResult(_last)
                    End Try
                End Sub)
        ' Never block Office's STA while a provider needs to send it messages.
        ' A stalled provider cannot trigger a late selection or write: probes are read-only.
        Dim finished As System.Threading.Tasks.Task = Await System.Threading.Tasks.Task.WhenAny(
            completion.Task, System.Threading.Tasks.Task.Delay(1200)).ConfigureAwait(False)
        If finished IsNot completion.Task Then completion.TrySetResult(Evidence.Unknown("uia_timeout"))
        Return Await completion.Task.ConfigureAwait(False)
    End Function

    Public Function IsCurrent(evidence As Evidence) As System.Boolean
        Return evidence IsNot Nothing AndAlso evidence.DocumentGeneration >= 0 AndAlso
            evidence.DocumentGeneration = System.Threading.Volatile.Read(_generation)
    End Function

    Public Sub Invalidate()
        System.Threading.Interlocked.Increment(_generation)
        Enqueue(Sub()
                    _last = Evidence.Unknown("document_changed")
                    _uncertainInputWindow = System.IntPtr.Zero
                    _lastCard = Nothing
                    _bodyNavigationPending = False
                    _commandSurfaceWindow = System.IntPtr.Zero
                End Sub)
    End Sub

    Public Sub WatchWindow(windowHandle As System.IntPtr)
        Enqueue(Sub() WatchWindowCore(windowHandle))
    End Sub

    Private Sub WatchWindowCore(windowHandle As System.IntPtr)
        If windowHandle = System.IntPtr.Zero Then
            ' Last document closed: unsubscribe on the UIA worker, never on Office's STA.
            RemoveSelectionSubscriptions()
            Return
        End If
        If _automation Is Nothing OrElse windowHandle = _watchedWindow OrElse Not IsWordWindow(windowHandle) Then Return
        RemoveSelectionSubscriptions()
        Dim root As NativeElement = Nothing
        Check(_automation.ElementFromHandle(windowHandle, root))
        If root Is Nothing Then Return
        _watchedRoot = root
        _watchedWindow = windowHandle
        If _selectionHandler Is Nothing Then _selectionHandler = New SelectionHandler(Me)
        For Each eventId As System.Int32 In _selectionEvents
            Try
                Check(_automation.AddAutomationEventHandler(eventId, root, 7, System.IntPtr.Zero, _selectionHandler))
                _registeredEvents.Add(eventId)
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Word comment selection subscription: " & ex.ToString())
            End Try
        Next
    End Sub

    Private Sub RemoveSelectionSubscriptions()
        If _automation IsNot Nothing AndAlso _watchedRoot IsNot Nothing Then
            For Each eventId As System.Int32 In _registeredEvents
                Try
                    Check(_automation.RemoveAutomationEventHandler(eventId, _watchedRoot, _selectionHandler))
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("Word comment selection unsubscribe: " & ex.ToString())
                End Try
            Next
        End If
        _registeredEvents.Clear()
        _watchedRoot = Nothing
        _watchedWindow = System.IntPtr.Zero
    End Sub

    Public Sub ObserveBodySelectionChange(windowHandle As System.IntPtr)
        If System.Threading.Interlocked.Exchange(_pendingBodyProbe, 1) <> 0 Then Return
        Dim generation As System.Int32 = System.Threading.Volatile.Read(_generation)
        Enqueue(Sub()
                    System.Threading.Interlocked.Exchange(_pendingBodyProbe, 0)
                    If _automation Is Nothing OrElse generation <> System.Threading.Volatile.Read(_generation) Then Return
                    Dim focused As NativeElement = Nothing
                    Check(_automation.GetFocusedElement(focused))
                    ReadFocus(focused, Observation.BodyProjection, windowHandle)
                End Sub)
    End Sub

    Private Function ObserveWindowMessage(code As System.Int32, removal As System.IntPtr,
                                           messagePointer As System.IntPtr) As System.IntPtr
        Try
            ' PM_REMOVE only: peeking a message is NOT a new user action.
            If code >= 0 AndAlso removal = New System.IntPtr(1) AndAlso
               System.Threading.Volatile.Read(_stopping) = 0 Then
                Dim message As NativeMessage = System.Runtime.InteropServices.Marshal.PtrToStructure(Of NativeMessage)(messagePointer)
                ' Observe completed clicks/drags once, after Word has received the
                ' matching down event. Double clicks also finish with WM_LBUTTONUP.
                Dim isPointer As System.Boolean = message.Message = &H202UI
                Dim isNavigation As System.Boolean = False
                If message.Message = &H100UI Then
                    Dim key As System.UInt64 = message.WParam.ToUInt64()
                    isNavigation = key = 8UL OrElse key = 9UL OrElse key = 27UL OrElse
                        (key >= 33UL AndAlso key <= 40UL) OrElse key = 46UL OrElse key = 117UL
                End If
                ' A character message is merely an activity signal. Its character
                ' value is never read, retained or logged by this observer.
                Dim isKeyboard As System.Boolean = isNavigation OrElse message.Message = &H102UI
                If isPointer OrElse isKeyboard Then
                    Dim windowHandle As System.IntPtr = GetAncestor(message.Hwnd, 2UI)
                    If IsWordWindow(windowHandle) Then
                        Dim generation As System.Int32 = System.Threading.Volatile.Read(_generation)
                        If isPointer Then
                            Dim point As NativePoint = message.Point
                            Enqueue(Sub()
                                        If generation <> System.Threading.Volatile.Read(_generation) OrElse _automation Is Nothing Then Return
                                        ' Resolve the element actually clicked, not the old keyboard
                                        ' focus and not the pointer position at command invocation.
                                        Try
                                            If GetAncestor(WindowFromPoint(point), 2UI) <> windowHandle Then
                                                RejectObservedInput(windowHandle, "click_surface_changed")
                                                Return
                                            End If
                                            Dim atPoint As NativeElement = Nothing
                                            Check(_automation.ElementFromPoint(point, atPoint))
                                            ReadFocus(atPoint, Observation.Pointer, windowHandle)
                                        Catch ex As System.Exception
                                            RejectObservedInput(windowHandle, "click_probe_failed")
                                            System.Diagnostics.Debug.WriteLine("Word comment click probe: " & ex.ToString())
                                        End Try
                                    End Sub)
                        ElseIf System.Threading.Interlocked.Exchange(_pendingKeyboardProbe, 1) = 0 Then
                            Enqueue(Sub()
                                        System.Threading.Interlocked.Exchange(_pendingKeyboardProbe, 0)
                                        If generation <> System.Threading.Volatile.Read(_generation) OrElse _automation Is Nothing Then Return
                                        _bodyNavigationPending = True
                                        Dim focused As NativeElement = Nothing
                                        Check(_automation.GetFocusedElement(focused))
                                        ReadFocus(focused, Observation.Keyboard, windowHandle)
                                    End Sub)
                        End If
                    End If
                End If
            End If
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("Word comment input observation: " & ex.ToString())
        End Try
        ' Always forward every input message, including Esc. Never swallow input.
        Return CallNextHookEx(_inputHook, code, removal, messagePointer)
    End Function

    Private Sub OnSelection(sender As NativeElement, eventId As System.Int32)
        If System.Threading.Volatile.Read(_stopping) <> 0 Then Return
        If System.Threading.Interlocked.Increment(_pendingSelectionEvents) > 64 Then
            System.Threading.Interlocked.Decrement(_pendingSelectionEvents)
            System.Threading.Interlocked.Increment(_generation)
            Return
        End If
        Dim generation As System.Int32 = System.Threading.Volatile.Read(_generation)
        Enqueue(Sub()
                    Try
                        If generation <> System.Threading.Volatile.Read(_generation) Then
                            _last = Evidence.Unknown("selection_events_invalidated")
                            _lastCard = Nothing
                            Return
                        End If
                        ' Removed-from-selection invalidates only this actual card,
                        ' not an unrelated card or the old document selection.
                        If eventId = 20011 Then
                            If SameElement(sender, _lastCard) AndAlso _commandSurfaceWindow <> _last.WindowHandle Then
                                _last = Evidence.Unknown("comment_deselected")
                                _lastCard = Nothing
                            End If
                            Return
                        End If
                        ReadFocus(sender, If(eventId = 20014, Observation.TextSelection, Observation.Selection))
                    Finally
                        System.Threading.Interlocked.Decrement(_pendingSelectionEvents)
                    End Try
                End Sub)
    End Sub

    Private Sub Enqueue(action As System.Action)
        If System.Threading.Volatile.Read(_stopping) <> 0 Then Return
        _queue.Enqueue(action)
        Try
            _wake.Set()
        Catch ex As System.ObjectDisposedException
            ' Host shutdown won the race. No work is performed after shutdown.
        End Try
    End Sub

    Private Sub WorkerMain()
        Dim subscribed As System.Boolean = False
        Try
            Dim automationType As System.Type = System.Type.GetTypeFromCLSID(
                New System.Guid("ff48dba4-60ef-4201-aa87-54103eef594e"), True)
            _automation = DirectCast(System.Activator.CreateInstance(automationType), NativeAutomation)
            Check(_automation.GetRawViewWalker(_walker))
            _handler = New FocusHandler(Me)
            Check(_automation.AddFocusChangedEventHandler(System.IntPtr.Zero, _handler))
            subscribed = True
        Catch ex As System.Exception
            _last = Evidence.Unknown("initialization_failed")
            System.Diagnostics.Debug.WriteLine("Word comment focus initialization: " & ex.ToString())
            _automation = Nothing
        End Try
        Try
            While System.Threading.Volatile.Read(_stopping) = 0
                _wake.WaitOne()
                Dim action As System.Action = Nothing
                While _queue.TryDequeue(action)
                    If System.Threading.Volatile.Read(_stopping) <> 0 Then Exit While
                    Try
                        action.Invoke()
                    Catch ex As System.Exception
                        RejectObservedInput(If(_last.WindowHandle <> System.IntPtr.Zero, _last.WindowHandle, _watchedWindow), "probe_failed")
                        System.Diagnostics.Debug.WriteLine("Word comment focus worker: " & ex.ToString())
                    End Try
                End While
            End While
        Finally
            ' Unsubscribe on the same MTA thread; never Join this thread on Office's STA.
            If subscribed AndAlso _automation IsNot Nothing Then
                Try
                    Check(_automation.RemoveFocusChangedEventHandler(_handler))
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("Word comment focus shutdown: " & ex.ToString())
                End Try
            End If
            RemoveSelectionSubscriptions()
            Dim discarded As System.Action = Nothing
            While _queue.TryDequeue(discarded)
            End While
            _last = Evidence.Unknown("stopped")
            _handler = Nothing
            _selectionHandler = Nothing
            _lastCard = Nothing
            _walker = Nothing
            _automation = Nothing
            _wake.Dispose()
        End Try
    End Sub

    Private Sub OnFocus(sender As NativeElement)
        If System.Threading.Volatile.Read(_stopping) <> 0 Then Return
        ' Retain the event's focused element, NOT a later GetFocusedElement result.
        ' Otherwise a fast comment -> ribbon click would lose the comment again.
        If System.Threading.Interlocked.Increment(_pendingFocusEvents) > 64 Then
            System.Threading.Interlocked.Decrement(_pendingFocusEvents)
            ' A focus burst must not leave an apparently valid old card. Queued
            ' events are invalidated, and the next command must obtain fresh proof.
            System.Threading.Interlocked.Increment(_generation)
            Return
        End If
        Dim generation As System.Int32 = System.Threading.Volatile.Read(_generation)
        Enqueue(Sub()
                    Try
                        If generation = System.Threading.Volatile.Read(_generation) Then
                            ReadFocus(sender, Observation.Focus)
                        Else
                            _last = Evidence.Unknown("focus_events_invalidated")
                            _commandSurfaceWindow = System.IntPtr.Zero
                        End If
                    Finally
                        System.Threading.Interlocked.Decrement(_pendingFocusEvents)
                    End Try
                End Sub)
    End Sub

    Private Sub ReadFocus(focused As NativeElement, source As Observation,
                          Optional expectedWindow As System.IntPtr = Nothing)
        If focused Is Nothing OrElse ReadInteger(focused, UIA_ProcessId) <> _processId Then
            ' Selection events from an unrelated provider must not become a target.
            If source = Observation.Selection OrElse source = Observation.TextSelection Then Return
            If source = Observation.Pointer Then _uncertainInputWindow = expectedWindow
            _last = Evidence.Unknown("outside_word_process")
            _lastCard = Nothing
            _commandSurfaceWindow = System.IntPtr.Zero
            Return
        End If
        Dim node As NativeElement = focused
        Dim focusPath As New System.Collections.Generic.List(Of NativeElement)()
        Dim card As NativeElement = Nothing
        Dim annotation As NativeElement = Nothing
        Dim nearestTreeItem As NativeElement = Nothing
        Dim contentSeen As System.Boolean = False
        Dim bodyFocusVerified As System.Boolean = False
        Dim commandSurfaceSeen As System.Boolean = False
        Dim rootWindow As System.IntPtr = System.IntPtr.Zero
        Dim watch As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
        For depth As System.Int32 = 0 To MaxAncestors - 1
            If node Is Nothing OrElse watch.ElapsedMilliseconds > 600 Then Exit For
            If ReadInteger(node, UIA_ProcessId) <> _processId Then Exit For
            focusPath.Add(node)
            Dim controlType As System.Int32 = ReadInteger(node, UIA_ControlType)
            Dim className As System.String = ReadString(node, UIA_ClassName)
            Dim automationId As System.String = ReadString(node, UIA_AutomationId)
            If controlType = UIA_TreeItem AndAlso nearestTreeItem Is Nothing Then nearestTreeItem = node
            If annotation Is Nothing AndAlso ReadInteger(node, UIA_AnnotationType) = UIA_CommentAnnotation Then annotation = node
            If card Is Nothing AndAlso ReadInteger(node, UIA_AnnotationType) = UIA_CommentAnnotation Then
                ' A focused reply tree-item is more specific than its thread's
                ' annotation. Do not silently promote it to the thread's body.
                card = If(nearestTreeItem, node)
            End If
            ' Resolved comments use a custom annotation type. Only consider that
            ' annotation on the focused tree-item path, never arbitrary annotations.
            If card Is Nothing AndAlso nearestTreeItem IsNot Nothing AndAlso
               ReadBoolean(node, UIA_IsAnnotationPatternAvailable) Then card = nearestTreeItem
            If System.String.Equals(className, "_WwG", System.StringComparison.Ordinal) OrElse
               automationId.StartsWith("UIA_AutomationId_Word_Content", System.StringComparison.Ordinal) Then
                contentSeen = True
                ' Merely being a descendant of the document is not proof of body
                ' focus: modern comments are embedded there too.
                If controlType = UIA_Document Then bodyFocusVerified = True
            End If
            If controlType = UIA_ToolBar OrElse controlType = UIA_Menu OrElse controlType = UIA_MenuBar OrElse
               System.String.Equals(className, "NetUIRibbon", System.StringComparison.Ordinal) Then
                commandSurfaceSeen = True
            End If
            Dim nativeHandle As System.Int32 = ReadInteger(node, UIA_NativeWindowHandle)
            If nativeHandle <> 0 Then
                Dim root As System.IntPtr = GetAncestor(New System.IntPtr(nativeHandle), 2UI)
                If IsWordWindow(root) Then rootWindow = root
            End If
            Dim parent As NativeElement = Nothing
            Check(_walker.GetParentElement(node, parent))
            node = parent
        Next
        If expectedWindow <> System.IntPtr.Zero AndAlso rootWindow <> expectedWindow Then
            If source = Observation.Pointer Then _uncertainInputWindow = expectedWindow
            _last = Evidence.Unknown("input_window_mismatch")
            _lastCard = Nothing
            Return
        End If
        If rootWindow = System.IntPtr.Zero Then
            If source = Observation.Selection OrElse source = Observation.TextSelection Then Return
            If source = Observation.Pointer Then _uncertainInputWindow = expectedWindow
            _last = Evidence.Unknown("unverified_word_window")
            _lastCard = Nothing
            _commandSurfaceWindow = System.IntPtr.Zero
            Return
        End If
        If (source = Observation.Selection OrElse source = Observation.TextSelection) AndAlso
           rootWindow <> _watchedWindow Then Return
        If card Is Nothing AndAlso contentSeen Then card = nearestTreeItem
        If card IsNot Nothing Then
            _uncertainInputWindow = System.IntPtr.Zero
            If source = Observation.Pointer OrElse source = Observation.Selection OrElse source = Observation.Keyboard Then
                _commandSurfaceWindow = System.IntPtr.Zero
            End If
            If source = Observation.Pointer OrElse source = Observation.Selection OrElse
               source = Observation.TextSelection OrElse source = Observation.Capture Then _bodyNavigationPending = False
            Dim texts As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim selectedTexts As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim author As System.String = ReadString(card, UIA_AnnotationAuthor)
            Dim visited As System.Int32 = 0
            ' A clicked label, selected read-only text and an editor are three
            ' different focus shapes. Identity comes from the complete owning card,
            ' not just the focused child's text and not the selected fragment.
            If annotation IsNot Nothing AndAlso SameElement(annotation, card) Then
                AddText(texts, ReadString(annotation, UIA_Name))
                author = ReadString(annotation, UIA_AnnotationAuthor)
            End If
            CollectTexts(card, card, texts, visited, watch)
            For Each pathNode As NativeElement In focusPath
                If watch.ElapsedMilliseconds > 600 Then Exit For
                AddProviderSelection(pathNode, selectedTexts)
                If SameElement(pathNode, card) Then Exit For
            Next
            Dim selectedText As System.String = System.String.Empty
            If selectedTexts.Count = 1 Then
                For Each value As System.String In selectedTexts
                    selectedText = value
                Next
            End If
            If selectedText.Length = 0 AndAlso (source = Observation.Refresh OrElse source = Observation.Capture) AndAlso
               SameElement(card, _lastCard) Then selectedText = _last.SelectedText
            _lastCard = card
            _last = New Evidence(rootWindow, True, True, texts, author,
                "source=" & source.ToString() & "; nodes=" & visited.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; textFields=" & texts.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; inputObserver=" & (_inputHook <> System.IntPtr.Zero).ToString(), selectedText:=selectedText)
        ElseIf commandSurfaceSeen Then
            If _last.WindowHandle <> rootWindow Then
                _last = Evidence.Unknown("command_without_prior_target")
                _lastCard = Nothing
            End If
            _bodyNavigationPending = False
            _commandSurfaceWindow = rootWindow
        ElseIf bodyFocusVerified Then
            If _uncertainInputWindow = rootWindow AndAlso source <> Observation.Pointer AndAlso source <> Observation.Keyboard Then
                RejectObservedInput(rootWindow, "body_after_unverified_input")
                Return
            End If
            If source = Observation.Pointer OrElse source = Observation.Keyboard Then _uncertainInputWindow = System.IntPtr.Zero
            ' Clicking a collapsed card need not change keyboard focus at all.
            ' Likewise Ribbon callbacks may restore the old body range. Neither is
            ' a NEW user selection. An actual body click/navigation clears the card.
            Dim keepCard As System.Boolean = _last.IsComment AndAlso _last.WindowHandle = rootWindow AndAlso
                source <> Observation.Pointer AndAlso source <> Observation.Keyboard AndAlso
                Not _bodyNavigationPending AndAlso
                (source <> Observation.BodyProjection OrElse _inputHook <> System.IntPtr.Zero)
            If Not keepCard Then
                _last = New Evidence(rootWindow, True, False, Nothing, Nothing, "body; source=" & source.ToString())
                _lastCard = Nothing
                _bodyNavigationPending = False
                _commandSurfaceWindow = System.IntPtr.Zero
            End If
        ElseIf source <> Observation.Selection AndAlso source <> Observation.TextSelection Then
            If source = Observation.Pointer Then _uncertainInputWindow = rootWindow
            _last = Evidence.Unknown("other_word_surface; source=" & source.ToString())
            _lastCard = Nothing
            _bodyNavigationPending = False
            _commandSurfaceWindow = System.IntPtr.Zero
        End If
    End Sub

    Private Sub RejectObservedInput(windowHandle As System.IntPtr, reason As System.String)
        _uncertainInputWindow = windowHandle
        _last = Evidence.Unknown(reason)
        _lastCard = Nothing
        _commandSurfaceWindow = System.IntPtr.Zero
    End Sub

    Private Function SameElement(left As NativeElement, right As NativeElement) As System.Boolean
        If left Is Nothing OrElse right Is Nothing OrElse _automation Is Nothing Then Return False
        Dim same As System.Boolean = False
        Return _automation.CompareElements(left, right, same) >= 0 AndAlso same
    End Function

    Private Sub CollectTexts(root As NativeElement, node As NativeElement,
                             texts As System.Collections.Generic.HashSet(Of System.String),
                             ByRef visited As System.Int32, watch As System.Diagnostics.Stopwatch)
        If node Is Nothing OrElse visited >= MaxNodes OrElse watch.ElapsedMilliseconds > 600 Then Return
        If ReadInteger(node, UIA_ProcessId) <> _processId Then Return
        Dim controlType As System.Int32 = ReadInteger(node, UIA_ControlType)
        ' Do not collect other replies/threads below the one that received focus.
        If Not SameElement(node, root) AndAlso (controlType = UIA_TreeItem OrElse
            ReadInteger(node, UIA_AnnotationType) = UIA_CommentAnnotation) Then Return
        visited += 1
        If ReadInteger(node, UIA_AnnotationType) = UIA_CommentAnnotation Then
            AddText(texts, ReadString(node, UIA_Name))
        End If
        ' The structural-card fallback inspects body controls, not labels such as
        ' author names/dates that might happen to equal another comment's text.
        If controlType = UIA_Edit OrElse controlType = UIA_Document OrElse controlType = UIA_Custom OrElse controlType = UIA_Text OrElse SameElement(node, root) Then
            AddProviderText(node, texts)
        End If
        Dim child As NativeElement = Nothing
        Check(_walker.GetFirstChildElement(node, child))
        While child IsNot Nothing AndAlso visited < MaxNodes AndAlso watch.ElapsedMilliseconds <= 600
            CollectTexts(root, child, texts, visited, watch)
            Dim sibling As NativeElement = Nothing
            Check(_walker.GetNextSiblingElement(child, sibling))
            child = sibling
        End While
    End Sub

    Private Shared Sub AddProviderText(node As NativeElement,
                                       texts As System.Collections.Generic.HashSet(Of System.String))
        Try
            ' Read-only annotations can expose ValuePattern without being Edit
            ' controls. A provider's value is content; an arbitrary UI label is not.
            AddText(texts, ReadString(node, UIA_ValueValue))
            Dim patternObject As System.Object = Nothing
            If node.GetCurrentPattern(UIA_TextPattern, patternObject) < 0 OrElse patternObject Is Nothing Then Return
            Dim pattern As NativeTextPattern = TryCast(patternObject, NativeTextPattern)
            If pattern Is Nothing Then Return
            Dim textRange As NativeTextRange = Nothing
            Check(pattern.GetDocumentRange(textRange))
            If textRange IsNot Nothing Then
                Dim text As System.String = Nothing
                Check(textRange.GetText(MaxTextLength + 1, text))
                AddText(texts, text)
            End If
        Catch ex As System.Runtime.InteropServices.COMException
            ' Incomplete/unavailable UI evidence is never replaced by label guesses.
        End Try
    End Sub

    Private Shared Sub AddProviderSelection(node As NativeElement,
                                            selectedTexts As System.Collections.Generic.HashSet(Of System.String))
        Try
            Dim patternObject As System.Object = Nothing
            If node.GetCurrentPattern(UIA_TextPattern, patternObject) < 0 Then Return
            Dim pattern As NativeTextPattern = TryCast(patternObject, NativeTextPattern)
            If pattern Is Nothing Then Return
            Dim ranges As NativeTextRangeArray = Nothing
            If pattern.GetSelection(ranges) < 0 OrElse ranges Is Nothing Then Return
            Dim count As System.Int32 = 0
            Check(ranges.GetLength(count))
            If count <> 1 Then Return
            Dim range As NativeTextRange = Nothing
            Check(ranges.GetElement(0, range))
            If range Is Nothing Then Return
            Dim text As System.String = Nothing
            Check(range.GetText(MaxTextLength + 1, text))
            If Not System.String.IsNullOrWhiteSpace(text) AndAlso text.Length <= MaxTextLength Then selectedTexts.Add(text)
        Catch unavailable As System.Runtime.InteropServices.COMException
            ' A selected read-only card need not expose a text selection.
        End Try
    End Sub

    Private Shared Sub AddText(texts As System.Collections.Generic.HashSet(Of System.String), value As System.String)
        If value Is Nothing OrElse value.Length > MaxTextLength Then Return
        Dim normalized As System.String = NormalizeText(value)
        If normalized.Length > 0 Then texts.Add(normalized)
    End Sub

    Public Shared Function NormalizeText(value As System.String) As System.String
        If System.String.IsNullOrEmpty(value) Then Return System.String.Empty
        ' Layout whitespace and Word's terminal reference/cell markers are not
        ' identity. Never case-fold, substring-match, or shorten comment content.
        Dim first As System.Int32 = 0
        Dim last As System.Int32 = value.Length - 1
        While first <= last AndAlso IsBoundaryCharacter(value(first))
            first += 1
        End While
        While last >= first AndAlso IsBoundaryCharacter(value(last))
            last -= 1
        End While
        If last < first Then Return System.String.Empty
        Return System.Text.RegularExpressions.Regex.Replace(value.Substring(first, last - first + 1), "\s+", " ").Trim()
    End Function

    Private Shared Function IsBoundaryCharacter(value As System.Char) As System.Boolean
        Dim code As System.Int32 = System.Convert.ToInt32(value)
        Return System.Char.IsWhiteSpace(value) OrElse code = 0 OrElse code = 5 OrElse code = 7
    End Function

    Public Shared Function Matches(evidence As Evidence, text As System.String, author As System.String) As System.Boolean
        If evidence Is Nothing OrElse Not evidence.KnownFocus OrElse Not evidence.IsComment Then Return False
        Dim normalized As System.String = NormalizeText(text)
        If normalized.Length = 0 Then Return False
        If Not System.String.IsNullOrWhiteSpace(evidence.Author) AndAlso
           Not System.String.Equals(NormalizeText(evidence.Author), NormalizeText(author), System.StringComparison.Ordinal) Then Return False
        For Each field As System.String In evidence.Texts
            If System.String.Equals(field, normalized, System.StringComparison.Ordinal) Then Return True
        Next
        Return False
    End Function

    Private Function IsWordWindow(hwnd As System.IntPtr) As System.Boolean
        If hwnd = System.IntPtr.Zero Then Return False
        Dim pid As System.UInt32 = 0UI
        GetWindowThreadProcessId(hwnd, pid)
        If pid <> CUInt(_processId) Then Return False
        Dim name As New System.Text.StringBuilder(64)
        GetClassName(hwnd, name, name.Capacity)
        Return System.String.Equals(name.ToString(), "OpusApp", System.StringComparison.Ordinal)
    End Function

    Private Shared Function ReadProperty(element As NativeElement, propertyId As System.Int32) As System.Object
        Dim value As System.Object = Nothing
        If element.GetCurrentPropertyValue(propertyId, value) < 0 Then Return Nothing
        Return value
    End Function

    Private Shared Function ReadInteger(element As NativeElement, propertyId As System.Int32) As System.Int32
        Dim value As System.Object = ReadProperty(element, propertyId)
        If TypeOf value Is System.Int32 Then Return DirectCast(value, System.Int32)
        Return 0
    End Function

    Private Shared Function ReadBoolean(element As NativeElement, propertyId As System.Int32) As System.Boolean
        Dim value As System.Object = ReadProperty(element, propertyId)
        Return TypeOf value Is System.Boolean AndAlso DirectCast(value, System.Boolean)
    End Function

    Private Shared Function ReadString(element As NativeElement, propertyId As System.Int32) As System.String
        Return If(TryCast(ReadProperty(element, propertyId), System.String), System.String.Empty)
    End Function

    Private Shared Sub Check(hresult As System.Int32)
        System.Runtime.InteropServices.Marshal.ThrowExceptionForHR(hresult)
    End Sub

    Public Sub Dispose() Implements System.IDisposable.Dispose
        If System.Threading.Interlocked.Exchange(_stopping, 1) = 0 Then
            If _inputHook <> System.IntPtr.Zero Then
                If UnhookWindowsHookEx(_inputHook) Then
                    Dim removedHandler As MessageHook = Nothing
                    LiveMessageHandlers.TryRemove(_inputHook, removedHandler)
                    _inputHook = System.IntPtr.Zero
                Else
                    ' Keep the delegate rooted. The stopped handler is pass-through
                    ' only, so a failed unhook can never call collected managed code.
                    System.Diagnostics.Debug.WriteLine("Word comment input observer shutdown: " &
                        System.Runtime.InteropServices.Marshal.GetLastWin32Error().ToString(System.Globalization.CultureInfo.InvariantCulture))
                End If
            End If
            Try
                _wake.Set()
            Catch ex As System.ObjectDisposedException
            End Try
        End If
    End Sub

    <System.Runtime.InteropServices.DllImport("user32.dll", ExactSpelling:=True)>
    Private Shared Function GetAncestor(hwnd As System.IntPtr, flags As System.UInt32) As System.IntPtr
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll", ExactSpelling:=True)>
    Private Shared Function GetWindowThreadProcessId(hwnd As System.IntPtr, ByRef processId As System.UInt32) As System.UInt32
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, EntryPoint:="GetClassNameW")>
    Private Shared Function GetClassName(hwnd As System.IntPtr, className As System.Text.StringBuilder, maximum As System.Int32) As System.Int32
    End Function

    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
    Private Structure NativePoint
        Public X As System.Int32
        Public Y As System.Int32
    End Structure

    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
    Private Structure NativeMessage
        Public Hwnd As System.IntPtr
        Public Message As System.UInt32
        Public WParam As System.UIntPtr
        Public LParam As System.IntPtr
        Public Time As System.UInt32
        Public Point As NativePoint
        Public PrivateValue As System.UInt32
    End Structure

    <System.Runtime.InteropServices.UnmanagedFunctionPointer(System.Runtime.InteropServices.CallingConvention.Winapi)>
    Private Delegate Function MessageHook(code As System.Int32, removal As System.IntPtr, message As System.IntPtr) As System.IntPtr

    <System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SetWindowsHookExW", SetLastError:=True)>
    Private Shared Function SetWindowsHookEx(kind As System.Int32, handler As MessageHook,
                                              moduleHandle As System.IntPtr, threadId As System.UInt32) As System.IntPtr
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll", ExactSpelling:=True)>
    Private Shared Function CallNextHookEx(hook As System.IntPtr, code As System.Int32,
                                            removal As System.IntPtr, message As System.IntPtr) As System.IntPtr
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll", ExactSpelling:=True, SetLastError:=True)>
    Private Shared Function UnhookWindowsHookEx(hook As System.IntPtr) As System.Boolean
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll", ExactSpelling:=True)>
    Private Shared Function WindowFromPoint(point As NativePoint) As System.IntPtr
    End Function

    <System.Runtime.InteropServices.DllImport("kernel32.dll", ExactSpelling:=True)>
    Private Shared Function GetCurrentThreadId() As System.UInt32
    End Function

    <System.Runtime.InteropServices.ComVisible(True),
     System.Runtime.InteropServices.Guid("146c3c17-f12e-4e22-8c27-f894b9b79c69"),
     System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface NativeSelectionHandler
        Sub HandleAutomationEvent(sender As NativeElement, eventId As System.Int32)
    End Interface

    <System.Runtime.InteropServices.ComVisible(True),
     System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)>
    Public NotInheritable Class SelectionHandler
        Implements NativeSelectionHandler
        Private ReadOnly _owner As WordCommentSelectionTracker
        Friend Sub New(owner As WordCommentSelectionTracker)
            _owner = owner
        End Sub
        Public Sub HandleAutomationEvent(sender As NativeElement, eventId As System.Int32) Implements NativeSelectionHandler.HandleAutomationEvent
            Try
                _owner.OnSelection(sender, eventId)
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Word comment selection event: " & ex.ToString())
            End Try
        End Sub
    End Class

    <System.Runtime.InteropServices.ComVisible(True),
     System.Runtime.InteropServices.Guid("c270f6b5-5c69-4290-9745-7a7f97169468"),
     System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface NativeFocusHandler
        Sub HandleFocusChangedEvent(sender As NativeElement)
    End Interface

    <System.Runtime.InteropServices.ComVisible(True),
     System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)>
    Public NotInheritable Class FocusHandler
        Implements NativeFocusHandler
        Private ReadOnly _owner As WordCommentSelectionTracker
        Friend Sub New(owner As WordCommentSelectionTracker)
            _owner = owner
        End Sub
        Public Sub HandleFocusChangedEvent(sender As NativeElement) Implements NativeFocusHandler.HandleFocusChangedEvent
            Try
                _owner.OnFocus(sender)
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Word comment focus event: " & ex.ToString())
            End Try
        End Sub
    End Class

    ' Prefix-only COM interfaces: preserve the SDK vtable order. Unused slots are
    ' deliberately parameterless and never called; no unsafe pointer arithmetic.
    <System.Runtime.InteropServices.ComImport,
     System.Runtime.InteropServices.Guid("30cbe57d-d9d0-452a-ab13-7ac5ac4825ee"),
     System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface NativeAutomation
        <System.Runtime.InteropServices.PreserveSig>
        Function CompareElements(left As NativeElement, right As NativeElement,
            <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)> ByRef same As System.Boolean) As System.Int32
        Sub UnusedCompareRuntimeIds()
        Sub UnusedGetRootElement()
        <System.Runtime.InteropServices.PreserveSig>
        Function ElementFromHandle(hwnd As System.IntPtr, ByRef element As NativeElement) As System.Int32
        <System.Runtime.InteropServices.PreserveSig>
        Function ElementFromPoint(point As NativePoint, ByRef element As NativeElement) As System.Int32
        <System.Runtime.InteropServices.PreserveSig>
        Function GetFocusedElement(ByRef element As NativeElement) As System.Int32
        Sub UnusedGetRootElementBuildCache()
        Sub UnusedElementFromHandleBuildCache()
        Sub UnusedElementFromPointBuildCache()
        Sub UnusedGetFocusedElementBuildCache()
        Sub UnusedCreateTreeWalker()
        Sub UnusedGetControlViewWalker()
        Sub UnusedGetContentViewWalker()
        <System.Runtime.InteropServices.PreserveSig>
        Function GetRawViewWalker(ByRef walker As NativeTreeWalker) As System.Int32
        Sub UnusedGetRawViewCondition()
        Sub UnusedGetControlViewCondition()
        Sub UnusedGetContentViewCondition()
        Sub UnusedCreateCacheRequest()
        Sub UnusedCreateTrueCondition()
        Sub UnusedCreateFalseCondition()
        Sub UnusedCreatePropertyCondition()
        Sub UnusedCreatePropertyConditionEx()
        Sub UnusedCreateAndCondition()
        Sub UnusedCreateAndConditionFromArray()
        Sub UnusedCreateAndConditionFromNativeArray()
        Sub UnusedCreateOrCondition()
        Sub UnusedCreateOrConditionFromArray()
        Sub UnusedCreateOrConditionFromNativeArray()
        Sub UnusedCreateNotCondition()
        <System.Runtime.InteropServices.PreserveSig>
        Function AddAutomationEventHandler(eventId As System.Int32, element As NativeElement,
            scope As System.Int32, cache As System.IntPtr, handler As NativeSelectionHandler) As System.Int32
        <System.Runtime.InteropServices.PreserveSig>
        Function RemoveAutomationEventHandler(eventId As System.Int32, element As NativeElement,
            handler As NativeSelectionHandler) As System.Int32
        Sub UnusedAddPropertyChangedEventHandlerNativeArray()
        Sub UnusedAddPropertyChangedEventHandler()
        Sub UnusedRemovePropertyChangedEventHandler()
        Sub UnusedAddStructureChangedEventHandler()
        Sub UnusedRemoveStructureChangedEventHandler()
        <System.Runtime.InteropServices.PreserveSig>
        Function AddFocusChangedEventHandler(cache As System.IntPtr, handler As NativeFocusHandler) As System.Int32
        <System.Runtime.InteropServices.PreserveSig>
        Function RemoveFocusChangedEventHandler(handler As NativeFocusHandler) As System.Int32
    End Interface

    <System.Runtime.InteropServices.ComImport, System.Runtime.InteropServices.ComVisible(True),
     System.Runtime.InteropServices.Guid("d22108aa-8ac5-49a5-837b-37bbb3d7591e"),
     System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface NativeElement
        Sub UnusedSetFocus()
        Sub UnusedGetRuntimeId()
        Sub UnusedFindFirst()
        Sub UnusedFindAll()
        Sub UnusedFindFirstBuildCache()
        Sub UnusedFindAllBuildCache()
        Sub UnusedBuildUpdatedCache()
        <System.Runtime.InteropServices.PreserveSig>
        Function GetCurrentPropertyValue(propertyId As System.Int32,
            <System.Runtime.InteropServices.Out, System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Struct)> ByRef value As System.Object) As System.Int32
        Sub UnusedGetCurrentPropertyValueEx()
        Sub UnusedGetCachedPropertyValue()
        Sub UnusedGetCachedPropertyValueEx()
        Sub UnusedGetCurrentPatternAs()
        Sub UnusedGetCachedPatternAs()
        <System.Runtime.InteropServices.PreserveSig>
        Function GetCurrentPattern(patternId As System.Int32,
            <System.Runtime.InteropServices.Out, System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.IUnknown)> ByRef pattern As System.Object) As System.Int32
    End Interface

    <System.Runtime.InteropServices.ComImport,
     System.Runtime.InteropServices.Guid("4042c624-389c-4afc-a630-9df854a541fc"),
     System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface NativeTreeWalker
        <System.Runtime.InteropServices.PreserveSig>
        Function GetParentElement(element As NativeElement, ByRef parent As NativeElement) As System.Int32
        <System.Runtime.InteropServices.PreserveSig>
        Function GetFirstChildElement(element As NativeElement, ByRef child As NativeElement) As System.Int32
        Sub UnusedGetLastChildElement()
        <System.Runtime.InteropServices.PreserveSig>
        Function GetNextSiblingElement(element As NativeElement, ByRef sibling As NativeElement) As System.Int32
    End Interface

    <System.Runtime.InteropServices.ComImport,
     System.Runtime.InteropServices.Guid("32eba289-3583-42c9-9c59-3b6d9a1e9b6a"),
     System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface NativeTextPattern
        Sub UnusedRangeFromPoint()
        Sub UnusedRangeFromChild()
        <System.Runtime.InteropServices.PreserveSig>
        Function GetSelection(ByRef ranges As NativeTextRangeArray) As System.Int32
        Sub UnusedGetVisibleRanges()
        <System.Runtime.InteropServices.PreserveSig>
        Function GetDocumentRange(ByRef range As NativeTextRange) As System.Int32
    End Interface

    <System.Runtime.InteropServices.ComImport,
     System.Runtime.InteropServices.Guid("ce4ae76a-e717-4c98-81ea-47371d028eb6"),
     System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface NativeTextRangeArray
        <System.Runtime.InteropServices.PreserveSig>
        Function GetLength(ByRef length As System.Int32) As System.Int32
        <System.Runtime.InteropServices.PreserveSig>
        Function GetElement(index As System.Int32, ByRef range As NativeTextRange) As System.Int32
    End Interface

    <System.Runtime.InteropServices.ComImport,
     System.Runtime.InteropServices.Guid("a543cc6a-f4ae-494b-8239-c814481187a8"),
     System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface NativeTextRange
        Sub UnusedClone()
        Sub UnusedCompare()
        Sub UnusedCompareEndpoints()
        Sub UnusedExpandToEnclosingUnit()
        Sub UnusedFindAttribute()
        Sub UnusedFindText()
        Sub UnusedGetAttributeValue()
        Sub UnusedGetBoundingRectangles()
        Sub UnusedGetEnclosingElement()
        <System.Runtime.InteropServices.PreserveSig>
        Function GetText(maxLength As System.Int32,
            <System.Runtime.InteropServices.Out, System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.BStr)> ByRef text As System.String) As System.Int32
    End Interface
End Class
