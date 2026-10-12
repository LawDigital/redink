' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: AssistantForm.vb
' Purpose: Ask Inky desktop functionality in the modular Windows client.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). Native title bar; the entire client area is local HTML/CSS.
Friend NotInheritable Class AssistantForm
    Inherits WebDialog
    Private ReadOnly context As AssistantContext
    Private selection As String = ""
    Private question As String = ""
    Private response As String = ""
    Private statusKey As String = "desktop.ready"
    Private sending As Boolean
    Private includeHistory As Boolean = True
    Private ReadOnly turns As New System.Collections.Generic.List(Of ConversationTurn)()
    Private displayedTurnIndex As Integer = -1
    Private NotInheritable Class ConversationTurn
        Public Property Question As String = ""
        Public Property Answer As String = ""
    End Class
    Private settingsOpen As Boolean
    Private exiting As Boolean
    Private manualResizeInProgress As Boolean
    Private suppressAutomaticGrowth As Boolean
    Private resizeStartClientHeight As Integer

    Protected Overrides Sub OnResizeBegin(e As System.EventArgs)
        manualResizeInProgress = True
        resizeStartClientHeight = Me.ClientSize.Height
        MyBase.OnResizeBegin(e)
    End Sub

    Protected Overrides Sub OnResizeEnd(e As System.EventArgs)
        MyBase.OnResizeEnd(e)
        If manualResizeInProgress AndAlso Me.ClientSize.Height < resizeStartClientHeight Then
            suppressAutomaticGrowth = True
        End If
        manualResizeInProgress = False
        If Me.Visible AndAlso Me.WindowState = System.Windows.Forms.FormWindowState.Normal Then
            context.RememberAssistantPosition(Me.Location)
        End If
    End Sub

    Public Sub New(parent As AssistantContext)
        MyBase.New("assistant.html", "Ask Inky", New System.Drawing.Size(600, 710))
        context = parent
        Dim available As System.Drawing.Rectangle = System.Windows.Forms.Screen.FromPoint(System.Windows.Forms.Cursor.Position).WorkingArea
        Dim estimatedNonClientHeight As Integer = Me.Height - Me.ClientSize.Height
        Dim maximumClientHeight As Integer = System.Math.Max(390, available.Height - estimatedNonClientHeight - 24)
        Dim minimumClientHeight As Integer = System.Math.Min(550, maximumClientHeight)
        Me.MinimumSize = New System.Drawing.Size(470, minimumClientHeight + estimatedNonClientHeight)
        Me.ClientSize = New System.Drawing.Size(600, System.Math.Min(710, maximumClientHeight))
        WindowPlacement.RestoreAssistant(Me, parent.Settings())
    End Sub

    Protected Overrides ReadOnly Property Language As String
        Get
            Return context.Settings().UiLanguage
        End Get
    End Property

    Protected Overrides Async Sub PageReady()
        SendView(New With {.[type] = "init", .strings = UiStrings.GetBundle(Language), .language = Language,
                           .darkMode = context.Settings().DarkMode, .fontSize = context.Settings().FontSize,
                           .selectedText = selection, .question = question, .html = ResponseClipboard.Render(response),
                           .turnIndex = displayedTurnIndex, .turnCount = turns.Count, .turnQuestion = DisplayedQuestion(),
                           .includeHistory = includeHistory, .hasAnswer = response.Length > 0,
                           .busy = sending, .locked = settingsOpen, .status = UiStrings.GetText(statusKey, Language),
                           .shortcut = HotkeyService.Format(context.Settings().HotkeyModifiers, context.Settings().HotkeyKey)})
        Try
            Dim available As Boolean = Await HelperClient.SupportsOcrAsync(context.Settings(), Lifetime.Token)
            SendView(New With {.[type] = "ocrCapability", .available = available, .shortcut = HotkeyService.Format(context.Settings().OcrHotkeyModifiers, context.Settings().OcrHotkeyKey)})
        Catch ex As System.Exception
            ClientLog.Write("ocr.capability.probe.failed", ex)
            SendView(New With {.[type] = "ocrCapability", .available = False})
        End Try
    End Sub

    Friend Async Sub ApplySettings()
        SendView(New With {.[type] = "preferences", .strings = UiStrings.GetBundle(Language), .language = Language,
                           .darkMode = context.Settings().DarkMode, .fontSize = context.Settings().FontSize,
                           .shortcut = HotkeyService.Format(context.Settings().HotkeyModifiers, context.Settings().HotkeyKey),
                           .status = UiStrings.GetText(statusKey, Language)})
        Try
            Dim available As Boolean = Await HelperClient.SupportsOcrAsync(context.Settings(), Lifetime.Token)
            SendView(New With {.[type] = "ocrCapability", .available = available, .shortcut = HotkeyService.Format(context.Settings().OcrHotkeyModifiers, context.Settings().OcrHotkeyKey)})
        Catch ex As System.Exception
            ClientLog.Write("ocr.capability.probe.failed", ex)
        End Try
    End Sub

    Friend Sub LoadSelection(value As String)
        If sending Then Return
        Dim incoming As String = If(value, "")
        If Not String.Equals(selection, incoming, System.StringComparison.Ordinal) Then
            selection = incoming
            suppressAutomaticGrowth = False
            turns.Clear()
            displayedTurnIndex = -1
            includeHistory = True
            response = ""
            question = ""
            SendView(New With {.[type] = "newConversation", .selectedText = selection, .includeHistory = includeHistory})
        End If
        statusKey = If(selection.Length > 0, "desktop.captured", "desktop.noSelection")
        SendView(New With {.[type] = "selection", .selectedText = selection, .status = UiStrings.GetText(statusKey, Language)})
    End Sub

    Friend Sub SetSettingsOpen(value As Boolean)
        settingsOpen = value
        SendView(New With {.[type] = "locked", .locked = value})
    End Sub

    Friend Sub FocusQuestion()
        SendView(New With {.[type] = "focus"})
    End Sub

    Protected Overrides Async Function HandleMessageAsync(message As System.Text.Json.JsonElement) As System.Threading.Tasks.Task
        Select Case HelperClient.TextValue(message, "action")
            Case "preferredHeight"
                Dim heightValue As System.Text.Json.JsonElement
                If message.TryGetProperty("height", heightValue) AndAlso heightValue.ValueKind = System.Text.Json.JsonValueKind.Number Then
                    Dim desiredHeight As Integer
                    If heightValue.TryGetInt32(desiredHeight) AndAlso desiredHeight > 0 Then
                        Dim workArea As System.Drawing.Rectangle = System.Windows.Forms.Screen.FromControl(Me).WorkingArea
                        Dim nonClientHeight As Integer = Me.Height - Me.ClientSize.Height
                        Dim maximumClientHeight As Integer = System.Math.Max(Me.MinimumSize.Height - nonClientHeight, workArea.Height - nonClientHeight - 24)
                        Dim targetHeight As Integer = System.Math.Min(System.Math.Max(Me.MinimumSize.Height - nonClientHeight, desiredHeight), maximumClientHeight)
                        If Not manualResizeInProgress AndAlso Not suppressAutomaticGrowth AndAlso targetHeight > Me.ClientSize.Height Then
                            Dim newTop As Integer = System.Math.Min(Me.Top, workArea.Bottom - nonClientHeight - targetHeight - 12)
                            Me.ClientSize = New System.Drawing.Size(Me.ClientSize.Width, targetHeight)
                            Me.Top = System.Math.Max(workArea.Top + 6, newTop)
                        End If
                    End If
                End If
            Case "questionChanged"
                Dim value As String = HelperClient.TextValue(message, "question")
                If value.Length <= 200000 AndAlso Not sending Then question = value
            Case "settings"
                If Not sending Then context.ShowSettings()
            Case "historyChanged"
                If Not sending AndAlso Not settingsOpen Then
                    Dim checkedValue As System.Text.Json.JsonElement
                    If message.TryGetProperty("includeHistory", checkedValue) Then
                        If checkedValue.ValueKind = System.Text.Json.JsonValueKind.True Then includeHistory = True
                        If checkedValue.ValueKind = System.Text.Json.JsonValueKind.False Then includeHistory = False
                    End If
                End If
            Case "navigateAnswer"
                If Not sending AndAlso Not settingsOpen AndAlso turns.Count > 0 Then
                    Dim direction As String = HelperClient.TextValue(message, "direction")
                    If direction = "previous" Then
                        ShowTurn(System.Math.Max(0, displayedTurnIndex - 1))
                    ElseIf direction = "next" Then
                        ShowTurn(System.Math.Min(turns.Count - 1, displayedTurnIndex + 1))
                    End If
                End If
            Case "deleteAnswer"
                If Not sending AndAlso Not settingsOpen AndAlso displayedTurnIndex >= 0 AndAlso displayedTurnIndex < turns.Count Then
                    turns.RemoveAt(displayedTurnIndex)
                    ShowTurn(System.Math.Min(displayedTurnIndex, turns.Count - 1))
                End If
            Case "clearAnswers"
                If Not sending AndAlso Not settingsOpen Then
                    turns.Clear()
                    ShowTurn(-1)
                End If
            Case "send"
                Await SendAsync(HelperClient.TextValue(message, "question"))
            Case "copySource"
                If sending OrElse settingsOpen OrElse System.String.IsNullOrEmpty(selection) Then Return
                Try
                    System.Windows.Forms.Clipboard.SetText(selection, System.Windows.Forms.TextDataFormat.UnicodeText)
                    statusKey = "desktop.copied"
                    SendView(New With {.[type] = "status", .status = UiStrings.GetText(statusKey, Language)})
                Catch ex As System.Exception
                    ClientLog.Write("clipboard.copy-source.failed", ex)
                    SendView(New With {.[type] = "error", .message = UiStrings.GetText("desktop.copyFailed", Language)})
                End Try
            Case "copy", "copySelection"
                If sending OrElse settingsOpen OrElse response.Length = 0 Then Return
                Try
                    Dim json As String = Await ReadCopyPayloadAsync(HelperClient.TextValue(message, "action") = "copySelection")
                    Using document As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(json)
                        Dim root As System.Text.Json.JsonElement = document.RootElement
                        Dim html As String = HelperClient.TextValue(root, "html")
                        Dim plain As String = HelperClient.TextValue(root, "text")
                        If html.Length = 0 OrElse html.Length > 16000000 OrElse plain.Length > 16000000 OrElse
                           (HelperClient.TextValue(message, "action") = "copySelection" AndAlso String.IsNullOrWhiteSpace(plain)) Then
                            Throw New System.IO.InvalidDataException("Invalid clipboard payload")
                        End If
                        ResponseClipboard.CopyRendered(html, plain)
                    End Using
                    statusKey = "desktop.copied"
                    SendView(New With {.[type] = "status", .status = UiStrings.GetText(statusKey, Language)})
                Catch ex As System.Exception
                    ClientLog.Write("clipboard.copy.failed", ex)
                    SendView(New With {.[type] = "error", .message = UiStrings.GetText("desktop.copyFailed", Language)})
                End Try
            Case "renderWarning"
                Dim imageCount As System.Text.Json.JsonElement
                Dim removedCount As System.Text.Json.JsonElement
                If message.TryGetProperty("images", imageCount) AndAlso imageCount.ValueKind = System.Text.Json.JsonValueKind.Number AndAlso imageCount.GetInt32() > 0 Then ClientLog.Write("render.external-image-replaced-with-alt")
                If message.TryGetProperty("removed", removedCount) AndAlso removedCount.ValueKind = System.Text.Json.JsonValueKind.Number AndAlso removedCount.GetInt32() > 0 Then ClientLog.Write("render.active-content-removed")
        End Select
    End Function

    Private Function DisplayedQuestion() As String
        If displayedTurnIndex < 0 OrElse displayedTurnIndex >= turns.Count Then Return ""
        Return turns(displayedTurnIndex).Question
    End Function

    Private Sub ShowTurn(index As Integer)
        displayedTurnIndex = index
        response = If(index >= 0 AndAlso index < turns.Count, turns(index).Answer, "")
        SendView(New With {.[type] = "answer", .html = ResponseClipboard.Render(response),
                           .question = question, .turnIndex = displayedTurnIndex, .turnCount = turns.Count,
                           .turnQuestion = DisplayedQuestion(), .includeHistory = includeHistory,
                           .status = UiStrings.GetText(statusKey, Language)})
    End Sub

    Private Function BuildQuestionWithHistory(currentQuestion As String) As String
        If Not includeHistory OrElse turns.Count = 0 Then Return currentQuestion
        Dim transcript As New System.Text.StringBuilder()
        transcript.AppendLine("Previous questions and AI responses for this selected text (chronological):")
        Dim index As Integer = 0
        For Each turn As ConversationTurn In turns
            index += 1
            transcript.AppendLine("Question " & index.ToString(System.Globalization.CultureInfo.InvariantCulture) & ":")
            transcript.AppendLine(turn.Question)
            transcript.AppendLine("Response " & index.ToString(System.Globalization.CultureInfo.InvariantCulture) & ":")
            transcript.AppendLine(turn.Answer)
        Next
        transcript.AppendLine("Current follow-up question:")
        transcript.Append(currentQuestion)
        If transcript.Length > 180000 Then Throw New HelperFailure("desktop.historyTooLong")
        Return transcript.ToString()
    End Function

    Private Async Function SendAsync(value As String) As System.Threading.Tasks.Task
        If sending OrElse settingsOpen Then
            SendView(New With {.[type] = "busy", .busy = sending, .status = UiStrings.GetText(statusKey, Language)})
            Return
        End If
        If String.IsNullOrWhiteSpace(value) OrElse value.Length > 200000 Then
            SendView(New With {.[type] = "error", .message = UiStrings.GetText("desktop.invalidQuestion", Language)})
            Return
        End If
        question = value
        sending = True
        context.SetBusy(True)
        statusKey = "desktop.sending"
        SendView(New With {.[type] = "busy", .busy = True, .status = UiStrings.GetText(statusKey, Language)})
        Try
            Dim effectiveQuestion As String = BuildQuestionWithHistory(question)
            Dim answer As String = Await HelperClient.AskAsync(ConfigStore.Clone(context.Settings()), effectiveQuestion, selection, Lifetime.Token)
            turns.Add(New ConversationTurn With {.Question = question, .Answer = answer})
            question = ""
            statusKey = "desktop.done"
            ShowTurn(turns.Count - 1)
        Catch ex As System.Exception
            If Not exiting Then
                ClientLog.Write("helper.ask.failed", ex)
                statusKey = "desktop.failed"
                SendView(New With {.[type] = "error", .message = HelperClient.Describe(ex, Language)})
            End If
        Finally
            sending = False
            context.SetBusy(False)
            SendView(New With {.[type] = "busy", .busy = False})
        End Try
    End Function

    Friend Sub Shutdown()
        exiting = True
        Lifetime.Cancel()
        Close()
        Dispose()
    End Sub

    Protected Overrides Sub OnFormClosing(e As System.Windows.Forms.FormClosingEventArgs)
        If Not exiting AndAlso e.CloseReason = System.Windows.Forms.CloseReason.UserClosing Then
            e.Cancel = True
            context.RememberAssistantPosition(Me.Location)
            Hide() ' Keep question and answer; closing the panel does not quit the tray application.
        End If
        MyBase.OnFormClosing(e)
    End Sub
End Class
