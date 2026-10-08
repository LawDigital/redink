' Part of Red Ink (AskInky). Native title bar; the entire client area is local HTML/CSS.
Friend NotInheritable Class AssistantForm
    Inherits WebDialog
    Private ReadOnly context As AssistantContext
    Private selection As String = ""
    Private question As String = ""
    Private response As String = ""
    Private statusKey As String = "desktop.ready"
    Private sending As Boolean
    Private settingsOpen As Boolean
    Private exiting As Boolean

    Public Sub New(parent As AssistantContext)
        MyBase.New("assistant.html", "Ask Inky", New System.Drawing.Size(468, 462))
        context = parent
    End Sub

    Protected Overrides ReadOnly Property Language As String
        Get
            Return context.Settings().UiLanguage
        End Get
    End Property

    Protected Overrides Sub PageReady()
        SendView(New With {.[type] = "init", .strings = UiStrings.GetBundle(Language), .language = Language,
                           .darkMode = context.Settings().DarkMode, .fontSize = context.Settings().FontSize,
                           .selectedText = selection, .question = question, .html = ResponseClipboard.Render(response),
                           .busy = sending, .locked = settingsOpen, .status = UiStrings.GetText(statusKey, Language),
                           .shortcut = HotkeyService.Format(context.Settings().HotkeyModifiers, context.Settings().HotkeyKey)})
    End Sub

    Friend Sub ApplySettings()
        SendView(New With {.[type] = "preferences", .strings = UiStrings.GetBundle(Language), .language = Language,
                           .darkMode = context.Settings().DarkMode, .fontSize = context.Settings().FontSize,
                           .shortcut = HotkeyService.Format(context.Settings().HotkeyModifiers, context.Settings().HotkeyKey),
                           .status = UiStrings.GetText(statusKey, Language)})
    End Sub

    Friend Sub LoadSelection(value As String)
        If sending Then Return
        selection = If(value, "")
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
            Case "questionChanged"
                Dim value As String = HelperClient.TextValue(message, "question")
                If value.Length <= 200000 AndAlso Not sending Then question = value
            Case "settings"
                If Not sending Then context.ShowSettings()
            Case "send"
                Await SendAsync(HelperClient.TextValue(message, "question"))
            Case "copy"
                If sending OrElse settingsOpen OrElse response.Length = 0 Then Return
                Try
                    Dim json As String = Await ReadCopyPayloadAsync()
                    Using document As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(json)
                        Dim root As System.Text.Json.JsonElement = document.RootElement
                        Dim html As String = HelperClient.TextValue(root, "html")
                        Dim plain As String = HelperClient.TextValue(root, "text")
                        If html.Length = 0 OrElse html.Length > 16000000 Then Throw New System.IO.InvalidDataException("Invalid clipboard payload")
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
            Dim answer As String = Await HelperClient.AskAsync(ConfigStore.Clone(context.Settings()), question, selection, Lifetime.Token)
            response = answer
            statusKey = "desktop.done"
            SendView(New With {.[type] = "answer", .html = ResponseClipboard.Render(response), .status = UiStrings.GetText(statusKey, Language)})
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
            Hide() ' Keep question and answer; closing the panel does not quit the tray application.
        End If
        MyBase.OnFormClosing(e)
    End Sub
End Class
