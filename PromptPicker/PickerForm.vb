' WebView2 presentation with a fully usable native WinForms fallback. The host owns
' prompt data and validates every selection; browser messages never carry file paths.
Friend NotInheritable Class PickerForm
    Inherits System.Windows.Forms.Form

    Private Const Origin As String = "https://prompt-picker.invalid/"
    Private ReadOnly config As LibraryConfiguration
    Private ReadOnly selectedText As String
    Private ReadOnly windowSettings As PickerSettings
    Private previewPercent As Integer
    Private nativeSplitter As System.Windows.Forms.SplitContainer
    Private commitPending As Boolean

    Private ReadOnly view As New Microsoft.Web.WebView2.WinForms.WebView2 With {.Dock = System.Windows.Forms.DockStyle.Fill}
    Private ReadOnly streams As New System.Collections.Generic.List(Of System.IO.Stream)()
    Private entries As System.Collections.Generic.List(Of PromptEntry)
    Private fallbackList As System.Windows.Forms.ListBox
    Private fallbackSearch As System.Windows.Forms.TextBox
    Private fallbackCategory As System.Windows.Forms.ComboBox
    Private fallbackPreview As System.Windows.Forms.TextBox
    Private fallbackSave As System.Windows.Forms.Button
    Private nativeMode As Boolean
    Private pageReady As Boolean
    Private result As String = ""

    Public ReadOnly Property ChosenPrompt As String
        Get
            Return result
        End Get
    End Property

    Public Sub New(settings As LibraryConfiguration, selection As String, prompts As System.Collections.Generic.List(Of PromptEntry), pickerSettings As PickerSettings)
        config = settings
        selectedText = If(selection, "")
        entries = prompts
        windowSettings = pickerSettings
        previewPercent = System.Math.Clamp(pickerSettings.PreviewPercent, 32, 68)
        Text = "Red Ink Prompt Picker"
        ClientSize = New System.Drawing.Size(System.Math.Clamp(pickerSettings.PickerClientWidth, 600, 2400), System.Math.Clamp(pickerSettings.PickerClientHeight, 390, 1600))
        MinimumSize = New System.Drawing.Size(600, 390)
        StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        TopMost = True
        BackColor = PickerUi.Surface
        Font = New System.Drawing.Font("Segoe UI", 9.0F)
        KeyPreview = True
        Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath)
        Controls.Add(view)
        AddHandler KeyDown, AddressOf HandleGlobalKeys
    End Sub

    Protected Overrides Sub OnLoad(e As System.EventArgs)
        MyBase.OnLoad(e)
        ' Apply stored physical client pixels *after* WinForms DPI autoscaling.
        ' Applying saved pixels in the constructor can scale them again on each
        ' start at 125/150/200 percent and make the window grow repeatedly.
        ClientSize = New System.Drawing.Size(System.Math.Clamp(windowSettings.PickerClientWidth, 600, 2400), System.Math.Clamp(windowSettings.PickerClientHeight, 390, 1600))
        RestoreWindowPosition()
    End Sub

    Private Sub RestoreWindowPosition()
        ' Keep the wider first-run default entirely reachable on smaller screens.
        ' Explicitly saved user sizes/positions remain authoritative whenever visible.
        If Not windowSettings.PickerX.HasValue OrElse Not windowSettings.PickerY.HasValue Then
            FitWindowToWorkArea(System.Windows.Forms.Screen.FromControl(Me).WorkingArea)
            Return
        End If
        Dim desired As New System.Drawing.Rectangle(windowSettings.PickerX.Value, windowSettings.PickerY.Value, Width, Height)
        Dim visibleScreen As System.Windows.Forms.Screen = Nothing
        Dim largest As Long = 0
        For Each display As System.Windows.Forms.Screen In System.Windows.Forms.Screen.AllScreens
            Dim intersection As System.Drawing.Rectangle = System.Drawing.Rectangle.Intersect(desired, display.WorkingArea)
            Dim area As Long = CLng(intersection.Width) * CLng(intersection.Height)
            If area > largest Then
                largest = area
                visibleScreen = display
            End If
        Next
        If visibleScreen Is Nothing OrElse largest < 80L * 80L Then
            FitWindowToWorkArea(System.Windows.Forms.Screen.FromControl(Me).WorkingArea)
            Return
        End If
        Dim limits As System.Drawing.Rectangle = visibleScreen.WorkingArea
        ' Bring an oversized old window back inside the visible work area.
        Dim bounds As New System.Drawing.Rectangle(desired.Location, New System.Drawing.Size(System.Math.Min(desired.Width, limits.Width), System.Math.Min(desired.Height, limits.Height)))
        bounds.X = System.Math.Clamp(bounds.X, limits.Left, limits.Right - bounds.Width)
        bounds.Y = System.Math.Clamp(bounds.Y, limits.Top, limits.Bottom - bounds.Height)
        StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Bounds = bounds
    End Sub

    Private Sub FitWindowToWorkArea(workArea As System.Drawing.Rectangle)
        If Width > workArea.Width OrElse Height > workArea.Height Then
            Size = New System.Drawing.Size(System.Math.Min(Width, workArea.Width), System.Math.Min(Height, workArea.Height))
        End If
    End Sub

    Public Sub CaptureWindowSettings()
        If WindowState = System.Windows.Forms.FormWindowState.Normal Then
            windowSettings.PickerX = Left
            windowSettings.PickerY = Top
            windowSettings.PickerClientWidth = ClientSize.Width
            windowSettings.PickerClientHeight = ClientSize.Height
        Else
            ' RestoreBounds uses outer dimensions; keep normal client size stable.
            windowSettings.PickerX = RestoreBounds.X
            windowSettings.PickerY = RestoreBounds.Y
        End If
        windowSettings.PreviewPercent = previewPercent
    End Sub

    Private Async Sub FocusFilter()
        Try
            If IsDisposed OrElse Not Visible Then Return
            If nativeMode Then
                If fallbackSearch IsNot Nothing AndAlso fallbackSearch.CanFocus Then fallbackSearch.Focus()
            ElseIf pageReady AndAlso view.CoreWebView2 IsNot Nothing Then
                view.Focus()
                ' A WebView2 HWND can be active without the HTML search input
                ' owning focus; put keyboard input there on each activation.
                Await view.CoreWebView2.ExecuteScriptAsync("document.getElementById('search')?.focus({preventScroll:true})")
            End If
        Catch ex As System.Exception
            AppLog.Write("picker.search.focus.failed", ex)
        End Try
    End Sub

    Protected Overrides Sub OnActivated(e As System.EventArgs)
        MyBase.OnActivated(e)
        If pageReady OrElse nativeMode Then BeginInvoke(New System.Action(AddressOf FocusFilter))
    End Sub

    Protected Overrides Async Sub OnShown(e As System.EventArgs)
        MyBase.OnShown(e)
        Try
            Dim root As String = System.IO.Path.Combine(AppLog.Root, "WebView2")
            Dim environment As Microsoft.Web.WebView2.Core.CoreWebView2Environment = Await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(Nothing, root)
            If IsDisposed Then Return
            Await view.EnsureCoreWebView2Async(environment)
            If IsDisposed Then Return
            Dim core As Microsoft.Web.WebView2.Core.CoreWebView2 = view.CoreWebView2
            core.Settings.AreHostObjectsAllowed = False
            core.Settings.AreDefaultScriptDialogsEnabled = False
            core.Settings.AreDefaultContextMenusEnabled = False
            core.Settings.AreDevToolsEnabled = False
            core.Settings.AreBrowserAcceleratorKeysEnabled = False
            core.Settings.IsStatusBarEnabled = False
            core.Settings.IsBuiltInErrorPageEnabled = False
            core.Settings.IsPasswordAutosaveEnabled = False
            core.Settings.IsGeneralAutofillEnabled = False
            core.Settings.IsWebMessageEnabled = True
            AddHandler core.NavigationStarting, AddressOf CheckNavigation
            AddHandler core.FrameNavigationStarting, Sub(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2NavigationStartingEventArgs) args.Cancel = True
            AddHandler core.NewWindowRequested, Sub(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2NewWindowRequestedEventArgs) args.Handled = True
            AddHandler core.PermissionRequested, Sub(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2PermissionRequestedEventArgs) args.State = Microsoft.Web.WebView2.Core.CoreWebView2PermissionState.Deny
            AddHandler core.DownloadStarting, Sub(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2DownloadStartingEventArgs) args.Cancel = True
            AddHandler core.WebResourceRequested, AddressOf ResourceRequested
            AddHandler core.WebMessageReceived, AddressOf WebMessageReceived
            AddHandler core.NavigationCompleted, AddressOf NavigationCompleted
            AddHandler core.ProcessFailed, Sub(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2ProcessFailedEventArgs)
                                               AppLog.Write("webview.process.failed")
                                               If Not IsDisposed Then ShowNativeFallback()
                                           End Sub
            core.AddWebResourceRequestedFilter("*", Microsoft.Web.WebView2.Core.CoreWebView2WebResourceContext.All)
            core.Navigate(Origin & "picker.html")
        Catch ex As System.Exception
            AppLog.Write("webview.start.failed.native-fallback", ex)
            If Not IsDisposed Then ShowNativeFallback()
        End Try
    End Sub

    Private Sub CheckNavigation(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2NavigationStartingEventArgs)
        If Not System.String.Equals(args.Uri, Origin & "picker.html", System.StringComparison.OrdinalIgnoreCase) Then args.Cancel = True
    End Sub

    Private Sub ResourceRequested(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2WebResourceRequestedEventArgs)
        Dim resource As System.Uri = Nothing
        Dim mime As String = "text/plain; charset=utf-8"
        Dim contents As Byte() = System.Array.Empty(Of Byte)()
        Dim status As Integer = 403
        Try
            If System.String.Equals(args.Request.Method, "GET", System.StringComparison.OrdinalIgnoreCase) AndAlso
               System.Uri.TryCreate(args.Request.Uri, System.UriKind.Absolute, resource) AndAlso
               System.String.Equals(resource.GetLeftPart(System.UriPartial.Authority) & "/", Origin, System.StringComparison.OrdinalIgnoreCase) AndAlso resource.Query = "" Then
                Dim name As String = resource.AbsolutePath.TrimStart("/"c)
                If name = "picker.html" OrElse name = "picker.css" OrElse name = "picker.js" Then
                    Dim path As String = "PromptPicker.Web." & name
                    Using body As System.IO.Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(path)
                        If body IsNot Nothing Then
                            Using output As New System.IO.MemoryStream()
                                body.CopyTo(output)
                                contents = output.ToArray()
                            End Using
                            status = 200
                            mime = If(name.EndsWith(".html", System.StringComparison.Ordinal), "text/html; charset=utf-8", If(name.EndsWith(".css", System.StringComparison.Ordinal), "text/css; charset=utf-8", "text/javascript; charset=utf-8"))
                        End If
                    End Using
                End If
            End If
        Catch ex As System.Exception
            AppLog.Write("webview.resource.failed", ex)
            status = 500
        End Try
        Try
            Dim bytes As New System.IO.MemoryStream(contents, False)
            streams.Add(bytes)
            args.Response = view.CoreWebView2.Environment.CreateWebResourceResponse(bytes, status, If(status = 200, "OK", "Blocked"), "Content-Type: " & mime & vbCrLf & "Cache-Control: no-store" & vbCrLf & "X-Content-Type-Options: nosniff")
        Catch ex As System.Exception
            AppLog.Write("webview.resource.response.failed", ex)
        End Try
    End Sub

    Private Sub NavigationCompleted(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs)
        If Not args.IsSuccess Then
            AppLog.Write("webview.navigation.failed.native-fallback")
            ShowNativeFallback()
            Return
        End If
        pageReady = True
        SendModel()
        FocusFilter()
    End Sub

    Private Sub SendModel()
        If nativeMode OrElse Not pageReady OrElse IsDisposed OrElse view.CoreWebView2 Is Nothing Then Return
        Try
            view.CoreWebView2.PostWebMessageAsJson(System.Text.Json.JsonSerializer.Serialize(New With {
                .[type] = "model",
                .items = entries,
                .hasSelection = selectedText.Length > 0,
                .canSave = selectedText.Length > 0 AndAlso PromptLibraryStore.CanSaveLocal(config),
                .hasLocal = config.HasLocal,
                .selectionLength = selectedText.Length,
                .previewPercent = previewPercent
            }))
        Catch ex As System.Exception
            AppLog.Write("webview.model.failed", ex)
            ShowNativeFallback()
        End Try
    End Sub

    Private Sub WebMessageReceived(sender As Object, e As Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs)
        If Not System.String.Equals(e.Source, Origin & "picker.html", System.StringComparison.OrdinalIgnoreCase) Then Return
        Try
            Dim json As String = e.WebMessageAsJson
            If json.Length > 1048576 Then Return
            Using document As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(json)
                Dim payload As System.Text.Json.JsonElement = document.RootElement
                If payload.ValueKind <> System.Text.Json.JsonValueKind.Object Then Return
                Dim actionField As System.Text.Json.JsonElement
                If Not payload.TryGetProperty("type", actionField) OrElse actionField.ValueKind <> System.Text.Json.JsonValueKind.String Then Return
                Select Case actionField.GetString()
                    Case "select"
                        Dim idField As System.Text.Json.JsonElement
                        If payload.TryGetProperty("id", idField) AndAlso idField.ValueKind = System.Text.Json.JsonValueKind.Number Then
                            Dim id As Integer
                            If idField.TryGetInt32(id) Then
                                ' Avoid disposing the WebView2 during its own input/message
                                ' dispatch (notably the second click of a double-click).
                                If Not commitPending Then
                                    commitPending = True
                                    BeginInvoke(New System.Action(Sub() CommitSelection(id)))
                                End If
                            End If
                        End If
                    Case "split"
                        Dim percentField As System.Text.Json.JsonElement
                        Dim value As Integer
                        If payload.TryGetProperty("previewPercent", percentField) AndAlso
                           percentField.ValueKind = System.Text.Json.JsonValueKind.Number AndAlso
                           percentField.TryGetInt32(value) AndAlso value >= 32 AndAlso value <= 68 Then
                            previewPercent = value
                        End If
                    Case "save"
                        Dim categoryField As System.Text.Json.JsonElement
                        Dim suggestedCategory As String = ""
                        If payload.TryGetProperty("category", categoryField) AndAlso categoryField.ValueKind = System.Text.Json.JsonValueKind.String Then
                            Dim candidate As String = categoryField.GetString()
                            If candidate IsNot Nothing AndAlso candidate.Length < 256 AndAlso
                               entries.Any(Function(item As PromptEntry) System.String.Equals(item.Category, candidate, System.StringComparison.OrdinalIgnoreCase)) Then
                                suggestedCategory = candidate
                            End If
                        End If
                        SaveSelection(suggestedCategory)
                    Case "close"
                        DialogResult = System.Windows.Forms.DialogResult.Cancel
                        Close()
                End Select
            End Using
        Catch ex As System.Exception
            AppLog.Write("webview.message.rejected", ex)
        End Try
    End Sub

    Private Sub CommitSelection(id As Integer)
        If id < 0 OrElse id >= entries.Count Then
            commitPending = False
            Return
        End If
        AppLog.Write("picker.selection.commit id=" & id.ToString(System.Globalization.CultureInfo.InvariantCulture))
        result = entries(id).Prompt.Replace("\n", vbCrLf)
        DialogResult = System.Windows.Forms.DialogResult.OK
        Close()
    End Sub

    Private Sub SaveSelection(Optional suggestedCategory As String = "")
        If selectedText.Length = 0 OrElse Not PromptLibraryStore.CanSaveLocal(config) Then Return
        Dim chosenCategory As String = suggestedCategory
        If nativeMode AndAlso fallbackCategory IsNot Nothing AndAlso fallbackCategory.SelectedItem IsNot Nothing Then
            chosenCategory = fallbackCategory.SelectedItem.ToString()
        End If
        Using saveForm As New SavePromptForm(selectedText, If(chosenCategory = "Alle Kategorien", "", chosenCategory))
            If saveForm.ShowDialog(Me) <> System.Windows.Forms.DialogResult.OK Then Return
            Try
                PromptLibraryStore.SaveLocal(config, saveForm.PromptTitle, saveForm.PromptBody, saveForm.PromptCategory)
                entries = PromptLibraryStore.Load(config)
                If nativeMode Then RefreshFallback() Else SendModel()
                System.Windows.Forms.MessageBox.Show(Me, "Der Prompt wurde in der lokalen Bibliothek gespeichert.", "Red Ink Prompt Picker", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
            Catch ex As System.Exception
                AppLog.Write("library.local.save.failed", ex)
                System.Windows.Forms.MessageBox.Show(Me, "Lokale Prompt Library konnte nicht geschrieben werden: " & ex.Message, "Red Ink Prompt Picker", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    Private Sub HandleGlobalKeys(sender As Object, e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = System.Windows.Forms.Keys.Escape Then
            DialogResult = System.Windows.Forms.DialogResult.Cancel
            Close()
            e.Handled = True
        End If
    End Sub

    Private Sub ShowNativeFallback()
        If nativeMode OrElse IsDisposed Then Return
        nativeMode = True
        view.Visible = False
        Dim root As New System.Windows.Forms.TableLayoutPanel With {.Dock = System.Windows.Forms.DockStyle.Fill, .Padding = New System.Windows.Forms.Padding(11), .ColumnCount = 2, .RowCount = 4, .BackColor = System.Drawing.Color.FromArgb(248, 248, 249)}
        root.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 55))
        root.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 45))
        root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 44))
        root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
        root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24))
        root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 42))
        fallbackSearch = New System.Windows.Forms.TextBox With {.Dock = System.Windows.Forms.DockStyle.Fill, .PlaceholderText = "Prompts durchsuchen"}
        fallbackCategory = New System.Windows.Forms.ComboBox With {.Dock = System.Windows.Forms.DockStyle.Fill, .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
        root.Controls.Add(fallbackSearch, 0, 0)
        root.Controls.Add(fallbackCategory, 1, 0)
        fallbackList = New System.Windows.Forms.ListBox With {.Dock = System.Windows.Forms.DockStyle.Fill, .DisplayMember = "Title", .IntegralHeight = False}
        fallbackPreview = New System.Windows.Forms.TextBox With {.Dock = System.Windows.Forms.DockStyle.Fill, .Multiline = True, .ReadOnly = True, .ScrollBars = System.Windows.Forms.ScrollBars.Vertical}
        nativeSplitter = New System.Windows.Forms.SplitContainer With {
            .Dock = System.Windows.Forms.DockStyle.Fill,
            .Orientation = System.Windows.Forms.Orientation.Vertical,
            .Size = New System.Drawing.Size(650, 350),
            .SplitterWidth = 6,
            .Panel1MinSize = 170,
            .Panel2MinSize = 180,
            .Margin = New System.Windows.Forms.Padding(0)
        }
        nativeSplitter.Panel1.Controls.Add(fallbackList)
        nativeSplitter.Panel2.Controls.Add(fallbackPreview)
        root.Controls.Add(nativeSplitter, 0, 1)
        root.SetColumnSpan(nativeSplitter, 2)
        AddHandler nativeSplitter.SplitterMoved, Sub(sender As Object, e As System.Windows.Forms.SplitterEventArgs)
                                                     If nativeSplitter.ClientSize.Width > nativeSplitter.SplitterWidth Then
                                                         previewPercent = System.Math.Clamp(CInt(System.Math.Round(100.0R * (nativeSplitter.ClientSize.Width - nativeSplitter.SplitterWidth - nativeSplitter.SplitterDistance) / (nativeSplitter.ClientSize.Width - nativeSplitter.SplitterWidth))), 32, 68)
                                                     End If
                                                 End Sub
        AddHandler nativeSplitter.SizeChanged, Sub(sender As Object, e As System.EventArgs)
                                                    If nativeSplitter.ClientSize.Width > nativeSplitter.Panel1MinSize + nativeSplitter.Panel2MinSize + nativeSplitter.SplitterWidth Then
                                                        Dim available As Integer = nativeSplitter.ClientSize.Width - nativeSplitter.SplitterWidth
                                                        nativeSplitter.SplitterDistance = System.Math.Clamp(CInt(System.Math.Round(available * (100 - previewPercent) / 100.0R)), nativeSplitter.Panel1MinSize, available - nativeSplitter.Panel2MinSize)
                                                    End If
                                                End Sub
        Dim sourceText As String = If(config.HasLocal, "Lokal + Zentral", "Nur zentrale Library") & "  ·  " & entries.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " Prompts"
        root.Controls.Add(New System.Windows.Forms.Label With {.Text = sourceText, .Dock = System.Windows.Forms.DockStyle.Fill, .TextAlign = System.Drawing.ContentAlignment.MiddleLeft}, 0, 2)

        Dim bar As New System.Windows.Forms.FlowLayoutPanel With {.Dock = System.Windows.Forms.DockStyle.Fill, .FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft, .WrapContents = False}
        Dim choose As System.Windows.Forms.Button = PickerUi.ActionButton("Einfügen", "insert", True, 112)
        Dim close As System.Windows.Forms.Button = PickerUi.ActionButton("Abbrechen", "close", False, 123)
        fallbackSave = PickerUi.ActionButton("Markierung speichern", "plus", False, 185)
        fallbackSave.Enabled = selectedText.Length > 0 AndAlso PromptLibraryStore.CanSaveLocal(config)
        bar.Controls.Add(choose)
        bar.Controls.Add(close)
        bar.Controls.Add(fallbackSave)
        root.Controls.Add(bar, 0, 3)
        root.SetColumnSpan(bar, 2)
        Controls.Add(root)
        root.BringToFront()
        AddHandler fallbackSearch.KeyDown, Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
                                               If fallbackList.Items.Count = 0 Then Return
                                               If e.KeyCode = System.Windows.Forms.Keys.Down OrElse e.KeyCode = System.Windows.Forms.Keys.Up Then
                                                   fallbackList.SelectedIndex = System.Math.Clamp(fallbackList.SelectedIndex + If(e.KeyCode = System.Windows.Forms.Keys.Down, 1, -1), 0, fallbackList.Items.Count - 1)
                                                   e.Handled = True
                                                   e.SuppressKeyPress = True
                                               ElseIf e.KeyCode = System.Windows.Forms.Keys.Enter Then
                                                   ChooseFromFallback()
                                                   e.Handled = True
                                                   e.SuppressKeyPress = True
                                               End If
                                           End Sub
        AddHandler fallbackSearch.TextChanged, Sub(sender As Object, e As System.EventArgs) RefreshFallbackList()
        AddHandler fallbackCategory.SelectedIndexChanged, Sub(sender As Object, e As System.EventArgs) RefreshFallbackList()
        AddHandler fallbackList.SelectedIndexChanged, Sub(sender As Object, e As System.EventArgs)
                                                          Dim item As PromptEntry = TryCast(fallbackList.SelectedItem, PromptEntry)
                                                          fallbackPreview.Text = If(item Is Nothing, "", item.Prompt.Replace("\n", vbCrLf))
                                                      End Sub
        AddHandler fallbackList.MouseDoubleClick, Sub(sender As Object, e As System.Windows.Forms.MouseEventArgs)
                                                      If fallbackList.IndexFromPoint(e.Location) >= 0 Then
                                                          fallbackList.SelectedIndex = fallbackList.IndexFromPoint(e.Location)
                                                          ChooseFromFallback()
                                                      End If
                                                  End Sub
        AddHandler fallbackList.KeyDown, Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
                                             If e.KeyCode = System.Windows.Forms.Keys.Enter Then
                                                 ChooseFromFallback()
                                                 e.Handled = True
                                                 e.SuppressKeyPress = True
                                             End If
                                         End Sub
        AddHandler choose.Click, Sub(sender As Object, e As System.EventArgs) ChooseFromFallback()
        AddHandler close.Click, Sub(sender As Object, e As System.EventArgs)
                                    DialogResult = System.Windows.Forms.DialogResult.Cancel
                                    Me.Close()
                                End Sub
        AddHandler fallbackSave.Click, Sub(sender As Object, e As System.EventArgs) SaveSelection()
        RefreshFallback()
        BeginInvoke(New System.Action(Sub()
            If nativeSplitter IsNot Nothing AndAlso nativeSplitter.ClientSize.Width > nativeSplitter.Panel1MinSize + nativeSplitter.Panel2MinSize + nativeSplitter.SplitterWidth Then
                Dim available As Integer = nativeSplitter.ClientSize.Width - nativeSplitter.SplitterWidth
                nativeSplitter.SplitterDistance = System.Math.Clamp(CInt(System.Math.Round(available * (100 - previewPercent) / 100.0R)), nativeSplitter.Panel1MinSize, available - nativeSplitter.Panel2MinSize)
            End If
            FocusFilter()
        End Sub))
    End Sub

    Private Sub ChooseFromFallback()
        Dim chosen As PromptEntry = TryCast(fallbackList.SelectedItem, PromptEntry)
        If chosen IsNot Nothing Then CommitSelection(chosen.Id)
    End Sub

    Private Sub RefreshFallback()
        If fallbackCategory Is Nothing Then Return
        Dim selected As String = If(fallbackCategory.SelectedItem Is Nothing, "Alle Kategorien", fallbackCategory.SelectedItem.ToString())
        fallbackCategory.Items.Clear()
        fallbackCategory.Items.Add("Alle Kategorien")
        Dim setOfCategories As New System.Collections.Generic.SortedSet(Of String)(System.StringComparer.CurrentCultureIgnoreCase)
        For Each entry As PromptEntry In entries
            setOfCategories.Add(entry.Category)
        Next
        For Each category As String In setOfCategories
            fallbackCategory.Items.Add(category)
        Next
        fallbackCategory.SelectedItem = If(fallbackCategory.Items.Contains(selected), selected, "Alle Kategorien")
        RefreshFallbackList()
    End Sub

    Private Sub RefreshFallbackList()
        If fallbackList Is Nothing OrElse fallbackCategory Is Nothing Then Return
        Dim filter As String = If(fallbackSearch.Text, "").Trim()
        Dim category As String = If(fallbackCategory.SelectedItem Is Nothing, "Alle Kategorien", fallbackCategory.SelectedItem.ToString())
        fallbackList.BeginUpdate()
        fallbackList.Items.Clear()
        For Each entry As PromptEntry In entries
            If category <> "Alle Kategorien" AndAlso Not System.String.Equals(entry.Category, category, System.StringComparison.OrdinalIgnoreCase) Then Continue For
            If filter.Length > 0 AndAlso entry.Title.IndexOf(filter, System.StringComparison.OrdinalIgnoreCase) < 0 AndAlso entry.Prompt.IndexOf(filter, System.StringComparison.OrdinalIgnoreCase) < 0 Then Continue For
            fallbackList.Items.Add(entry)
        Next
        fallbackList.EndUpdate()
        fallbackList.SelectedIndex = If(fallbackList.Items.Count > 0, 0, -1)
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            For Each item As System.IO.Stream In streams
                Try
                    item.Dispose()
                Catch
                End Try
            Next
            streams.Clear()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class
