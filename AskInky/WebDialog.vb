' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: WebDialog.vb
' Purpose: Ask Inky desktop functionality in the modular Windows client.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). Shared deterministic WebView2 shell and strict message boundary.
Friend MustInherit Class WebDialog
    Inherits System.Windows.Forms.Form
    Protected ReadOnly View As New Microsoft.Web.WebView2.WinForms.WebView2 With {.Dock = System.Windows.Forms.DockStyle.Fill}
    Protected ReadOnly Lifetime As New System.Threading.CancellationTokenSource()
    Private ReadOnly pageName As String
    Private ReadOnly resourceStreams As New System.Collections.Generic.List(Of System.IO.Stream)()
    Private initialized As Boolean
    Private initializing As Boolean
    Private disposedResources As Boolean
    Private failurePanel As System.Windows.Forms.Panel
    Protected Property ViewReady As Boolean
    Private Const Origin As String = "https://redink-ui.invalid/"

    Protected Sub New(page As String, title As String, size As System.Drawing.Size)
        pageName = page
        Text = title
        ClientSize = size
        MinimumSize = New System.Drawing.Size(376, 420)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Font = New System.Drawing.Font("Segoe UI", 9.0F)
        BackColor = System.Drawing.Color.FromArgb(246, 246, 247)
        Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath)
        Controls.Add(View)
    End Sub

    Protected MustOverride ReadOnly Property Language As String
    Protected MustOverride Sub PageReady()
    Protected MustOverride Function HandleMessageAsync(message As System.Text.Json.JsonElement) As System.Threading.Tasks.Task

    Protected Overrides Async Sub OnShown(e As System.EventArgs)
        MyBase.OnShown(e)
        Await InitializeViewAsync()
    End Sub

    Private Async Function InitializeViewAsync() As System.Threading.Tasks.Task
        If initialized OrElse initializing OrElse IsDisposed Then Return
        initializing = True
        Try
            If View.CoreWebView2 Is Nothing Then
                Dim profile As String = System.IO.Path.Combine(StoragePaths.Root, "WebView2")
                Dim environment As Microsoft.Web.WebView2.Core.CoreWebView2Environment = Await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(Nothing, profile)
                If IsDisposed Then Return
                Await View.EnsureCoreWebView2Async(environment)
            End If
            If IsDisposed Then Return
            Dim core As Microsoft.Web.WebView2.Core.CoreWebView2 = View.CoreWebView2
            core.Settings.AreHostObjectsAllowed = False
            core.Settings.AreDefaultScriptDialogsEnabled = False
            core.Settings.AreDefaultContextMenusEnabled = False
            core.Settings.AreDevToolsEnabled = False
            core.Settings.AreBrowserAcceleratorKeysEnabled = False
            core.Settings.IsWebMessageEnabled = True
            core.Settings.IsScriptEnabled = True
            core.Settings.IsStatusBarEnabled = False
            core.Settings.IsBuiltInErrorPageEnabled = False
            core.Settings.IsGeneralAutofillEnabled = False
            core.Settings.IsPasswordAutosaveEnabled = False
            AddHandler core.NavigationStarting, AddressOf NavigationStarting
            AddHandler core.NavigationCompleted, Sub(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs)
                                                     If Not args.IsSuccess Then
                                                         ClientLog.Write("webview.navigation.failed")
                                                         ShowFailure()
                                                     End If
                                                 End Sub
            AddHandler core.FrameNavigationStarting, Sub(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2NavigationStartingEventArgs) args.Cancel = True
            AddHandler core.NewWindowRequested, Sub(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2NewWindowRequestedEventArgs) args.Handled = True
            AddHandler core.DownloadStarting, Sub(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2DownloadStartingEventArgs) args.Cancel = True
            AddHandler core.PermissionRequested, Sub(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2PermissionRequestedEventArgs) args.State = Microsoft.Web.WebView2.Core.CoreWebView2PermissionState.Deny
            AddHandler core.WebResourceRequested, AddressOf ResourceRequested
            AddHandler core.WebMessageReceived, AddressOf MessageReceived
            AddHandler core.ProcessFailed, AddressOf BrowserFailed
            core.AddWebResourceRequestedFilter("*", Microsoft.Web.WebView2.Core.CoreWebView2WebResourceContext.All)
            initialized = True
            If failurePanel IsNot Nothing Then
                failurePanel.Dispose()
                failurePanel = Nothing
            End If
            View.Visible = True
            core.Navigate(Origin & pageName)
        Catch ex As System.Exception
            ClientLog.Write("webview.initialize.failed", ex)
            ShowFailure()
        Finally
            initializing = False
        End Try
    End Function

    Private Sub NavigationStarting(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2NavigationStartingEventArgs)
        ' No remote page, file URI, data URI, fragment navigation or reload with altered query.
        If Not String.Equals(args.Uri, Origin & pageName, System.StringComparison.Ordinal) Then
            args.Cancel = True
            ClientLog.Write("webview.navigation.blocked")
        End If
    End Sub

    Private Sub ResourceRequested(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2WebResourceRequestedEventArgs)
        Dim bytes As Byte() = System.Array.Empty(Of Byte)()
        Dim contentType As String = "text/plain; charset=utf-8"
        Dim code As Integer = 403
        Dim resource As System.Uri = Nothing
        Try
            If System.Uri.TryCreate(args.Request.Uri, System.UriKind.Absolute, resource) AndAlso resource.GetLeftPart(System.UriPartial.Authority) & "/" = Origin AndAlso resource.Query.Length = 0 AndAlso args.Request.Method = "GET" Then
                Dim name As String = resource.AbsolutePath.TrimStart("/"c)
                If WebAssets.IsPublicAsset(name) Then
                    bytes = WebAssets.ReadBytes(name)
                    contentType = WebAssets.ContentType(name)
                    code = 200
                End If
            End If
        Catch ex As System.Exception
            ClientLog.Write("webview.resource.failed", ex)
            code = 500
        End Try
        ' Creating a response can fail if the browser process is gone while a request is
        ' in flight. Keep only streams actually handed off; a native event must not throw.
        Dim stream As System.IO.MemoryStream = Nothing
        Try
            stream = New System.IO.MemoryStream(bytes, False)
            args.Response = View.CoreWebView2.Environment.CreateWebResourceResponse(stream, code, If(code = 200, "OK", "Blocked"), "Content-Type: " & contentType & vbCrLf & "Cache-Control: no-store" & vbCrLf & "X-Content-Type-Options: nosniff")
            resourceStreams.Add(stream) ' WebView2 may still be reading after this callback.
        Catch ex As System.Exception
            If stream IsNot Nothing Then stream.Dispose()
            ClientLog.Write("webview.resource.response-failed", ex)
        End Try
    End Sub

    Protected Function IsTrustedPage() As Boolean
        Try
            Return Not IsDisposed AndAlso Not View.IsDisposed AndAlso View.CoreWebView2 IsNot Nothing AndAlso String.Equals(View.CoreWebView2.Source, Origin & pageName, System.StringComparison.Ordinal)
        Catch ex As System.Exception
            ClientLog.Write("webview.origin.unavailable", ex)
            Return False
        End Try
    End Function

    Private Async Sub MessageReceived(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs)
        If Not IsTrustedPage() OrElse Not String.Equals(args.Source, Origin & pageName, System.StringComparison.Ordinal) Then Return
        Try
            Dim json As String = args.WebMessageAsJson
            If json.Length > 1048576 Then Throw New System.IO.InvalidDataException("UI message too large")
            Using document As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(json)
                Dim message As System.Text.Json.JsonElement = document.RootElement
                If message.ValueKind <> System.Text.Json.JsonValueKind.Object OrElse HelperClient.TextValue(message, "version") <> "1" Then Return
                If HelperClient.TextValue(message, "action") = "ready" Then
                    ViewReady = True
                    PageReady()
                ElseIf ViewReady AndAlso HelperClient.TextValue(message, "action") = "openLink" Then
                    OpenLink(HelperClient.TextValue(message, "url"))
                ElseIf ViewReady Then
                    Await HandleMessageAsync(message)
                End If
            End Using
        Catch ex As System.Exception
            ClientLog.Write("webview.message.failed", ex)
            SendView(New With {.[type] = "error", .message = HelperClient.Describe(ex, Language)})
        End Try
    End Sub

    Private Sub OpenLink(value As String)
        Dim target As System.Uri = Nothing
        If value.Length > 8192 OrElse Not System.Uri.TryCreate(value, System.UriKind.Absolute, target) Then Return
        If target.UserInfo.Length <> 0 OrElse (target.Scheme <> "https" AndAlso target.Scheme <> "http" AndAlso target.Scheme <> "mailto") Then Return
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(target.AbsoluteUri) With {.UseShellExecute = True})
    End Sub

    Protected Sub SendView(value As Object)
        If Not ViewReady OrElse Not IsTrustedPage() Then Return
        Try
            View.CoreWebView2.PostWebMessageAsJson(System.Text.Json.JsonSerializer.Serialize(value))
        Catch ex As System.Exception
            ClientLog.Write("webview.post.failed", ex)
            ViewReady = False
            ShowFailure()
        End Try
    End Sub

    Protected Async Function ReadCopyPayloadAsync(Optional selectionOnly As Boolean = False) As System.Threading.Tasks.Task(Of String)
        If Not ViewReady OrElse Not IsTrustedPage() Then Throw New System.InvalidOperationException("View not ready")
        Dim script As String = If(selectionOnly, "window.redInkAnswer.getSelectedCopyPayload()", "window.redInkAnswer.getCopyPayload()")
        Dim value As String = Await View.CoreWebView2.ExecuteScriptAsync(script)
        If Not IsTrustedPage() Then Throw New System.InvalidOperationException("View origin changed")
        Return value
    End Function

    Private Sub BrowserFailed(sender As Object, args As Microsoft.Web.WebView2.Core.CoreWebView2ProcessFailedEventArgs)
        ViewReady = False
        ClientLog.Write("webview.process.failed." & args.ProcessFailedKind.ToString())
        ShowFailure()
    End Sub

    Private Sub ShowFailure()
        If IsDisposed OrElse failurePanel IsNot Nothing Then Return
        View.Visible = False
        failurePanel = New System.Windows.Forms.Panel With {.Dock = System.Windows.Forms.DockStyle.Fill, .Padding = New System.Windows.Forms.Padding(20)}
        Dim text As New System.Windows.Forms.Label With {.Dock = System.Windows.Forms.DockStyle.Fill, .Text = UiStrings.GetText("desktop.webviewMissing", Language), .AutoSize = False}
        failurePanel.Controls.Add(text)
        Controls.Add(failurePanel)
        failurePanel.BringToFront()
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso Not disposedResources Then
            disposedResources = True
            Lifetime.Cancel()
            ViewReady = False
            View.Dispose()
            For Each stream As System.IO.Stream In resourceStreams
                stream.Dispose()
            Next
            resourceStreams.Clear()
            Lifetime.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class
