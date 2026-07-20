' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: MCPOAuthBrowserForm.vb
' Purpose: Hosts an embedded WebView2-based OAuth sign-in window for MCP and
'          related authentication flows, captures the redirect callback URL,
'          and returns authorization results without requiring copy/paste.
'
' Responsibilities:
'  - Initializes an embedded WebView2 browser with a temporary user-data folder.
'  - Navigates to the supplied authorization URL.
'  - Intercepts navigation to the configured redirect URI prefix.
'  - Captures callback URLs and OAuth error parameters for the caller.
'  - Redirects popup/new-window requests back into the same embedded browser.
'
' Notes:
'  - Intended for interactive OAuth sign-in within desktop add-ins.
'  - Uses WebView2 to support modern identity providers and SSO flows.
' =============================================================================



Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary

    ''' <summary>
    ''' Embedded OAuth sign-in window using WebView2 (Edge/Chromium).
    ''' Hosts the authorization endpoint and intercepts navigation to the configured
    ''' redirect URI prefix. The full callback URL (containing code/state/error) is
    ''' exposed via <see cref="CapturedRedirectUrl"/>. No copy/paste is required.
    ''' </summary>
    Public Class MCPOAuthBrowserForm
        Inherits Form

        Private WithEvents _webView As WebView2

        Private ReadOnly _authUrl As String
        Private ReadOnly _redirectUriPrefix As String
        Private _capturedReported As Boolean = False

        ''' <summary>Full callback URL captured from the redirect (contains code/state).</summary>
        Public Property CapturedRedirectUrl As String = ""

        ''' <summary>Optional error description captured from the callback (if `error=` was in the query).</summary>
        Public Property CapturedError As String = ""

        Public Sub New(authUrl As String, redirectUriPrefix As String)
            _authUrl = authUrl
            _redirectUriPrefix = redirectUriPrefix

            InitializeComponent()
            InitializeWebViewAsync()
        End Sub

        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)

            Try
                Me.TopMost = True
                Me.Activate()
                Me.BringToFront()
            Catch
            End Try
        End Sub

        Private Sub InitializeComponent()
            Me.Text = $"{AN} – OAuth Sign-In"
            Me.ClientSize = New Size(720, 800)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MinimizeBox = True
            Me.MaximizeBox = True
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.ShowInTaskbar = True
            Me.TopMost = True
            Me.KeyPreview = True

            Try
                Dim bmp As New Bitmap(GetLogoBitmap(LogoType.Standard))
                Me.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            _webView = New WebView2() With {.Dock = DockStyle.Fill}
            Me.Controls.Add(_webView)
        End Sub

        Private Async Sub InitializeWebViewAsync()
            Try
                Dim userDataFolder As String = SharedLibrary.SharedMethods.GetWebView2UserDataFolder()

                Dim env As CoreWebView2Environment =
                    Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder)

                Await _webView.EnsureCoreWebView2Async(env)

                AddHandler _webView.CoreWebView2.ProcessFailed,
                    Sub(s, e)
                        SharedLibrary.SharedMethods.LogWebView2ProcessFailed("OAuth", e.ProcessFailedKind.ToString(), e.ExitCode.ToString())
                        Try
                            If e.ProcessFailedKind = CoreWebView2ProcessFailedKind.RenderProcessExited Then
                                _webView.Reload()
                            End If
                        Catch
                        End Try
                    End Sub

                _webView.CoreWebView2.Settings.AreDevToolsEnabled = False
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = True
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = True
                _webView.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = True
                _webView.CoreWebView2.Settings.IsZoomControlEnabled = True

                AddHandler _webView.CoreWebView2.NavigationStarting, AddressOf OnNavigationStarting
                AddHandler _webView.CoreWebView2.NewWindowRequested, AddressOf OnNewWindowRequested

                _webView.CoreWebView2.Navigate(_authUrl)

            Catch ex As Exception
                MessageBox.Show(
                    $"Could not initialize the OAuth sign-in window: {ex.Message}",
                    $"{AN} – OAuth",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
                Me.DialogResult = DialogResult.Cancel
                Me.Close()
            End Try
        End Sub

        Private Sub OnNavigationStarting(sender As Object, e As CoreWebView2NavigationStartingEventArgs)
            Try
                Dim url As String = If(e.Uri, "")
                If String.IsNullOrEmpty(url) Then Return

                If url.StartsWith(_redirectUriPrefix, StringComparison.OrdinalIgnoreCase) Then
                    e.Cancel = True
                    ReportCaptured(url)
                End If
            Catch
            End Try
        End Sub

        Private Sub OnNewWindowRequested(sender As Object, e As CoreWebView2NewWindowRequestedEventArgs)
            Try
                ' Force any popup (e.g. SSO) to navigate inside the same WebView2 so we can still capture the redirect.
                e.NewWindow = _webView.CoreWebView2
                e.Handled = True
            Catch
            End Try
        End Sub

        Private Sub ReportCaptured(url As String)
            If _capturedReported Then Return
            _capturedReported = True

            CapturedRedirectUrl = url

            Try
                Dim uri As New Uri(url)
                Dim qs = System.Web.HttpUtility.ParseQueryString(uri.Query)
                CapturedError = If(qs("error"), "")
            Catch
            End Try

            Me.BeginInvoke(New Action(
                Sub()
                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                End Sub))
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            ' If the user closes the window without completing auth, return Cancel.
            If Not _capturedReported AndAlso Me.DialogResult = DialogResult.None Then
                Me.DialogResult = DialogResult.Cancel
            End If
            MyBase.OnFormClosing(e)
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            Try
                If disposing AndAlso _webView IsNot Nothing Then
                    _webView.Dispose()
                    _webView = Nothing
                End If
            Catch
            End Try
            MyBase.Dispose(disposing)
        End Sub

    End Class

End Namespace
