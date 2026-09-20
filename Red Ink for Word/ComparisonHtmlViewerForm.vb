' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland.
' All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ComparisonHtmlViewerForm.vb
' Purpose:
'   Hosts Word comparison HTML in WebView2 so tracked deletions, superscripts,
'   subscripts, and other complex revision markup render with the modern Edge
'   engine instead of the legacy WinForms WebBrowser control.
'
' Architecture / Function:
'   - Accepts already-prepared HTML plus a temp host folder.
'   - Writes a dedicated viewer HTML file into that folder and navigates WebView2
'     to the local file path so relative resources continue to resolve.
'   - Recreates the existing compare-window button model: OK plus optional
'     additional action buttons with per-button close behavior.
'   - Invokes the supplied onClose callback exactly once when the form closes.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports SharedLibrary.SharedLibrary.SharedMethods

Public Class ComparisonHtmlViewerForm
    Inherits Form

    Private WithEvents _webView As WebView2

    Private ReadOnly _htmlContent As String
    Private ReadOnly _viewerHostFolder As String
    Private ReadOnly _additionalButtons As System.Tuple(Of System.String, System.Action, System.Boolean)()
    Private ReadOnly _onClose As System.Action

    Private _onCloseInvoked As Boolean
    Private _viewerFilePath As String = Nothing

    Public Sub New(
        htmlContent As String,
        header As String,
        viewerHostFolder As String,
        Optional additionalButtons As System.Tuple(Of System.String, System.Action, System.Boolean)() = Nothing,
        Optional onClose As System.Action = Nothing)

        _htmlContent = If(htmlContent, "")
        _viewerHostFolder = viewerHostFolder
        _additionalButtons = additionalButtons
        _onClose = onClose

        InitializeComponent(header)
        InitializeWebViewAsync()
    End Sub

    Protected Overrides Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)

        Try
            Me.TopMost = True
            ForceDialogToForeground(Me)
            AttachForeignForegroundWatchdog(Me)
            Me.Activate()
            Me.BringToFront()
        Catch
        End Try
    End Sub

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        MyBase.OnFormClosed(e)

        If _onCloseInvoked Then Return
        _onCloseInvoked = True

        If _onClose IsNot Nothing Then
            Try
                _onClose.Invoke()
            Catch
            End Try
        End If
    End Sub

    Private Sub InitializeComponent(header As String)
        Me.Text = If(String.IsNullOrWhiteSpace(header), AN, header)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.MaximizeBox = True
        Me.MinimizeBox = True
        Me.ShowInTaskbar = True
        Me.TopMost = False
        Me.KeyPreview = True
        Me.AutoScaleMode = AutoScaleMode.Font
        Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
        Me.MinimumSize = New Size(800, 500)
        Me.Size = New Size(1100, 760)

        Try
            Dim bmp As New Bitmap(GetLogoBitmap(LogoType.Standard))
            Me.Icon = Icon.FromHandle(bmp.GetHicon())
        Catch
        End Try

        AddHandler Me.KeyDown,
            Sub(sender, e)
                If e.KeyCode = Keys.Escape Then
                    Me.Close()
                    e.SuppressKeyPress = True
                End If
            End Sub

        Dim outer As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 2
        }
        outer.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        outer.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        outer.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        Dim browserHost As New Panel() With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(20, 0, 20, 0),
            .Margin = New Padding(0)
        }

        _webView = New WebView2() With {
            .Dock = DockStyle.Fill,
            .Margin = New Padding(0)
        }
        browserHost.Controls.Add(_webView)

        Dim okButton As New Button() With {
            .Text = "OK",
            .AutoSize = True,
            .Font = Me.Font,
            .Margin = New Padding(0)
        }
        AddHandler okButton.Click,
            Sub()
                Me.Close()
            End Sub

        Dim bottomFlow As New FlowLayoutPanel() With {
            .FlowDirection = FlowDirection.LeftToRight,
            .Dock = DockStyle.Fill,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Padding = New Padding(20),
            .WrapContents = False
        }
        bottomFlow.Controls.Add(okButton)

        If _additionalButtons IsNot Nothing AndAlso _additionalButtons.Length > 0 Then
            For Each btnDef In _additionalButtons
                If String.IsNullOrWhiteSpace(btnDef.Item1) OrElse btnDef.Item2 Is Nothing Then
                    Continue For
                End If

                Dim addBtn As New Button() With {
                    .Text = btnDef.Item1,
                    .AutoSize = True,
                    .Font = Me.Font,
                    .Margin = New Padding(10, okButton.Margin.Top, 0, okButton.Margin.Bottom)
                }

                Dim closeAfter As Boolean = btnDef.Item3
                Dim action As System.Action = btnDef.Item2

                AddHandler addBtn.Click,
                    Sub()
                        Try
                            action.Invoke()
                        Catch
                        End Try

                        If closeAfter Then
                            Me.Close()
                        End If
                    End Sub

                bottomFlow.Controls.Add(addBtn)
            Next
        End If

        bottomFlow.PerformLayout()
        Dim totalButtonWidth As Integer = 0
        For Each ctrl As Control In bottomFlow.Controls
            totalButtonWidth += ctrl.PreferredSize.Width + ctrl.Margin.Left + ctrl.Margin.Right
        Next
        totalButtonWidth += bottomFlow.Padding.Left + bottomFlow.Padding.Right + 40

        Dim minFormWidth As Integer = Math.Max(800, totalButtonWidth)
        Me.MinimumSize = New Size(minFormWidth, 500)
        Me.Size = New Size(Math.Max(minFormWidth, 1100), 760)

        outer.Controls.Add(browserHost, 0, 0)
        outer.Controls.Add(bottomFlow, 0, 1)

        Me.Controls.Add(outer)
    End Sub

    Private Async Sub InitializeWebViewAsync()
        Try
            If String.IsNullOrWhiteSpace(_viewerHostFolder) Then
                Throw New InvalidOperationException("Viewer host folder is missing.")
            End If

            Directory.CreateDirectory(_viewerHostFolder)

            _viewerFilePath = Path.Combine(_viewerHostFolder, "comparison.viewer.htm")
            File.WriteAllText(_viewerFilePath, _htmlContent, System.Text.Encoding.UTF8)

            Dim userDataFolder As String = GetWebView2UserDataFolder()
            Dim env As CoreWebView2Environment =
                Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder)

            Await _webView.EnsureCoreWebView2Async(env)

            AddHandler _webView.CoreWebView2.ProcessFailed,
                Sub(s, e)
                    LogWebView2ProcessFailed("ComparisonViewer", e.ProcessFailedKind.ToString(), e.ExitCode.ToString())
                    Try
                        If e.ProcessFailedKind = CoreWebView2ProcessFailedKind.RenderProcessExited Then
                            _webView.Reload()
                        End If
                    Catch
                    End Try
                End Sub

            _webView.CoreWebView2.Settings.AreDevToolsEnabled = False
            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = True
            _webView.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = True
            _webView.CoreWebView2.Settings.IsStatusBarEnabled = True
            _webView.CoreWebView2.Settings.IsZoomControlEnabled = True

            Dim viewerUri As New Uri(_viewerFilePath)
            _webView.CoreWebView2.Navigate(viewerUri.AbsoluteUri)

        Catch ex As Exception
            ShowCustomMessageBox($"Could not initialize the comparison viewer: {ex.Message}", AN)
            Me.Close()
        End Try
    End Sub
End Class
