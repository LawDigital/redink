' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: LogWindow.vb
' Purpose: Provides a log window for displaying tooling operations in real-time.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary

''' <summary>
''' A modeless log window that displays tooling operations and allows cancellation.
''' Supports color-coded entries by severity.
''' </summary>
Public Class LogWindow
    Inherits Form

    Private ReadOnly rtbLog As RichTextBox
    Private ReadOnly btnCancel As Button
    Private ReadOnly btnCopy As Button
    Private ReadOnly btnClear As Button

    Private _initialPositionSet As Boolean = False

    Private _autoCloseTimer As Timer = Nothing
    Private _autoCloseSecondsRemaining As Integer = 0
    Private _closeButtonBaseText As String = "Close"

    ''' <summary>
    ''' Raised when the user requests cancellation.
    ''' </summary>
    Public Event CancelRequested As EventHandler

    ''' <summary>
    ''' Initializes a new instance of the LogWindow.
    ''' </summary>
    Public Sub New()
        Me.Text = $"{ThisAddIn.AN} Tooling Log"
        Me.Width = 600
        Me.Height = 400
        Me.StartPosition = FormStartPosition.Manual
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.MinimumSize = New Size(450, 300)
        Me.TopMost = False
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.AutoScaleMode = AutoScaleMode.Font

        Try
            Me.Icon = Icon.FromHandle((New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))).GetHicon())
        Catch
        End Try

        ' Main layout
        Dim mainPanel As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 2,
            .Padding = New Padding(10)
        }
        mainPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        mainPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        ' Log box (RichTextBox for per-entry coloring)
        rtbLog = New RichTextBox() With {
            .Dock = DockStyle.Fill,
            .ReadOnly = True,
            .BorderStyle = BorderStyle.None,
            .ScrollBars = RichTextBoxScrollBars.Vertical,
            .WordWrap = False,
            .DetectUrls = False,
            .BackColor = Me.BackColor,
            .HideSelection = False,
            .ShortcutsEnabled = True,
            .Font = New Font("Consolas", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
        }
        mainPanel.Controls.Add(rtbLog, 0, 0)

        ' Button panel
        Dim buttonPanel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.RightToLeft,
            .AutoSize = True
        }

        btnCancel = New Button() With {
            .Text = "Cancel",
            .AutoSize = True,
            .Padding = New Padding(10, 5, 10, 5)
        }
        AddHandler btnCancel.Click, AddressOf OnCancelClick

        btnCopy = New Button() With {
            .Text = "Copy Log",
            .AutoSize = True,
            .Padding = New Padding(10, 5, 10, 5)
        }
        AddHandler btnCopy.Click, AddressOf OnCopyClick

        btnClear = New Button() With {
            .Text = "Clear",
            .AutoSize = True,
            .Padding = New Padding(10, 5, 10, 5)
        }
        AddHandler btnClear.Click, AddressOf OnClearClick

        buttonPanel.Controls.Add(btnCancel)
        buttonPanel.Controls.Add(btnCopy)
        buttonPanel.Controls.Add(btnClear)

        mainPanel.Controls.Add(buttonPanel, 0, 1)

        Me.Controls.Add(mainPanel)
    End Sub

    Protected Overrides Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)

        If Not _initialPositionSet Then
            PositionWindowInitial()
            _initialPositionSet = True
        End If

        Me.BringToFront()
    End Sub

    Private Sub PositionWindowInitial()
        Const marginRight As Integer = 40
        Const marginBottom As Integer = 40

        Dim wa = Screen.PrimaryScreen.WorkingArea
        Me.Location = New Point(
            wa.Right - Me.Width - marginRight,
            wa.Bottom - Me.Height - marginBottom
        )
    End Sub

    Public Sub StartAutoCloseCountdown(Optional seconds As Integer = 30)
        If seconds <= 0 Then seconds = 1

        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub() StartAutoCloseCountdown(seconds))
            Return
        End If

        _autoCloseSecondsRemaining = seconds

        If _autoCloseTimer Is Nothing Then
            _autoCloseTimer = New Timer() With {.Interval = 1000}
            AddHandler _autoCloseTimer.Tick, AddressOf AutoCloseTimerTick
        End If

        btnCancel.Text = $"{_closeButtonBaseText} ({_autoCloseSecondsRemaining}s)"
        _autoCloseTimer.Stop()
        _autoCloseTimer.Start()
    End Sub

    Private Sub AutoCloseTimerTick(sender As Object, e As EventArgs)
        _autoCloseSecondsRemaining -= 1

        If _autoCloseSecondsRemaining <= 0 Then
            _autoCloseTimer.Stop()
            Try
                Me.Close()
            Catch
            End Try
            Return
        End If

        btnCancel.Text = $"{_closeButtonBaseText} ({_autoCloseSecondsRemaining}s)"
    End Sub

    Private Sub StopAutoCloseCountdown()
        If _autoCloseTimer IsNot Nothing Then
            _autoCloseTimer.Stop()
        End If
        _autoCloseSecondsRemaining = 0
        If btnCancel IsNot Nothing Then
            btnCancel.Text = _closeButtonBaseText
        End If
    End Sub

    ''' <summary>
    ''' Appends a log entry (default level: info).
    ''' </summary>
    Public Sub AppendLog(message As String)
        AppendLog(message, "info")
    End Sub

    ''' <summary>
    ''' Appends a log entry with a severity level (info, warn, error, success, step, llm).
    ''' </summary>
    Public Sub AppendLog(message As String, Optional level As String = "info")
        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub() AppendLogInternal(message, level))
        Else
            AppendLogInternal(message, level)
        End If
    End Sub

    Private Sub AppendLogInternal(message As String, level As String)
        If String.IsNullOrEmpty(message) Then Return

        Dim timestamp = DateTime.Now.ToString("HH:mm:ss.fff")
        Dim prefix = $"[{timestamp}] "

        Dim textColor As Color
        Select Case (If(level, "info")).ToLowerInvariant()
            Case "error", "err", "fail"
                textColor = Color.DarkRed
            Case "warn", "warning"
                textColor = Color.DarkOrange
            Case "success", "ok"
                textColor = Color.DarkGreen
            Case "step"
                textColor = Color.DarkBlue
            Case "llm"
                textColor = Color.DarkMagenta
            Case Else
                textColor = Color.Black
        End Select

        rtbLog.SuspendLayout()
        Try
            rtbLog.SelectionStart = rtbLog.TextLength
            rtbLog.SelectionLength = 0
            rtbLog.SelectionColor = Color.Gray
            rtbLog.AppendText(prefix)

            rtbLog.SelectionStart = rtbLog.TextLength
            rtbLog.SelectionLength = 0
            rtbLog.SelectionColor = textColor
            rtbLog.AppendText(message & Environment.NewLine)

            rtbLog.SelectionStart = rtbLog.TextLength
            rtbLog.ScrollToCaret()
        Finally
            rtbLog.ResumeLayout()
        End Try
    End Sub

    Private Sub OnCancelClick(sender As Object, e As EventArgs)
        btnCancel.Enabled = False
        btnCancel.Text = "Cancelling..."
        RaiseEvent CancelRequested(Me, EventArgs.Empty)
    End Sub

    Private Sub OnCopyClick(sender As Object, e As EventArgs)
        If Not String.IsNullOrEmpty(rtbLog.Text) Then
            Clipboard.SetText(rtbLog.Text)
        End If
    End Sub

    Private Sub OnClearClick(sender As Object, e As EventArgs)
        rtbLog.Clear()
    End Sub

    ''' <summary>
    ''' Updates the cancel button to indicate session is complete.
    ''' </summary>
    Public Sub MarkComplete()
        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub() MarkCompleteInternal())
        Else
            MarkCompleteInternal()
        End If
    End Sub

    Private Sub MarkCompleteInternal()
        StopAutoCloseCountdown()

        btnCancel.Text = "Close"
        _closeButtonBaseText = "Close"
        btnCancel.Enabled = True
        RemoveHandler btnCancel.Click, AddressOf OnCancelClick
        AddHandler btnCancel.Click, Sub()
                                        StopAutoCloseCountdown()
                                        Me.Close()
                                    End Sub

        Me.Text = $"{ThisAddIn.AN} Tooling Log (Complete)"

        StartAutoCloseCountdown(ThisAddIn.ToolingLog_AutoCloseDefaultSeconds)
    End Sub
End Class